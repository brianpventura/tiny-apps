"""Conexao e escrita no MySQL do warehouse (tiny_dw).

Driver: PyMySQL (Python puro, nada para compilar no 3.14) + cryptography, que
NAO e opcional: o plugin default do MySQL 8 e caching_sha2_password e sem ela o
handshake RSA sobre conexao nao-TLS falha com erro obscuro.

Credenciais em data/db.json (mesmo tratamento de arquivo-com-segredo que o
data/tenants.json recebe: fora do codigo, nunca impresso no banner). As variaveis
de ambiente TINY_DB_HOST/PORTA/USUARIO/SENHA/BANCO sobrescrevem campo a campo, e
a CLI sobrescreve as duas.

Tres garantias que este modulo da a quem carrega:

  1. transacao(conn) -- commit no fim, rollback na excecao. Nunca ha "dado
     gravado sem o estado marcado".
  2. com_reconexao(fn) -- separa erro de INFRA (retentavel: refaz o lote inteiro,
     seguro porque todo write e ON DUPLICATE KEY UPDATE com chave natural) de
     erro de DADO (nao retentavel: o chamador refaz linha a linha para isolar a
     linha ruim, que vai para etl_erro, e o resto do lote entra).
  3. travar_tenant(conn, tenant) -- GET_LOCK que impede dois processos do mesmo
     tenant. Dois processos = dois rate limiters independentes = 4 req/s na mesma
     conta = 429 em cadeia. O lock morre com a conexao, entao crash nao deixa
     cadeado preso.
"""

import json
import os
import time

import _bootstrap
import pymysql
from pymysql.err import OperationalError, InterfaceError

# Erros de INFRA: vale reconectar e refazer o lote inteiro.
_ERROS_INFRA = {
    2006,  # MySQL server has gone away
    2013,  # Lost connection during query
    2003,  # nao conseguiu conectar
    1205,  # Lock wait timeout exceeded
    1213,  # Deadlock found
}

# Erros de DADO: refazer o lote nao adianta, a linha e que esta ruim.
ERROS_DADO = {
    1062,  # Duplicate entry
    1264,  # Out of range value
    1292,  # Incorrect datetime value
    1366,  # Incorrect string value
    1406,  # Data too long for column
    1452,  # Cannot add or update a child row (FK)
}

_PADRAO = {"host": "127.0.0.1", "porta": 3306, "usuario": "root",
           "senha": "", "banco": "tiny_dw"}


def carregar_config(overrides: dict | None = None) -> dict:
    """data/db.json -> variaveis de ambiente -> overrides da CLI."""
    cfg = dict(_PADRAO)
    if _bootstrap.DB_FILE.exists():
        try:
            cfg.update(json.loads(_bootstrap.DB_FILE.read_text(encoding="utf-8")))
        except Exception as e:
            raise SystemExit(f"  [ERRO] {_bootstrap.DB_FILE} ilegivel: {e}")
    for chave, env in (("host", "TINY_DB_HOST"), ("porta", "TINY_DB_PORTA"),
                       ("usuario", "TINY_DB_USUARIO"), ("senha", "TINY_DB_SENHA"),
                       ("banco", "TINY_DB_BANCO")):
        v = os.environ.get(env)
        if v:
            cfg[chave] = v
    for chave, v in (overrides or {}).items():
        if v is not None:
            cfg[chave] = v
    cfg["porta"] = int(cfg["porta"])
    return cfg


def descrever(cfg: dict) -> str:
    """Identificacao da conexao para o banner. NUNCA inclui a senha."""
    return "{}@{}:{}/{}".format(cfg["usuario"], cfg["host"], cfg["porta"],
                                cfg["banco"])


def conectar(cfg: dict | None = None, com_banco: bool = True):
    """Abre conexao. com_banco=False conecta sem selecionar o schema (usado so
    para o CREATE DATABASE da fase 0)."""
    cfg = cfg or carregar_config()
    try:
        return pymysql.connect(
            host=cfg["host"], port=cfg["porta"],
            user=cfg["usuario"], password=cfg["senha"],
            database=cfg["banco"] if com_banco else None,
            charset="utf8mb4",
            autocommit=False,
            cursorclass=pymysql.cursors.DictCursor,
            # DATETIME em vez de TIMESTAMP em todo o schema, entao o offset fixo
            # basta. As tabelas de timezone nomeadas nao vem carregadas no MySQL
            # do Windows: SET time_zone='America/Sao_Paulo' daria ERROR 1298.
            init_command="SET SESSION time_zone='-03:00'",
            connect_timeout=15,
            read_timeout=180,
            write_timeout=180,
        )
    except OperationalError as e:
        codigo = e.args[0] if e.args else "?"
        detalhe = e.args[1] if len(e.args) > 1 else e
        dica = ""
        if not _bootstrap.DB_FILE.exists():
            dica = ("\n         " + str(_bootstrap.DB_FILE) + " nao existe. "
                    "Copie data/db.exemplo.json e preencha.")
        raise SystemExit(
            "  [ERRO] nao conectou em {} (MySQL {}).\n         {}\n"
            "         Confira o servico MySQL80.{}".format(
                descrever(cfg), codigo, detalhe, dica))


class transacao:
    """Commit no fim do bloco, rollback em qualquer excecao."""

    def __init__(self, conn):
        self.conn = conn

    def __enter__(self):
        return self.conn

    def __exit__(self, exc_tipo, exc, tb):
        if exc_tipo is None:
            self.conn.commit()
        else:
            try:
                self.conn.rollback()
            except Exception:
                pass
        return False


def codigo_erro(e: Exception):
    return e.args[0] if getattr(e, "args", None) else None


def eh_infra(e: Exception) -> bool:
    if isinstance(e, InterfaceError):
        return True
    if isinstance(e, OperationalError):
        return codigo_erro(e) in _ERROS_INFRA
    return False


def executar_lote(conn, sql: str, linhas, tamanho: int = 500) -> int:
    """executemany em blocos. tamanho limitado porque o driver monta um unico
    INSERT gigante e max_allowed_packet e 64M."""
    linhas = list(linhas)
    if not linhas:
        return 0
    n = 0
    with conn.cursor() as cur:
        for i in range(0, len(linhas), tamanho):
            bloco = linhas[i:i + tamanho]
            cur.executemany(sql, bloco)
            n += len(bloco)
    return n


def com_reconexao(ref: dict, fn, tentativas: int = 4):
    """Roda fn(conn) e, em erro de INFRA, reconecta e REFAZ.

    Refazer e seguro porque todo write do pipeline e idempotente
    (INSERT ... ON DUPLICATE KEY UPDATE com chave natural). ref e um dict
    {"conn": conn, "cfg": cfg} para que a reconexao seja visivel ao chamador.
    """
    ultimo = None
    for tentativa in range(tentativas):
        try:
            return fn(ref["conn"])
        except Exception as e:
            ultimo = e
            if not eh_infra(e) or tentativa == tentativas - 1:
                raise
            espera = (2, 5, 10)[min(tentativa, 2)]
            print("  [RETRY {}/{}] MySQL caiu ({}); reconectando em {}s".format(
                tentativa + 1, tentativas - 1, e, espera), flush=True)
            try:
                ref["conn"].close()
            except Exception:
                pass
            time.sleep(espera)
            ref["conn"] = conectar(ref["cfg"])
    raise ultimo


def ping(conn) -> None:
    """Custa ~0,3 ms e imuniza contra reinicio do servico entre dois lotes."""
    try:
        conn.ping(reconnect=True)
    except Exception:
        pass


def travar_tenant(conn, tenant: str) -> bool:
    """GET_LOCK nao-bloqueante. False = outro processo ja roda este tenant."""
    with conn.cursor() as cur:
        cur.execute("SELECT GET_LOCK(%s, 0) AS ok", ("etl_pedidos_" + tenant,))
        linha = cur.fetchone()
    return bool(linha and linha.get("ok") == 1)


def _executar(cur, sql: str, params):
    """PyMySQL so interpola quando `args` nao e None -- e `()` NAO e None.

    Passar tupla vazia faria o driver rodar `sql % ()`, e a DDL deste schema
    tem `STR_TO_DATE(..., '%Y-%m-%d')`: o %Y viraria placeholder e estouraria
    com "not enough arguments for format string". SQL sem parametro tem de ir
    sem o segundo argumento.
    """
    if params:
        return cur.execute(sql, params)
    return cur.execute(sql)


def um_valor(conn, sql: str, params=None):
    with conn.cursor() as cur:
        _executar(cur, sql, params)
        linha = cur.fetchone()
    if not linha:
        return None
    return next(iter(linha.values()))


def consultar(conn, sql: str, params=None) -> list:
    with conn.cursor() as cur:
        _executar(cur, sql, params)
        return list(cur.fetchall())


def executar(conn, sql: str, params=None) -> int:
    with conn.cursor() as cur:
        return _executar(cur, sql, params)
