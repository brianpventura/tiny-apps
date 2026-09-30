"""Fase 0 -- cria o banco, aplica sql/*.sql e semeia as tabelas cfg_*/dim_*.

Nao ha ferramenta de migracao no repo, entao o esquema tem dono aqui:
os arquivos de `sql/` sao aplicados em ordem de nome e registrados em
`schema_migracao` com o sha256 do conteudo -- assim um arquivo editado depois de
aplicado e detectado em vez de passar em silencio.

Toda DDL e idempotente (`CREATE TABLE IF NOT EXISTS`, `CREATE OR REPLACE VIEW`)
porque no MySQL a DDL causa COMMIT IMPLICITO: nao existe migracao atomica, entao
a defesa e poder reaplicar sempre.

O seed fica em Python, e nao num `.sql`, porque depende de `data/tenants.json` e
das constantes de regra de `core/regras_vendas.py` -- gravar aquilo num arquivo
estatico criaria a segunda copia das regras que o repo inteiro evita.
"""


import hashlib
from datetime import date, datetime, timedelta

import _bootstrap
import etl_db
import etl_regras as regras

_DIAS_SEM = ["seg", "ter", "qua", "qui", "sex", "sab", "dom"]

CAL_INICIO = date(2023, 1, 1)
CAL_FIM = date(2029, 12, 31)


# ------------------------------------------------------------------ #
#  Separador de comandos SQL                                          #
# ------------------------------------------------------------------ #

def dividir_comandos(sql: str) -> list:
    """Divide um script em comandos, respeitando strings e comentarios.

    Um split ingenuo por ';' nao serve: varios COMMENT='...' deste schema
    contem '--' dentro da string (ex.: 'do cadastro -- unica fonte de kit'), e
    uma limpeza ingenua de comentarios os truncaria.
    """
    comandos, atual = [], []
    i, n = 0, len(sql)
    em_string = False
    while i < n:
        c = sql[i]
        if em_string:
            atual.append(c)
            if c == "\\" and i + 1 < n:
                atual.append(sql[i + 1])
                i += 2
                continue
            if c == "'":
                if i + 1 < n and sql[i + 1] == "'":   # aspa escapada ''
                    atual.append(sql[i + 1])
                    i += 2
                    continue
                em_string = False
            i += 1
            continue
        if c == "'":
            em_string = True
            atual.append(c)
            i += 1
            continue
        if c == "-" and sql.startswith("--", i):
            fim = sql.find("\n", i)
            i = n if fim == -1 else fim
            continue
        if c == "/" and sql.startswith("/*", i):
            fim = sql.find("*/", i + 2)
            i = n if fim == -1 else fim + 2
            continue
        if c == ";":
            comandos.append("".join(atual).strip())
            atual = []
            i += 1
            continue
        atual.append(c)
        i += 1
    resto = "".join(atual).strip()
    if resto:
        comandos.append(resto)
    return [c for c in comandos if c]


# ------------------------------------------------------------------ #
#  Criacao do banco e aplicacao dos arquivos                          #
# ------------------------------------------------------------------ #

def criar_banco(cfg: dict) -> bool:
    """CREATE DATABASE IF NOT EXISTS. Devolve True se criou agora."""
    conn = etl_db.conectar(cfg, com_banco=False)
    try:
        existia = etl_db.um_valor(
            conn,
            "SELECT COUNT(*) FROM information_schema.schemata "
            "WHERE schema_name = %s", (cfg["banco"],))
        if not existia:
            etl_db.executar(
                conn,
                "CREATE DATABASE `{}` CHARACTER SET utf8mb4 "
                "COLLATE utf8mb4_0900_ai_ci".format(cfg["banco"]))
            conn.commit()
        return not existia
    finally:
        conn.close()


def aplicar(conn, forcar: bool = False) -> int:
    """Aplica sql/*.sql em ordem de nome. Devolve quantos arquivos rodaram."""
    arquivos = sorted(_bootstrap.SQL_DIR.glob("*.sql"))
    if not arquivos:
        raise SystemExit("  [ERRO] nenhum .sql em " + str(_bootstrap.SQL_DIR))

    # A propria tabela de controle precisa existir antes de ser consultada.
    primeiro = arquivos[0].read_text(encoding="utf-8")
    for cmd in dividir_comandos(primeiro):
        if "schema_migracao" in cmd and cmd.upper().startswith("CREATE TABLE"):
            etl_db.executar(conn, cmd)
            conn.commit()
            break

    aplicados = {r["arquivo"]: r["sha256"] for r in etl_db.consultar(
        conn, "SELECT arquivo, sha256 FROM schema_migracao")}

    rodados = 0
    for arq in arquivos:
        texto = arq.read_text(encoding="utf-8")
        digest = hashlib.sha256(texto.encode("utf-8")).hexdigest()
        anterior = aplicados.get(arq.name)
        if anterior == digest and not forcar:
            print("  [esquema] {} ja aplicado.".format(arq.name))
            continue
        if anterior and anterior != digest:
            print("  [AVISO] {} mudou depois de aplicado; reaplicando "
                  "(a DDL e idempotente).".format(arq.name))
        comandos = dividir_comandos(texto)
        for cmd in comandos:
            try:
                etl_db.executar(conn, cmd)
            except Exception as e:
                cabeca = " ".join(cmd.split())[:120]
                raise SystemExit(
                    "  [ERRO] {} falhou em: {}\n         {}".format(
                        arq.name, cabeca, e))
        # DDL ja commitou sozinha; o commit aqui fecha o INSERT de controle.
        etl_db.executar(
            conn,
            "INSERT INTO schema_migracao (arquivo, sha256, aplicado_em) "
            "VALUES (%s, %s, %s) AS novo "
            "ON DUPLICATE KEY UPDATE sha256 = novo.sha256, "
            "aplicado_em = novo.aplicado_em",
            (arq.name, digest, datetime.now()))
        conn.commit()
        print("  [esquema] {}: {} comandos.".format(arq.name, len(comandos)))
        rodados += 1
    return rodados


# ------------------------------------------------------------------ #
#  Seed                                                               #
# ------------------------------------------------------------------ #

def _seed_calendario(conn) -> int:
    ja = etl_db.um_valor(conn, "SELECT COUNT(*) FROM dim_calendario")
    esperado = (CAL_FIM - CAL_INICIO).days + 1
    if ja and ja >= esperado:
        return 0
    linhas, d = [], CAL_INICIO
    while d <= CAL_FIM:
        dow = d.weekday()
        linhas.append((d, d.year, d.month, d.day, (d.month - 1) // 3 + 1,
                       "{:04d}-{:02d}".format(d.year, d.month), dow,
                       _DIAS_SEM[dow], 1 if dow >= 5 else 0))
        d += timedelta(days=1)
    sql = ("INSERT INTO dim_calendario "
           "(dia, ano, mes, dia_mes, trimestre, ano_mes, dia_semana, nome_dia, "
           " eh_fim_semana) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) AS novo "
           "ON DUPLICATE KEY UPDATE ano_mes = novo.ano_mes")
    return etl_db.executar_lote(conn, sql, linhas, tamanho=1000)


def _seed_situacoes(conn) -> int:
    # Vocabulario do Tiny + validade no perfil da EMPRESA (sem override por
    # tenant: o que vale por pedido e a coluna pedido.eh_situacao_valida).
    p = regras.perfil()
    linhas = [(cod, nome, 1 if p.eh_situacao_valida(cod) else 0)
              for cod, nome in sorted(regras.SITUACOES.items())]
    sql = ("INSERT INTO cfg_situacao (codigo, nome, eh_valida) "
           "VALUES (%s,%s,%s) AS novo "
           "ON DUPLICATE KEY UPDATE nome = novo.nome, "
           "eh_valida = novo.eh_valida")
    return etl_db.executar_lote(conn, sql, linhas)


def _seed_tenants(conn, tenants: list) -> int:
    import json
    cfg = {}
    try:
        cfg = json.loads(_bootstrap.TENANTS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    linhas = []
    for t in tenants:
        conf = cfg.get(t, {})
        linhas.append((t, regras.NOMES.get(t, t), conf.get("razao_social") or ""))
    sql = ("INSERT INTO cfg_tenant (tenant, empresa, razao_social) "
           "VALUES (%s,%s,%s) AS novo "
           "ON DUPLICATE KEY UPDATE empresa = novo.empresa, "
           "razao_social = novo.razao_social")
    return etl_db.executar_lote(conn, sql, linhas)


def _seed_clientes_excluidos(conn, tenants: list) -> int:
    linhas = []
    for t in tenants:
        # clientes_excluidos() ja soma a razao_social do PROPRIO tenant e
        # normaliza tudo -- sem isso a auto-venda entra como faturamento.
        for nome in sorted(regras.clientes_excluidos(t)):
            linhas.append((t, nome[:160]))
    etl_db.executar_lote(
        conn, "DELETE FROM cfg_cliente_excluido WHERE tenant = %s",
        [(t,) for t in tenants])
    sql = "INSERT INTO cfg_cliente_excluido (tenant, nome_norm) VALUES (%s,%s)"
    return etl_db.executar_lote(conn, sql, linhas)


def _seed_sentinelas(conn, tenants: list) -> int:
    """Linha id=0 em cada dimensao, para o LEFT JOIN nunca perder pedido."""
    agora = datetime.now()
    n = 0
    tabelas = [
        ("dim_vendedor", "id_vendedor"),
        ("dim_ecommerce", "id_ecommerce"),
        ("dim_natureza_operacao", "id_natureza"),
        ("dim_transportador", "id_transportador"),
        ("dim_deposito", "id_deposito"),
        ("dim_intermediador", "id_intermediador"),
    ]
    for tabela, coluna in tabelas:
        sql = ("INSERT INTO {} (tenant, {}, nome, ultimo_visto) "
               "VALUES (%s, 0, %s, %s) AS novo "
               "ON DUPLICATE KEY UPDATE nome = novo.nome".format(tabela, coluna))
        n += etl_db.executar_lote(
            conn, sql, [(t, "(nao informado)", agora) for t in tenants])
    sql_cli = ("INSERT INTO dim_cliente (tenant, id_cliente, nome, nome_norm, "
               "primeiro_visto, ultimo_visto) VALUES (%s, 0, %s, %s, %s, %s) "
               "AS novo ON DUPLICATE KEY UPDATE nome = novo.nome")
    n += etl_db.executar_lote(conn, sql_cli, [
        (t, "(nao informado)", "(nao informado)", agora, agora) for t in tenants])
    sql_prod = ("INSERT INTO dim_produto (tenant, id_produto, sku, descricao, "
                "atualizado_em) VALUES (%s, 0, %s, %s, %s) AS novo "
                "ON DUPLICATE KEY UPDATE descricao = novo.descricao")
    n += etl_db.executar_lote(conn, sql_prod, [
        (t, "", "(nao informado)", agora) for t in tenants])
    return n


def semear(conn, tenants: list) -> None:
    """Popula cfg_* e as sentinelas. Idempotente, roda a cada fase 0."""
    with etl_db.transacao(conn):
        cal = _seed_calendario(conn)
        sit = _seed_situacoes(conn)
        ten = _seed_tenants(conn, tenants)
        exc = _seed_clientes_excluidos(conn, tenants)
        sen = _seed_sentinelas(conn, tenants)
    print("  [esquema] seed: {} dias de calendario, {} situacoes, "
          "{} tenants, {} clientes excluidos, {} sentinelas.".format(
              cal, sit, ten, exc, sen))


def fase0(cfg: dict, tenants: list, forcar: bool = False):
    """Cria o banco se preciso, aplica o esquema e semeia. Devolve a conexao."""
    if criar_banco(cfg):
        print("  [esquema] banco {} criado.".format(cfg["banco"]))
    conn = etl_db.conectar(cfg)
    aplicar(conn, forcar=forcar)
    semear(conn, tenants)
    return conn
