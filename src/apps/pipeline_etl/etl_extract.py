"""Extract: fase 1 (listagem por janela) e fase 2 (detalhe pedido a pedido).

O que este modulo NAO reimplementa: retry de 5xx, Retry-After no 429, backoff e
recriacao da Session. Tudo isso ja esta em `core/tiny_client.py:_get`. Duplicar
so criaria uma segunda politica divergente.

O que ele acrescenta, e por que:

  * `blindar_token` -- o `TinyTokenManager` do repo NAO tem lock e o `_save()`
    dele nao e atomico. Com 4 workers, dois refresh simultaneos fazem o segundo
    tomar HTTP 400 (que vira TokenExpiredError ESPURIA) e podem deixar o
    `data/tokens_<tenant>.json` truncado -- e ai o proximo processo nem carrega.
    Resolvido aqui, sem tocar em `core/`, que e compartilhado por ~30 scripts.
  * captura EXPLICITA de `TokenExpiredError` antes do `except Exception`. Ela e
    levantada em `auth_header()` ANTES de tocar a rede; engolida pelo except
    generico, os 35 mil pedidos restantes "falhariam" em segundos sem gastar
    request e o pipeline terminaria com tudo em erro, parecendo problema de
    dado. Ao ve-la, o pipeline para limpo e sai com codigo 2.
  * janela que estoura vira subdivisao. `GET /pedidos?offset=57841` devolve
    HTTP 400 "Sua consulta levou muito tempo"; o teto pratico e ~6k.
"""

import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import _bootstrap  # noqa: F401
import etl_db
import etl_load
import etl_regras as regras
import etl_transform as T
from core.auth_tiny import TokenExpiredError
from core.produtos_utils import desempacotar_data

_MAX_WORKERS = 4
_LOTE_PADRAO = 200       # ~100 s de trabalho a 2 req/s
_INTERVALO_FLUSH = 120   # segundos: comita mesmo com o buffer curto
_TENTATIVAS_LISTA = 6

_TOKEN_LOCK = threading.Lock()


class ParadaToken(Exception):
    """Refresh token morreu no meio da coleta. Estado salvo, retomavel."""


def blindar_token(client) -> None:
    """Serializa auth_header()/_refresh() do TinyTokenManager, que nao tem lock.

    Custa nada: a 2 req/s o lock e adquirido 2x/s para uma leitura de dict.
    """
    tm = getattr(client, "_token_manager", None)
    if tm is None or getattr(tm, "_blindado", False):
        return
    original = tm.auth_header

    def _sob_lock():
        with _TOKEN_LOCK:
            return original()

    tm.auth_header = _sob_lock
    tm._blindado = True


def aquecer_token(client) -> None:
    """Renova o access token ANTES de abrir o pool, para que nenhum worker
    precise disparar refresh no meio do lote."""
    tm = getattr(client, "_token_manager", None)
    if tm is None:
        return
    with _TOKEN_LOCK:
        tm.get_access_token()


# ------------------------------------------------------------------ #
#  Fase 1 -- listagem da janela (SEM filtro nenhum)                   #
# ------------------------------------------------------------------ #

def _total_da_janela(client, ini, fim):
    """1 request com limit=1 so para ler paginacao.total.

    E a valvula de seguranca do backfill: janela grande demais e subdividida
    ANTES de paginar, em vez de estourar o offset no meio da madrugada.
    """
    resp = desempacotar_data(client.listar_pedidos(
        offset=0, limit=1, dataInicial=ini.isoformat(),
        dataFinal=fim.isoformat()))
    return (resp.get("paginacao") or {}).get("total")


def _gravar_janela(conn, tenant, ini, fim, total_api, gravado, status, msg=None):
    with etl_db.transacao(conn):
        etl_db.executar(
            conn,
            "INSERT INTO etl_janela (tenant, ini, fim, total_api, "
            "total_gravado, status, mensagem, atualizado_em) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s) AS novo "
            "ON DUPLICATE KEY UPDATE total_api = novo.total_api, "
            "total_gravado = novo.total_gravado, status = novo.status, "
            "mensagem = novo.mensagem, atualizado_em = novo.atualizado_em",
            (tenant, ini, fim, total_api, gravado, status,
             (msg or None), datetime.now()))


def _paginar(client, ref, tenant, ini, fim) -> tuple:
    """Pagina uma janela inteira, commitando a cada pagina de 100."""
    di, dfim = ini.isoformat(), fim.isoformat()
    offset, total, gravados, fora = 0, None, 0, 0
    while True:
        resp = None
        for tentativa in range(_TENTATIVAS_LISTA):
            try:
                resp = desempacotar_data(client.listar_pedidos(
                    offset=offset, dataInicial=di, dataFinal=dfim))
                break
            except TokenExpiredError:
                raise
            except Exception as e:
                if tentativa < _TENTATIVAS_LISTA - 1:
                    time.sleep(3 * (tentativa + 1))
                else:
                    # Diferenca deliberada do repo: o padrao atual segue como se
                    # a pagina estivesse vazia e PERDE dado em silencio. Aqui a
                    # janela vai para status='erro' e e re-paginada depois.
                    raise RuntimeError(
                        "lista {}..{} off={}: {}".format(di, dfim, offset, e))
        itens = resp.get("itens") or []
        total = (resp.get("paginacao") or {}).get("total", total)
        if not itens:
            break

        registros = []
        for ped in itens:
            pid = T.inteiro(ped.get("id"))
            if pid is None:
                continue
            d = regras.data_de(ped)
            if not d or not (ini <= d <= fim):
                fora += 1
                continue
            registros.append({
                "tenant": tenant, "id_pedido": pid,
                "payload": T.serializar(ped), "hash": T.hash_payload(ped),
                "data_ref": d, "numero": T.inteiro(ped.get("numeroPedido")),
                "situacao": T.inteiro(ped.get("situacao")),
            })
        gravados += etl_db.com_reconexao(
            ref, lambda c, r=registros: etl_load.gravar_lote_lista(c, r))

        offset += len(itens)
        if total is not None and offset >= total:
            break
    return gravados, total, fora


def fase1_listar(client, ref, tenant, inicio, fim, max_janela=5000,
                 refazer=False) -> int:
    """Lista a janela inteira -- TUDO, sem filtro de situacao ou natureza."""
    conn = ref["conn"]
    pendentes = list(regras.janelas_mensais(inicio, fim))
    # Do mais recente para o mais antigo: se o backfill for interrompido, o que
    # esta no banco e o periodo que mais interessa.
    pendentes.sort(key=lambda j: j[0], reverse=True)

    total_gravado, erros = 0, 0
    fila = list(pendentes)
    while fila:
        ini, fim_j = fila.pop(0)
        if not refazer:
            st = etl_db.um_valor(
                conn, "SELECT status FROM etl_janela WHERE tenant=%s AND "
                      "ini=%s AND fim=%s", (tenant, ini, fim_j))
            if st == "ok":
                print("  [{}] janela {}..{} ja ok (cache do banco).".format(
                    tenant, ini, fim_j))
                continue

        try:
            total = _total_da_janela(client, ini, fim_j)
        except TokenExpiredError:
            raise
        except Exception as e:
            print("  [{}] [ERRO] total de {}..{}: {}".format(tenant, ini, fim_j, e))
            _gravar_janela(ref["conn"], tenant, ini, fim_j, None, 0, "erro", str(e))
            erros += 1
            continue

        if total and total > max_janela and ini < fim_j:
            partes = regras.metade(ini, fim_j)
            print("  [{}] janela {}..{} tem {} pedidos (> {}); subdividindo em "
                  "{}.".format(tenant, ini, fim_j, total, max_janela, len(partes)))
            fila = partes + fila
            continue

        try:
            gravados, total_api, fora = _paginar(client, ref, tenant, ini, fim_j)
        except TokenExpiredError:
            raise
        except Exception as e:
            print("  [{}] [ERRO] {}".format(tenant, e))
            _gravar_janela(ref["conn"], tenant, ini, fim_j, total, 0, "erro", str(e))
            erros += 1
            continue

        status = "ok" if (total_api is None or gravados == total_api) else "erro"
        _gravar_janela(ref["conn"], tenant, ini, fim_j, total_api, gravados,
                       status,
                       None if status == "ok" else
                       "gravado {} de {}".format(gravados, total_api))
        total_gravado += gravados
        aviso = "" if status == "ok" else "  [AVISO] nao fecha com a API!"
        extra = "" if not fora else "  (fora do periodo: {})".format(fora)
        print("  [{}] janela {}..{}: {} gravados / {} na API{}{}".format(
            tenant, ini, fim_j, gravados, total_api, extra, aviso), flush=True)

    if erros:
        print("  [{}] [AVISO] {} janela(s) em erro; rode a fase 1 de novo "
              "para re-paginar so elas.".format(tenant, erros))
    return total_gravado


# ------------------------------------------------------------------ #
#  Fase 2 -- detalhe: 1 request por pedido, inevitavel                #
# ------------------------------------------------------------------ #

def pendentes_detalhe(conn, tenant, max_tentativas=5, limite=None) -> list:
    sql = ("SELECT id_pedido FROM etl_pedido_estado "
           "WHERE tenant = %s AND estado = 'listado' AND tentativas < %s "
           "ORDER BY data_ref DESC, id_pedido DESC")
    params = [tenant, max_tentativas]
    if limite:
        sql += " LIMIT %s"
        params.append(int(limite))
    return [r["id_pedido"] for r in etl_db.consultar(conn, sql, tuple(params))]


def fase2_detalhar(client, ref, tenant, lote=_LOTE_PADRAO,
                   workers=_MAX_WORKERS, limite=None, max_tentativas=5) -> tuple:
    """GET /pedidos/{id} de tudo que esta em 'listado'.

    O commit acontece a cada `lote` pedidos (ou a cada _INTERVALO_FLUSH
    segundos), e o estado de cada pedido e marcado na MESMA transacao do dado.
    Depois do commit, nenhum deles sera buscado de novo.
    """
    pendentes = pendentes_detalhe(ref["conn"], tenant, max_tentativas, limite)
    if not pendentes:
        print("  [{}] fase 2: nada pendente.".format(tenant))
        return 0, 0
    if limite:
        print("  [{}] fase 2 LIMITADA a {} pedidos (smoke test).".format(
            tenant, len(pendentes)))
    print("  [{}] fase 2: {} detalhes a buscar. ~{:.1f} min a 2 req/s.".format(
        tenant, len(pendentes), len(pendentes) / 2 / 60))

    blindar_token(client)
    aquecer_token(client)

    parar = threading.Event()
    lock = threading.Lock()
    buffer = []
    contador = {"ok": 0, "erro": 0, "feitos": 0}
    t0 = time.time()
    ultimo_flush = {"t": time.time()}

    def _flush(forcado=False):
        """Chamado SEMPRE com o lock em maos."""
        if not buffer and not forcado:
            return
        pendente, buffer[:] = list(buffer), []
        if not pendente:
            return
        etl_db.ping(ref["conn"])
        ok, falhas = etl_db.com_reconexao(
            ref, lambda c: etl_load.gravar_lote_detalhe(
                c, tenant, pendente, max_tentativas))
        ultimo_flush["t"] = time.time()
        feito = contador["feitos"]
        taxa = feito / max(1e-9, time.time() - t0)
        resta = (len(pendentes) - feito) / max(1e-9, taxa) / 60
        print("  [{}] {}/{} detalhes ({:.1f}/s, ~{:.1f} min restantes, "
              "{} erros)".format(tenant, feito, len(pendentes), taxa, resta,
                                 contador["erro"]), flush=True)
        return ok, falhas

    def _proc(pid):
        if parar.is_set():
            return None
        try:
            bruto = client.obter_pedido(pid)
        except TokenExpiredError:
            # ANTES do except Exception, de proposito. Ver a docstring do modulo.
            parar.set()
            return {"id_pedido": pid, "token": True}
        except Exception as e:
            return {"id_pedido": pid, "erro": str(e)[:500]}
        det = desempacotar_data(bruto)
        return {"id_pedido": pid, "payload": T.serializar(det),
                "hash": T.hash_payload(det),
                "data_ref": regras.data_de(det),
                "situacao": T.inteiro(det.get("situacao"))}

    token_morreu = False
    try:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = [ex.submit(_proc, pid) for pid in pendentes]
            for fut in as_completed(futs):
                r = fut.result()
                if r is None:
                    continue
                if r.get("token"):
                    token_morreu = True
                    continue
                with lock:
                    buffer.append(r)
                    contador["feitos"] += 1
                    if r.get("erro"):
                        contador["erro"] += 1
                    else:
                        contador["ok"] += 1
                    if (len(buffer) >= lote
                            or time.time() - ultimo_flush["t"] >= _INTERVALO_FLUSH):
                        _flush()
                if token_morreu:
                    ex.shutdown(wait=False, cancel_futures=True)
                    break
    finally:
        # Flush tambem no caminho de erro, de Ctrl+C e de token morto: o que ja
        # foi baixado tem de estar no banco antes de sair.
        with lock:
            _flush()

    if token_morreu or parar.is_set():
        raise ParadaToken(tenant)

    print("  [{}] fase 2: {} detalhados, {} com erro.".format(
        tenant, contador["ok"], contador["erro"]))
    return contador["ok"], contador["erro"]
