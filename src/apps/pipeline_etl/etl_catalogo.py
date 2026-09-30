"""Fase 3 -- completa dim_produto com o cadastro (tipo, situacao, categoria).

Por que existe: o `produto.tipo` que vem DENTRO do item do pedido usa outro
vocabulario (visto "P") e NAO serve para detectar kit. Quem responde isso e
`GET /produtos/{id}` com `tipo == "K"`. Categoria tambem so existe no detalhe do
produto (a listagem de /produtos nem devolve, e ?idCategoria= e ignorado).

Custo: 1 request por produto NOVO. So os produtos que apareceram em pedido, nao
o catalogo inteiro -- algumas centenas na primeira carga e algumas dezenas
depois, porque o resultado fica no banco.
"""

import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import _bootstrap  # noqa: F401
import etl_db
import etl_extract
import etl_transform as T
from core.auth_tiny import TokenExpiredError
from core.produtos_utils import desempacotar_data

SQL_UPSERT = (
    "INSERT INTO dim_produto (tenant, id_produto, sku, descricao, "
    "tipo_cadastro, situacao, categoria_caminho, atualizado_em) "
    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s) AS novo "
    "ON DUPLICATE KEY UPDATE sku = novo.sku, descricao = novo.descricao, "
    "tipo_cadastro = novo.tipo_cadastro, situacao = novo.situacao, "
    "categoria_caminho = novo.categoria_caminho, "
    "atualizado_em = novo.atualizado_em")


def _categoria(prod: dict):
    cat = T.obj(prod, "categoria")
    for chave in ("caminhoCompleto", "caminho", "descricao", "nome"):
        v = cat.get(chave)
        if v:
            return T.txt(v, 500)
    v = prod.get("categoria")
    return T.txt(v, 500) if isinstance(v, str) else None


def pendentes(conn, tenant: str, limite: int | None = None) -> list:
    """Produtos que apareceram em pedido e ainda nao tem cadastro lido.

    Le os ids direto do JSON cru com JSON_TABLE, e NAO de dim_produto. A
    diferenca importa: dim_produto so ganha linha na fase 4, que roda DEPOIS
    da 3 na ordem natural das fases -- lendo de la, a fase 3 nao acharia nada
    numa carga limpa e `tipo_cadastro` ficaria eternamente NULL, ou seja, kit
    nenhum seria detectado. Lendo do raw, a fase 3 fica independente da 4.
    """
    sql = ("SELECT DISTINCT jt.id_produto "
           "FROM pedido_raw r "
           "JOIN etl_pedido_estado e "
           "  ON e.tenant = r.tenant AND e.id_pedido = r.id_pedido "
           " AND r.payload_hash = e.hash_detalhe "
           "CROSS JOIN JSON_TABLE(r.payload, '$.itens[*]' COLUMNS ("
           "  id_produto BIGINT PATH '$.produto.id')) jt "
           "LEFT JOIN dim_produto d "
           "  ON d.tenant = r.tenant AND d.id_produto = jt.id_produto "
           "WHERE r.tenant = %s AND r.fonte = 'detalhe' "
           "  AND jt.id_produto IS NOT NULL AND jt.id_produto <> 0 "
           "  AND (d.id_produto IS NULL OR d.tipo_cadastro IS NULL) "
           "ORDER BY jt.id_produto")
    params = [tenant]
    if limite:
        sql += " LIMIT %s"
        params.append(int(limite))
    return [r["id_produto"] for r in etl_db.consultar(conn, sql, tuple(params))]


def fase3_catalogo(client, ref, tenant: str, workers: int = 4,
                   lote: int = 200, limite: int | None = None) -> tuple:
    ids = pendentes(ref["conn"], tenant, limite)
    if not ids:
        print("  [{}] fase 3: catalogo completo, nada a buscar.".format(tenant))
        return 0, 0
    print("  [{}] fase 3: {} produtos novos. ~{:.1f} min a 2 req/s.".format(
        tenant, len(ids), len(ids) / 2 / 60))

    etl_extract.blindar_token(client)
    etl_extract.aquecer_token(client)

    parar = threading.Event()
    lock = threading.Lock()
    buffer, contador = [], {"ok": 0, "erro": 0}
    agora = datetime.now()

    def _flush():
        if not buffer:
            return
        pendente, buffer[:] = list(buffer), []
        etl_db.ping(ref["conn"])
        etl_db.com_reconexao(ref, lambda c: _gravar(c, pendente))
        print("  [{}] {} produtos no catalogo ({} erros)".format(
            tenant, contador["ok"], contador["erro"]), flush=True)

    def _gravar(conn, linhas):
        with etl_db.transacao(conn):
            etl_db.executar_lote(conn, SQL_UPSERT, linhas, tamanho=200)

    def _proc(pid):
        if parar.is_set():
            return None
        try:
            prod = desempacotar_data(client.obter_produto(pid))
        except TokenExpiredError:
            parar.set()
            return {"token": True}
        except Exception:
            return {"erro": True}
        return (tenant, pid, T.txt(prod.get("sku"), 80) or "",
                T.txt(prod.get("descricao") or prod.get("nome"), 255) or "",
                T.txt(prod.get("tipo"), 1), T.txt(prod.get("situacao"), 1),
                _categoria(prod), agora)

    token_morreu = False
    try:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = [ex.submit(_proc, pid) for pid in ids]
            for fut in as_completed(futs):
                r = fut.result()
                if r is None:
                    continue
                if isinstance(r, dict):
                    if r.get("token"):
                        token_morreu = True
                        ex.shutdown(wait=False, cancel_futures=True)
                        break
                    with lock:
                        contador["erro"] += 1
                    continue
                with lock:
                    buffer.append(r)
                    contador["ok"] += 1
                    if len(buffer) >= lote:
                        _flush()
    finally:
        with lock:
            _flush()

    if token_morreu:
        raise etl_extract.ParadaToken(tenant)
    print("  [{}] fase 3: {} produtos, {} erros.".format(
        tenant, contador["ok"], contador["erro"]))
    return contador["ok"], contador["erro"]
