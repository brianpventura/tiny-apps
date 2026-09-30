"""Load: upsert em lote no MySQL, com o estado marcado na MESMA transacao.

E aqui que mora o requisito explicito do pedido -- "registrar os dados conforme o
sistema rode, para que, caso o pipeline de erro, os dados estejam no banco":
`gravar_lote_lista` e `gravar_lote_detalhe` persistem N pedidos E marcam o estado
de cada um numa transacao so. Depois do commit, nenhum deles sera buscado de novo.

Tres regras que o arquivo inteiro obedece:

  * `INSERT ... ON DUPLICATE KEY UPDATE` com alias de linha (`AS novo`). Nunca
    `REPLACE INTO`: ele faz DELETE+INSERT e, com ON DELETE CASCADE, apagaria em
    silencio todos os itens do pedido. E nunca `VALUES()`, depreciado desde
    8.0.20 -- encheria o log de warning a cada statement de um backfill.
  * Refazer um lote e sempre seguro, porque todo write e idempotente sobre
    chave natural. E isso que permite ao `com_reconexao` simplesmente repetir.
  * Nunca commitar entre o DELETE e o INSERT dos itens: uma queda ali deixaria
    o pedido sem itens, erro que some em qualquer SUM().
"""

import json
from datetime import datetime

import _bootstrap  # noqa: F401
import etl_db

# ------------------------------------------------------------------ #
#  Colunas -- a ordem vive AQUI e em nenhum outro lugar               #
# ------------------------------------------------------------------ #

COLS_PEDIDO = [
    "tenant", "id_pedido", "numero_pedido", "id_nota_fiscal", "origem_pedido",
    "data_pedido", "data_prevista", "data_entrega", "data_faturamento",
    "data_envio", "situacao", "situacao_nome",
    "valor_total_produtos", "valor_total_pedido", "valor_desconto",
    "valor_frete", "valor_outras_despesas",
    "numero_ordem_compra", "observacoes", "observacoes_internas",
    "id_cliente", "cliente_nome", "cliente_cpf_cnpj", "cliente_tipo_pessoa",
    "cliente_email", "cliente_telefone", "cliente_celular", "cliente_uf",
    "cliente_municipio",
    "id_natureza", "natureza_nome", "natureza_norm",
    "id_vendedor", "vendedor_nome",
    "id_ecommerce", "ecommerce_nome", "ecommerce_canal",
    "numero_pedido_ecommerce", "numero_pedido_canal",
    "id_transportador", "transportador_nome", "frete_por_conta",
    "codigo_rastreamento", "url_rastreamento", "id_forma_envio",
    "forma_envio_nome", "id_forma_frete", "forma_frete_nome",
    "id_deposito", "deposito_nome", "id_intermediador", "intermediador_nome",
    "id_lista_preco", "lista_preco_nome", "lista_preco_acrescimo",
    "id_forma_recebimento", "forma_recebimento_nome", "id_meio_pagamento",
    "meio_pagamento_nome", "condicao_pagamento", "id_categoria_financeira",
    "categoria_financeira",
    "eh_situacao_valida", "eh_natureza_alvo", "cliente_excluido", "canal",
    "conflito_canal", "tem_detalhe", "regras_versao",
    "qtd_itens", "total_unidades", "criado_em", "atualizado_em",
]

COLS_ITEM = ["tenant", "id_pedido", "seq", "id_produto", "sku",
             "descricao_item", "tipo_item", "quantidade", "valor_unitario",
             "info_adicional"]

COLS_ENDERECO = ["tenant", "id_pedido", "tipo", "nome_destinatario", "cpf_cnpj",
                 "tipo_pessoa", "telefone", "inscricao_estadual", "endereco",
                 "numero", "complemento", "bairro", "municipio", "cep", "uf",
                 "pais"]

COLS_PARCELA = ["tenant", "id_pedido", "seq", "dias", "data_vencimento",
                "valor", "observacoes", "id_forma_recebimento",
                "forma_recebimento_nome", "id_meio_pagamento",
                "meio_pagamento_nome"]

COLS_PAGAMENTO = ["tenant", "id_pedido", "seq", "valor", "tipo_pagamento",
                  "cnpj_intermediador", "codigo_autorizacao", "codigo_bandeira"]

COLS_MARCADOR = ["tenant", "id_pedido", "marcador"]

# (tabela, colunas, chaves da PK, colunas que o UPDATE nao pode sobrescrever)
_DIMS = {
    "cliente": ("dim_cliente",
                ["tenant", "id_cliente", "nome", "nome_norm", "codigo",
                 "fantasia", "tipo_pessoa", "cpf_cnpj", "inscricao_estadual",
                 "rg", "telefone", "celular", "email", "endereco", "numero",
                 "complemento", "bairro", "municipio", "cep", "uf", "pais",
                 "eh_cliente_grupo", "primeiro_visto", "ultimo_visto"],
                ["tenant", "id_cliente"],
                ["primeiro_visto"]),
    "vendedor": ("dim_vendedor",
                 ["tenant", "id_vendedor", "nome", "eh_pessoa", "ultimo_visto"],
                 ["tenant", "id_vendedor"], []),
    "ecommerce": ("dim_ecommerce",
                  ["tenant", "id_ecommerce", "nome", "canal_venda",
                   "ultimo_visto"],
                  ["tenant", "id_ecommerce"], []),
    "natureza": ("dim_natureza_operacao",
                 ["tenant", "id_natureza", "nome", "nome_norm",
                  "eh_natureza_alvo", "ultimo_visto"],
                 ["tenant", "id_natureza"], []),
    "transportador": ("dim_transportador",
                      ["tenant", "id_transportador", "nome", "ultimo_visto"],
                      ["tenant", "id_transportador"], []),
    "deposito": ("dim_deposito",
                 ["tenant", "id_deposito", "nome", "ultimo_visto"],
                 ["tenant", "id_deposito"], []),
    "intermediador": ("dim_intermediador",
                      ["tenant", "id_intermediador", "nome", "cnpj",
                       "cnpj_instituicao", "ultimo_visto"],
                      ["tenant", "id_intermediador"], []),
    "referencia": ("dim_referencia",
                   ["tenant", "dominio", "id_ref", "nome", "ultimo_visto"],
                   ["tenant", "dominio", "id_ref"], []),
    # O produto vindo do item traz so sku/descricao; tipo_cadastro e categoria
    # sao da fase 3 e NAO podem ser sobrescritos aqui.
    "produto": ("dim_produto",
                ["tenant", "id_produto", "sku", "descricao", "atualizado_em"],
                ["tenant", "id_produto"], []),
}


def _sql_upsert(tabela: str, colunas: list, chaves: list,
                nao_atualizar: list | None = None) -> str:
    nao = set(chaves) | set(nao_atualizar or [])
    sets = ["{c} = novo.{c}".format(c=c) for c in colunas if c not in nao]
    return ("INSERT INTO {tab} ({cols}) VALUES ({ph}) AS novo"
            " ON DUPLICATE KEY UPDATE {sets}".format(
                tab=tabela, cols=", ".join(colunas),
                ph=", ".join(["%s"] * len(colunas)),
                sets=", ".join(sets) if sets else colunas[0] + " = novo." + colunas[0]))


def _linhas(dicts, colunas):
    return [tuple(d.get(c) for c in colunas) for d in dicts]


# data_raw entra aqui, e nao como coluna GENERATED, porque sob sql_mode
# estrito uma data torta dentro de coluna gerada levanta erro no INSERT e
# derruba o lote inteiro. Ver sql/004_data_raw_comum.sql.
SQL_RAW = _sql_upsert(
    "pedido_raw",
    ["tenant", "id_pedido", "fonte", "payload_hash", "payload", "data_raw",
     "extraido_em", "visto_em"],
    ["tenant", "id_pedido", "fonte", "payload_hash"],
    nao_atualizar=["payload", "data_raw", "extraido_em"])

SQL_PEDIDO = _sql_upsert("pedido", COLS_PEDIDO, ["tenant", "id_pedido"],
                         nao_atualizar=["criado_em"])
SQL_ITEM = _sql_upsert("pedido_item", COLS_ITEM,
                       ["tenant", "id_pedido", "seq"])
SQL_ENDERECO = _sql_upsert("pedido_endereco", COLS_ENDERECO,
                           ["tenant", "id_pedido", "tipo"])
SQL_PARCELA = _sql_upsert("pedido_parcela", COLS_PARCELA,
                          ["tenant", "id_pedido", "seq"])
SQL_PAGAMENTO = _sql_upsert("pedido_pagamento_integrado", COLS_PAGAMENTO,
                            ["tenant", "id_pedido", "seq"])
SQL_MARCADOR = ("INSERT IGNORE INTO pedido_marcador (tenant, id_pedido, "
                "marcador) VALUES (%s, %s, %s)")

SQL_ESTADO_LISTA = (
    "INSERT INTO etl_pedido_estado (tenant, id_pedido, estado, data_ref, "
    "numero_pedido, situacao, hash_lista, listado_em, atualizado_em) "
    "VALUES (%s,%s,'listado',%s,%s,%s,%s,%s,%s) AS novo "
    "ON DUPLICATE KEY UPDATE "
    "  data_ref = novo.data_ref, numero_pedido = novo.numero_pedido, "
    "  situacao = novo.situacao, listado_em = novo.listado_em, "
    "  atualizado_em = novo.atualizado_em, "
    # Conteudo da listagem mudou (situacao, valor, o que for): o pedido volta
    # para 'listado' e sera re-detalhado. Identico: mantem o estado atual.
    "  estado = IF(etl_pedido_estado.hash_lista <> novo.hash_lista, "
    "              'listado', etl_pedido_estado.estado), "
    "  tentativas = IF(etl_pedido_estado.hash_lista <> novo.hash_lista, "
    "                  0, etl_pedido_estado.tentativas), "
    "  hash_lista = novo.hash_lista")

SQL_ESTADO_DETALHE_OK = (
    "UPDATE etl_pedido_estado SET estado = 'detalhado', hash_detalhe = %s, "
    "situacao = COALESCE(%s, situacao), detalhado_em = %s, atualizado_em = %s, "
    "tentativas = 0, erro_msg = NULL "
    "WHERE tenant = %s AND id_pedido = %s")

SQL_ESTADO_DETALHE_ERRO = (
    "UPDATE etl_pedido_estado SET tentativas = tentativas + 1, erro_msg = %s, "
    "estado = IF(tentativas + 1 >= %s, 'erro', estado), atualizado_em = %s "
    "WHERE tenant = %s AND id_pedido = %s")

SQL_ESTADO_TRANSFORMADO = (
    "UPDATE etl_pedido_estado SET estado = 'transformado', "
    "transformado_em = %s, atualizado_em = %s "
    "WHERE tenant = %s AND id_pedido = %s")

SQL_ERRO = ("INSERT INTO etl_erro (tenant, id_pedido, etapa, tabela, codigo, "
            "erro, criado_em) VALUES (%s,%s,%s,%s,%s,%s,%s)")


# ------------------------------------------------------------------ #
#  Gravacao incremental                                               #
# ------------------------------------------------------------------ #

def gravar_lote_lista(conn, registros: list) -> int:
    """Persiste uma pagina da listagem e marca cada pedido como 'listado'.

    registros: [{"id_pedido","payload","hash","data_ref","numero","situacao"}]
    Commit por pagina de 100 -- uma queda custa no maximo 1 request.
    """
    if not registros:
        return 0
    agora = datetime.now()
    raw, estados = [], []
    for r in registros:
        raw.append((r["tenant"], r["id_pedido"], "listagem", r["hash"],
                    r["payload"], r.get("data_ref"), agora, agora))
        estados.append((r["tenant"], r["id_pedido"], r.get("data_ref"),
                        r.get("numero"), r.get("situacao"), r["hash"],
                        agora, agora))
    with etl_db.transacao(conn):
        etl_db.executar_lote(conn, SQL_RAW, raw, tamanho=200)
        etl_db.executar_lote(conn, SQL_ESTADO_LISTA, estados, tamanho=200)
    return len(registros)


def gravar_lote_detalhe(conn, tenant: str, registros: list,
                        max_tentativas: int = 5) -> tuple:
    """Persiste N detalhes e marca o estado DE CADA UM na MESMA transacao.

    registros: [{"id_pedido", "payload", "hash", "situacao"}] para sucesso,
               [{"id_pedido", "erro": "..."}] para falha.
    Depois do commit, nenhum desses pedidos sera buscado de novo -- e a funcao
    que garante que uma queda no meio de um backfill de 5h nao perde trabalho.
    """
    if not registros:
        return 0, 0
    agora = datetime.now()
    raw, ok, falhas = [], [], []
    for r in registros:
        if r.get("erro"):
            falhas.append((str(r["erro"])[:500], max_tentativas, agora,
                           tenant, r["id_pedido"]))
            continue
        raw.append((tenant, r["id_pedido"], "detalhe", r["hash"], r["payload"],
                    r.get("data_ref"), agora, agora))
        ok.append((r["hash"], r.get("situacao"), agora, agora, tenant,
                   r["id_pedido"]))
    with etl_db.transacao(conn):
        etl_db.executar_lote(conn, SQL_RAW, raw, tamanho=200)
        etl_db.executar_lote(conn, SQL_ESTADO_DETALHE_OK, ok, tamanho=500)
        etl_db.executar_lote(conn, SQL_ESTADO_DETALHE_ERRO, falhas, tamanho=500)
    return len(ok), len(falhas)


def gravar_dims(conn, dims_acumuladas: dict) -> int:
    """Upsert das dimensoes. Roda dentro da transacao do chamador."""
    n = 0
    for chave, registros in dims_acumuladas.items():
        if not registros:
            continue
        tabela, colunas, chaves, nao_atualizar = _DIMS[chave]
        # Deduplica dentro do lote: o mesmo cliente aparece em varios pedidos e
        # executemany com a chave repetida so gasta round-trip.
        vistos, unicos = set(), []
        for r in registros:
            k = tuple(r.get(c) for c in chaves)
            if k not in vistos:
                vistos.add(k)
                unicos.append(r)
        sql = _sql_upsert(tabela, colunas, chaves, nao_atualizar)
        n += etl_db.executar_lote(conn, sql, _linhas(unicos, colunas))
    return n


def carregar_bloco(conn, tenant: str, blocos: list) -> int:
    """Fase 4: grava pedido + filhas + dims de N pedidos e marca 'transformado'.

    blocos: lista do que `etl_transform.transformar` devolveu.
    Tudo numa transacao so -- em especial o DELETE+INSERT dos itens.
    """
    if not blocos:
        return 0
    agora = datetime.now()
    dims = {k: [] for k in _DIMS}
    for b in blocos:
        for k, v in b["dims"].items():
            dims[k].extend(v)

    with etl_db.transacao(conn):
        etl_db.executar_lote(conn, SQL_PEDIDO,
                             _linhas([b["pedido"] for b in blocos], COLS_PEDIDO),
                             tamanho=200)
        gravar_dims(conn, dims)

        ids = [(b["pedido"]["tenant"], b["pedido"]["id_pedido"]) for b in blocos]
        for tabela in ("pedido_item", "pedido_endereco", "pedido_parcela",
                       "pedido_pagamento_integrado", "pedido_marcador"):
            etl_db.executar_lote(
                conn,
                "DELETE FROM {} WHERE tenant = %s AND id_pedido = %s".format(
                    tabela), ids, tamanho=500)

        for sql, chave, colunas in (
                (SQL_ITEM, "itens", COLS_ITEM),
                (SQL_ENDERECO, "enderecos", COLS_ENDERECO),
                (SQL_PARCELA, "parcelas", COLS_PARCELA),
                (SQL_PAGAMENTO, "pagamentos", COLS_PAGAMENTO),
                (SQL_MARCADOR, "marcadores", COLS_MARCADOR)):
            linhas = []
            for b in blocos:
                linhas.extend(_linhas(b[chave], colunas))
            etl_db.executar_lote(conn, sql, linhas, tamanho=500)

        etl_db.executar_lote(
            conn, SQL_ESTADO_TRANSFORMADO,
            [(agora, agora, b["pedido"]["tenant"], b["pedido"]["id_pedido"])
             for b in blocos], tamanho=500)
    return len(blocos)


def carregar_bloco_resiliente(ref: dict, tenant: str, blocos: list) -> tuple:
    """carregar_bloco com a separacao INFRA x DADO.

    INFRA (servidor caiu, deadlock): `com_reconexao` reconecta e refaz o lote --
    seguro porque todo write e idempotente.
    DADO (string longa demais, FK, datetime invalido): refazer o lote nao
    adianta. Refaz pedido a pedido para isolar o culpado, manda a linha ruim
    para etl_erro e deixa o resto do lote entrar. Um pedido esquisito nao pode
    travar a carga inteira.
    """
    try:
        return etl_db.com_reconexao(
            ref, lambda c: carregar_bloco(c, tenant, blocos)), 0
    except Exception as e:
        if etl_db.eh_infra(e):
            raise
        print("  [{}] lote falhou por erro de DADO ({}); isolando pedido a "
              "pedido.".format(tenant, e), flush=True)

    ok, ruins = 0, 0
    for b in blocos:
        try:
            ok += etl_db.com_reconexao(
                ref, lambda c, bb=b: carregar_bloco(c, tenant, [bb]))
        except Exception as e2:
            ruins += 1
            registrar_erro(ref["conn"], tenant, b["pedido"]["id_pedido"],
                           "load", "pedido", etl_db.codigo_erro(e2), str(e2))
    return ok, ruins


def registrar_erro(conn, tenant, id_pedido, etapa, tabela, codigo, erro):
    try:
        with etl_db.transacao(conn):
            etl_db.executar(conn, SQL_ERRO, (tenant, id_pedido, etapa, tabela,
                                             codigo if isinstance(codigo, int)
                                             else None,
                                             str(erro)[:4000], datetime.now()))
    except Exception:
        pass


# ------------------------------------------------------------------ #
#  Fase 4 -- transform + load, 100% OFFLINE                           #
# ------------------------------------------------------------------ #

def _payloads(conn, tenant: str, ids: list) -> dict:
    """Busca a versao CORRENTE do raw de cada pedido.

    O raw e append-only e versionado; quem diz qual versao vale e o
    hash_lista/hash_detalhe do etl_pedido_estado.
    """
    if not ids:
        return {}
    marcas = ", ".join(["%s"] * len(ids))
    sql = (
        "SELECT r.id_pedido, r.fonte, r.payload "
        "FROM pedido_raw r "
        "JOIN etl_pedido_estado e "
        "  ON e.tenant = r.tenant AND e.id_pedido = r.id_pedido "
        " AND ((r.fonte = 'listagem' AND r.payload_hash = e.hash_lista) "
        "   OR (r.fonte = 'detalhe'  AND r.payload_hash = e.hash_detalhe)) "
        "WHERE r.tenant = %s AND r.id_pedido IN ({})".format(marcas))
    saida = {}
    for r in etl_db.consultar(conn, sql, tuple([tenant] + list(ids))):
        saida.setdefault(r["id_pedido"], {})[r["fonte"]] = r["payload"]
    return saida


def pendentes_transform(conn, tenant: str, refazer: bool = False,
                        regras_versao: int | None = None,
                        limite: int | None = None) -> list:
    """O que falta carregar. Sem heuristica: e uma query."""
    if refazer:
        sql = ("SELECT id_pedido FROM etl_pedido_estado "
               "WHERE tenant = %s AND estado IN ('detalhado','transformado') ")
        params = [tenant]
    elif regras_versao is not None:
        # TRES motivos para (re)normalizar, e o primeiro ja foi esquecido uma
        # vez -- um pedido mudou de situacao 5 para 6, foi re-detalhado, e a
        # fase 4 o PULOU porque ele ja estava na
        # regras_versao corrente. O warehouse ficou com a situacao velha sem
        # acusar nada. Como e exatamente esse o caminho da atualizacao diaria,
        # o bug teria feito todo incremental falhar em silencio.
        sql = ("SELECT e.id_pedido FROM etl_pedido_estado e "
               "LEFT JOIN pedido p ON p.tenant = e.tenant "
               "                  AND p.id_pedido = e.id_pedido "
               "WHERE e.tenant = %s AND e.estado IN ('detalhado','transformado') "
               "  AND (e.estado = 'detalhado'"      # raw mais novo que o pedido
               "    OR p.id_pedido IS NULL"          # nunca carregado
               "    OR p.regras_versao < %s) ")      # regra de negocio mudou
        params = [tenant, regras_versao]
    else:
        sql = ("SELECT id_pedido FROM etl_pedido_estado "
               "WHERE tenant = %s AND estado = 'detalhado' ")
        params = [tenant]
    sql += "ORDER BY data_ref DESC, id_pedido DESC"
    if limite:
        sql += " LIMIT %s"
        params.append(int(limite))
    return [r["id_pedido"] for r in etl_db.consultar(conn, sql, tuple(params))]


def fase4_carregar(ref: dict, tenant: str, regras_versao: int, lote: int = 200,
                   refazer: bool = False, limite: int | None = None) -> tuple:
    """Le o raw, transforma e grava. Nao valida token e nao faz request.

    `regras_versao` e a versao corrente do perfil de regras
    (etl_regras.versao_regras): e gravada em cada pedido e todo pedido com
    versao menor e renormalizado.

    E de proposito que esta fase seja offline: com o raw ja em casa, da para
    re-normalizar o banco inteiro depois que o refresh token morrer.
    """
    import etl_transform as T

    ids = pendentes_transform(ref["conn"], tenant, refazer, regras_versao,
                              limite)
    if not ids:
        print("  [{}] fase 4: nada a carregar.".format(tenant))
        return 0, 0
    print("  [{}] fase 4: {} pedidos a normalizar (offline).".format(
        tenant, len(ids)))

    ctx = T.contexto(tenant, regras_versao)
    carregados, ruins = 0, 0
    t0 = datetime.now()
    for i in range(0, len(ids), lote):
        pedaco = ids[i:i + lote]
        brutos = _payloads(ref["conn"], tenant, pedaco)
        blocos = []
        for pid in pedaco:
            fontes = brutos.get(pid) or {}
            if "detalhe" not in fontes and "listagem" not in fontes:
                continue
            try:
                lista_json = json.loads(fontes["listagem"]) if fontes.get("listagem") else {}
                det_json = json.loads(fontes["detalhe"]) if fontes.get("detalhe") else None
                blocos.append(T.transformar(ctx, pid, lista_json, det_json))
            except Exception as e:
                ruins += 1
                registrar_erro(ref["conn"], tenant, pid, "transform", None,
                               None, e)
        ok, falhas = carregar_bloco_resiliente(ref, tenant, blocos)
        carregados += ok
        ruins += falhas
        seg = max(1e-9, (datetime.now() - t0).total_seconds())
        print("  [{}] {}/{} normalizados ({:.0f}/s)".format(
            tenant, min(i + lote, len(ids)), len(ids), carregados / seg),
            flush=True)
    print("  [{}] fase 4: {} carregados, {} com problema.".format(
        tenant, carregados, ruins))
    return carregados, ruins
