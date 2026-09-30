"""Sondas -- descobrem a realidade da API gastando poucos requests.

Duas coisas que o repositorio NAO sabia antes deste pipeline:

1. `sondar_schema` (~4 requests). Nao existe nenhum dump cru de pedido no repo:
   todos os caches guardam recorte feito pelo app antes de gravar. A doc oficial
   lista blocos que o repo nunca tocou (pagamento/parcelas, transportador,
   deposito, intermediador, listaPreco, enderecoEntrega completo,
   pagamentosIntegrados, marcadores) e nao ha evidencia local de quais vem
   preenchidos. Esta sonda grava o JSON integro em data/ e diz quais blocos
   existem de verdade.

2. `sondar_filtro` (~4 requests). A API do Tiny ja provou ACEITAR E IGNORAR
   filtro: `dataInicio`/`dataFim` sao ignorados e so `dataInicial`/`dataFinal`
   funcionam. Um filtro ignorado responde 200 e nao dispara alarme nenhum -- no
   incremental isso significaria re-detalhar 130 mil pedidos ou concluir que
   nada mudou. Por isso `dataAtualizacao` so e usado depois de aprovado aqui.
"""

import json
from datetime import date, datetime, timedelta

import _bootstrap
import etl_db
from core.produtos_utils import desempacotar_data

# Blocos que a doc oficial promete mas o repositorio nunca leu.
BLOCOS_NAO_COMPROVADOS = [
    "pagamento", "parcelas", "pagamentosIntegrados", "transportador",
    "deposito", "intermediador", "listaPreco", "enderecoEntrega", "marcadores",
    "observacoes", "observacoesInternas", "numeroOrdemCompra", "idNotaFiscal",
    "dataFaturamento", "dataEntrega", "dataEnvio", "origemPedido",
]


def _preenchido(v) -> bool:
    return v not in (None, "", [], {}, 0)


def sondar_schema(client, tenant: str, inicio: date, fim: date,
                  quantos: int = 3) -> dict:
    """Grava pedidos crus em data/amostra_pedido_<tenant>.json e relata."""
    print("\n  [{}] sondando o schema de {} a {} ...".format(tenant, inicio, fim))
    resp = desempacotar_data(client.listar_pedidos(
        offset=0, limit=quantos, dataInicial=inicio.isoformat(),
        dataFinal=fim.isoformat()))
    itens = resp.get("itens") or []
    total = (resp.get("paginacao") or {}).get("total")
    print("  paginacao.total da janela = {}".format(total))
    if not itens:
        print("  [AVISO] nenhum pedido na janela; escolha outra com --inicio/--fim.")
        return {}

    print("  chaves da LISTAGEM: {}".format(sorted(itens[0].keys())))

    amostras, cobertura = [], {}
    for it in itens[:quantos]:
        pid = it.get("id")
        det = desempacotar_data(client.obter_pedido(pid))
        amostras.append({"id": pid, "listagem": it, "detalhe": det})
        for bloco in BLOCOS_NAO_COMPROVADOS:
            if _preenchido(det.get(bloco)):
                cobertura[bloco] = cobertura.get(bloco, 0) + 1

    d0 = amostras[0]["detalhe"]
    print("  chaves do DETALHE: {}".format(sorted(d0.keys())))
    linhas = d0.get("itens") or []
    if linhas:
        print("  chaves do ITEM[0]: {}".format(sorted(linhas[0].keys())))
        print("  item[0] = {}".format(
            json.dumps(linhas[0], ensure_ascii=False)[:500]))
        soma = sum(float(i.get("quantidade") or 0) * float(i.get("valorUnitario") or 0)
                   for i in linhas)
        vprod = float(d0.get("valorTotalProdutos") or 0)
        print("  integridade: soma(qtd*valorUnitario)={:.2f} vs "
              "valorTotalProdutos={:.2f} (dif {:+.2f})".format(
                  soma, vprod, soma - vprod))

    print("\n  Blocos NAO comprovados pelo repositorio, em {} pedido(s):".format(
        len(amostras)))
    for bloco in BLOCOS_NAO_COMPROVADOS:
        n = cobertura.get(bloco, 0)
        print("    {:<22} {}".format(
            bloco, "PREENCHIDO em {}/{}".format(n, len(amostras)) if n
            else "vazio/ausente"))

    _bootstrap.DATA_DIR.mkdir(parents=True, exist_ok=True)
    destino = _bootstrap.DATA_DIR / "amostra_pedido_{}.json".format(tenant)
    destino.write_text(json.dumps(amostras, ensure_ascii=False, indent=2),
                       encoding="utf-8")
    print("\n  Amostra crua gravada em {}".format(destino))
    print("  [AVISO] o arquivo tem nome, CPF/CNPJ e endereco de cliente real -- "
          "nao versionar nem colar em chat.")
    return {"cobertura": cobertura, "arquivo": str(destino)}


def sondar_filtro(client, conn, tenant: str) -> dict:
    """4 requests: dataAtualizacao filtra de verdade? E com que semantica?"""
    print("\n  [{}] sondando o filtro dataAtualizacao (4 requests) ...".format(
        tenant))
    ontem = (date.today() - timedelta(days=1)).isoformat()
    antigo = "2024-01-01"

    def _total(**filtros):
        resp = desempacotar_data(client.listar_pedidos(offset=0, limit=1,
                                                       **filtros))
        return (resp.get("paginacao") or {}).get("total"), resp

    t_base, _ = _total()
    t_ontem, _ = _total(dataAtualizacao=ontem)
    t_antigo, _ = _total(dataAtualizacao=antigo)
    _, amostra = _total(dataAtualizacao=ontem)

    print("    total sem filtro           = {}".format(t_base))
    print("    total dataAtualizacao=ontem= {}".format(t_ontem))
    print("    total dataAtualizacao=2024 = {}".format(t_antigo))

    if t_base is None or t_ontem is None:
        suportado, semantica = 0, "indefinido"
    elif t_ontem == t_base:
        # Mesmo sintoma de dataInicio/dataFim: a API aceita e ignora.
        suportado, semantica = 0, "ignorado"
    elif t_antigo is not None and t_antigo >= t_base * 0.8:
        suportado, semantica = 1, "maior_igual"
    else:
        suportado, semantica = 1, "igual"

    veredito = {"total_base": t_base, "total_ontem": t_ontem,
                "total_antigo": t_antigo,
                "chaves_item": sorted((amostra.get("itens") or [{}])[0].keys())}

    etl_db.executar(
        conn,
        "INSERT INTO etl_sonda_filtro (tenant, filtro, suportado, semantica, "
        "evidencia, testado_em) VALUES (%s,%s,%s,%s,%s,%s) AS novo "
        "ON DUPLICATE KEY UPDATE suportado = novo.suportado, "
        "semantica = novo.semantica, evidencia = novo.evidencia, "
        "testado_em = novo.testado_em",
        (tenant, "dataAtualizacao", suportado, semantica,
         json.dumps(veredito, ensure_ascii=False, default=str), datetime.now()))
    conn.commit()

    if suportado:
        print("    veredito: SUPORTADO, semantica '{}'. "
              "--usar-data-atualizacao liberado.".format(semantica))
    else:
        print("    veredito: NAO usar. O filtro foi {} -- o incremental fica "
              "na janela deslizante (--dias-retro), que custa ~25 s de "
              "listagem por tenant.".format(semantica))
    return {"suportado": suportado, "semantica": semantica}


def filtro_aprovado(conn, tenant: str) -> bool:
    return bool(etl_db.um_valor(
        conn, "SELECT suportado FROM etl_sonda_filtro "
              "WHERE tenant = %s AND filtro = 'dataAtualizacao'", (tenant,)))
