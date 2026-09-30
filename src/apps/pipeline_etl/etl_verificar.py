"""Fase 5 -- reconciliacao e regressao. ZERO requests.

Toda checagem sai do que ja esta no banco. A ancora e `etl_janela.total_api`:
aquele numero veio do `paginacao.total` que o servidor respondeu durante a fase
1, entao da para provar completude sem falar com a API de novo.

Saida: reports/etl_verificacao_<tenants>_<timestamp>.csv (utf-8-sig, que e o que
o Excel pt-BR abre certo). Exit 4 no orquestrador se alguma checagem reprovar.
"""

import csv
import json
from datetime import datetime

import _bootstrap
import etl_db

# Regressao e casos aceitos vem de verificacao.json (TINY_APPS_CONFIG ou data/,
# porque sao numeros de faturamento e ids de pedido reais). Modelo em
# data/verificacao.exemplo.json. Sem o arquivo a regressao e PULADA com aviso e
# as views de reconciliacao rodam normalmente.
#
# REGRESSAO: uma janela curta ja carregada cujos numeros foram medidos por um
# codigo independente. Reproduzi-los ao centavo prova a cadeia inteira
# (extract -> raw -> transform -> load -> SQL) num comando so.
#
# AJUSTE: o numero de referencia e um RETRATO, e pedido muda de situacao depois
# (ex.: Enviada -> Cancelada). Quando o delta for explicado pedido a pedido, ele
# entra em "ajuste" com o motivo, para o teste continuar servindo de alarme de
# regressao de CODIGO em vez de reprovar para sempre e ser ignorado. Se o
# ajuste crescer com o tempo, troque a janela de referencia.
#
# ARITMETICA ACEITA: caso conhecido, medido e irreparavel DESTE lado (cabecalho
# da propria API inconsistente com os itens) fica listado com id, para a
# checagem continuar sendo alarme do que e NOVO. Nao acrescente id sem antes
# abrir o payload.
from core import config as _config  # noqa: E402

VERIFICACAO_FILE = _config.caminho("verificacao.json")
CAMPOS_REGRESSAO = ("pedidos", "linhas", "unidades", "skus", "faturamento")


def carregar_verificacao() -> dict:
    if not VERIFICACAO_FILE.exists():
        return {}
    return json.loads(VERIFICACAO_FILE.read_text(encoding="utf-8"))


def _aritmetica_aceita(cfg: dict) -> set:
    return {(str(t), int(i)) for t, i, *_ in cfg.get("aritmetica_aceita", [])}


SQL_REGRESSAO = """
SELECT p.tenant,
       COUNT(DISTINCT p.id_pedido)                    AS pedidos,
       COUNT(*)                                       AS linhas,
       SUM(i.quantidade)                              AS unidades,
       COUNT(DISTINCT i.sku)                          AS skus,
       ROUND(SUM(i.quantidade * i.valor_unitario), 2) AS faturamento
  FROM pedido p
  JOIN pedido_item i ON i.tenant = p.tenant AND i.id_pedido = p.id_pedido
 WHERE p.tenant = %s
   AND p.data_pedido BETWEEN %s AND %s
   AND p.situacao IN ({situacoes})
   AND p.cliente_excluido = 0
   AND i.qtd_zero = 0
 GROUP BY p.tenant
"""

# (rotulo, view, reprova?, aceitos) -- view que volta vazia = passou.
_VAZIAS = [
    ("Completude da janela (total_api x gravado)", "v_recon_janela", True, None),
    ("Pedido transformado sem linha em pedido", "v_recon_raw_sem_pedido", True, None),
    ("Pedido sem registro de estado", "v_recon_pedido_sem_raw", True, None),
    ("Pedido detalhado sem item", "v_recon_pedido_sem_item", True, None),
    ("Aritmetica (soma itens x valorTotalProdutos)", "v_recon_aritmetica", True,
     "aritmetica_aceita"),
    ("Fidelidade raw x normalizado", "v_recon_fidelidade", True, None),
    ("numero_pedido duplicado", "v_recon_numero_duplicado", False, None),
]


def _linha(checagem, esperado, obtido, status, detalhe=""):
    return {"checagem": checagem, "esperado": esperado, "obtido": obtido,
            "status": status, "detalhe": detalhe}


def verificar(conn, tenants: list, com_regressao: bool = True) -> tuple:
    """Devolve (linhas_do_relatorio, passou)."""
    linhas, passou = [], True
    cfg = carregar_verificacao()

    for rotulo, view, reprova, chave_aceitos in _VAZIAS:
        aceitos = _aritmetica_aceita(cfg) if chave_aceitos else None
        try:
            if aceitos:
                # Com lista de aceitos a view vem inteira (sao poucas linhas por
                # construcao) e o filtro e feito aqui, para o relatorio poder
                # dizer quantos casos conhecidos foram perdoados.
                todas = etl_db.consultar(conn, "SELECT * FROM " + view)
                restantes = [r for r in todas
                             if (r["tenant"], r["id_pedido"]) not in aceitos]
                n, perdoados = len(restantes), len(todas) - len(restantes)
            else:
                n = etl_db.um_valor(conn, "SELECT COUNT(*) FROM " + view)
                restantes, perdoados = None, 0
        except Exception as e:
            linhas.append(_linha(rotulo, "0 linhas", "erro", "ERRO", str(e)[:300]))
            passou = False
            continue
        ok = (n == 0)
        if not ok and reprova:
            passou = False
        detalhe = ""
        if not ok:
            amostra = (restantes[:3] if restantes is not None else
                       etl_db.consultar(conn, "SELECT * FROM {} LIMIT 3".format(view)))
            detalhe = "; ".join(str(a)[:160] for a in amostra)
        elif perdoados:
            detalhe = ("{} caso(s) conhecido(s) da origem, aceitos "
                       "em verificacao.json").format(perdoados)
        linhas.append(_linha(rotulo, "0 linhas", "{} linhas".format(n),
                             "OK" if ok else ("FALHOU" if reprova else "AVISO"),
                             detalhe))

    # Censo por estado -- 'listado' significa fase 2 incompleta; 'erro' sai
    # nomeado para o operador decidir.
    for r in etl_db.consultar(conn, "SELECT * FROM v_recon_estado ORDER BY 1,2"):
        estado = r["estado"]
        ruim = estado in ("listado", "erro")
        if ruim:
            passou = False
        linhas.append(_linha(
            "Estado [{}] {}".format(r["tenant"], estado),
            "0 (fora de 'transformado')" if ruim else "-",
            "{} pedidos".format(r["pedidos"]),
            "FALHOU" if ruim else "OK",
            "{} a {}".format(r["de"], r["ate"])))

    if com_regressao:
        linhas.extend(regressao(conn, tenants, cfg.get("regressao")))
        for lin in linhas:
            if lin["checagem"].startswith("Regressao") and lin["status"] == "FALHOU":
                passou = False

    return linhas, passou


def regressao(conn, tenants: list, cfg: dict | None) -> list:
    """Compara a janela de referencia com os numeros conhecidos."""
    if not cfg:
        return [_linha("Regressao", "-", "sem referencia", "PULADO",
                       "verificacao.json ausente ou sem 'regressao'")]
    ini, fim = cfg["inicio"], cfg["fim"]
    situacoes = [int(s) for s in cfg.get("situacoes", (1, 3, 4, 5, 6, 7))]
    sql = SQL_REGRESSAO.format(situacoes=",".join(["%s"] * len(situacoes)))
    linhas = []
    for tenant in tenants:
        esperado = (cfg.get("esperado") or {}).get(tenant)
        if not esperado:
            continue
        r = etl_db.consultar(conn, sql, (tenant, ini, fim, *situacoes))
        if not r:
            linhas.append(_linha(
                "Regressao {} ({} a {})".format(tenant, ini, fim),
                "{} pedidos".format(esperado["pedidos"]), "sem dados",
                "PULADO", "a janela de referencia nao esta carregada neste banco"))
            continue
        got = r[0]
        ajuste = (cfg.get("ajuste") or {}).get(tenant) or {}
        for campo in CAMPOS_REGRESSAO:
            referencia = esperado[campo]
            # skus e contagem distinta e portanto nao seria aditiva em teoria;
            # o delta do ajuste tem de ser MEDIDO, nao deduzido.
            alvo = round(referencia - ajuste.get(campo, 0), 2)
            obtido = float(got[campo] or 0)
            ok = abs(obtido - float(alvo)) < (0.01 if campo == "faturamento" else 0.5)
            detalhe = ""
            if ajuste.get(campo):
                detalhe = "referencia {} menos {} ({})".format(
                    referencia, ajuste[campo], ajuste.get("motivo", ""))
            linhas.append(_linha(
                "Regressao {} / {}".format(tenant, campo),
                alvo, round(obtido, 2) if campo == "faturamento" else int(obtido),
                "OK" if ok else "FALHOU", detalhe))
    return linhas


def resumo_negocio(conn) -> list:
    """Nao reprova nada -- so mostra o retrato, para um mes esquisito aparecer
    antes de virar relatorio errado."""
    saida = []
    try:
        saida.append(("Por canal", etl_db.consultar(
            conn, "SELECT * FROM v_resumo_canal ORDER BY tenant, ano_mes, "
                  "faturamento DESC")))
        saida.append(("Top integracoes", etl_db.consultar(
            conn, "SELECT tenant, integracao, SUM(pedidos) AS pedidos, "
                  "ROUND(SUM(faturamento),2) AS faturamento "
                  "FROM v_resumo_integracao GROUP BY tenant, integracao "
                  "ORDER BY tenant, pedidos DESC LIMIT 20")))
    except Exception:
        pass
    return saida


def salvar_csv(linhas: list, tenants: list):
    _bootstrap.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    nome = "etl_verificacao_{}_{}.csv".format(
        "_".join(tenants), datetime.now().strftime("%Y%m%d_%H%M%S"))
    caminho = _bootstrap.REPORTS_DIR / nome
    campos = ["checagem", "esperado", "obtido", "status", "detalhe"]
    with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        for lin in linhas:
            w.writerow({c: lin.get(c, "") for c in campos})
    return caminho


def imprimir(linhas: list) -> None:
    largura = max((len(str(lin["checagem"])) for lin in linhas), default=20)
    for lin in linhas:
        marca = {"OK": "  ok  ", "FALHOU": " FALHA", "AVISO": " aviso",
                 "PULADO": " pula ", "ERRO": " ERRO "}.get(lin["status"], "  ?   ")
        print("  [{}] {:<{w}}  esperado={}  obtido={}".format(
            marca, lin["checagem"], lin["esperado"], lin["obtido"], w=largura))
        if lin.get("detalhe"):
            print("           {}".format(lin["detalhe"][:200]))


def fase5_verificar(conn, tenants: list, com_regressao: bool = True) -> bool:
    print("\n  Verificacao (0 requests):")
    linhas, passou = verificar(conn, tenants, com_regressao)
    imprimir(linhas)
    caminho = salvar_csv(linhas, tenants)
    print("\n  Relatorio: {}".format(caminho))
    print("  Resultado: {}".format("TUDO OK" if passou else "REPROVOU"))
    return passou
