"""
Produtos mais vendidos por CANAL/INTEGRACAO numa janela — le o warehouse.

Responde: quais produtos sairam por UMA integracao especifica no periodo, com
SKU, nome, preco unitario, quantidade e faturamento bruto. Le `pedido`,
`pedido_item` e `dim_produto` do MySQL `tiny_dw` (carregado por
`pipeline_pedidos.py`); nao faz request nenhum a API. Qualquer janela ja
carregada no banco pode ser consultada.

Pre-requisito: a janela precisa estar carregada no warehouse
(`pipeline_pedidos.py <tenant> --inicio ... --fim ...` ou `--incremental`). A
fase 3 do pipeline e o que da a coluna Tipo (kit/materia-prima) e o nome
canonico do cadastro.

DECISOES DE NEGOCIO (do perfil de regras da empresa, core/regras_vendas.py --
um lugar so, para nao divergir; relatorio 'produtos_canal' pode ter recorte
proprio em regras.json -> relatorios.produtos_canal):
  - Situacoes validas: vendas.situacoes_validas do perfil. Sobrescrevivel com
    --situacoes.
  - Naturezas: so as de venda (vendas.naturezas do perfil).
    --sem-filtro-natureza desliga.
  - Transferencia interna entre empresas do grupo excluida, incluindo a
    razao_social do proprio tenant.
  - Kits NAO sao explodidos (o Tiny entrega o kit destrinchado em componentes
    dentro do pedido); linha de kit fica rotulada Tipo=K e o Diagnostico conta
    quantas apareceram.
  - Valor por produto e BRUTO: soma(quantidade x valorUnitario). Nao existe
    desconto em nivel de item na API, e valorDesconto/valorFrete sao do pedido
    inteiro — nao ha rateio.
  - "Preco unit. medio" e media PONDERADA (faturamento / quantidade). As
    colunas Preco min/max e Precos dist. mostram a dispersao, porque o mesmo
    SKU sai a precos diferentes (promocao, atacado, cortesia).
  - Brinde NAO e descartado por padrao: item a valorUnitario 0 saiu do estoque
    e conta como unidade vendida, com faturamento zero. Em loja virtual isso
    pode ser uma fatia grande das unidades, entao o xlsx traz DUAS abas de
    ranking: "Top N" por quantidade e "Top N faturamento" por dinheiro.
    Quantidade e faturamento contam historias diferentes aqui.
  - `--sem-brinde` descarta as LINHAS a valorUnitario 0 — corte por linha, nao
    por SKU: o SKU que sai de graca E vendido fica so com as unidades pagas
    (ex.: 98 un a R$ 0 + 2 a R$ 5,00 viram 2 un a R$ 5,00, em vez de 100 un a
    R$ 0,10 de media). O faturamento do canal NAO muda (linha de valor zero
    soma zero); mudam quantidade, preco medio e o ranking. As contagens de
    brinde saem no Diagnostico nos dois modos.
  - Item de quantidade <= 0 fica fora (o warehouse guarda, marcado `qtd_zero`).

PRECO DE TABELA (`--precos-xlsx`): em vez do preco praticado no pedido, o
ranking pode usar o preco de uma planilha externa (ex.: export de precos da
loja virtual), casado por **SKU** (join por codigo, nao por nome). Regras do
join:
  - SKU repetido na planilha: vence a linha `product_status = ACTIVE` quando a
    coluna existe (padrao do export do Shopify); o desempate fica explicito;
  - SKU do ranking que **nao existe** na planilha nao e descartado em silencio:
    cai no preco praticado do pedido, com `Fonte do preco = pedido (fallback)`,
    e sai listado na aba `Sem preco na planilha`. `--so-com-preco` descarta.
    A tabela de uma loja virtual raramente cobre todo o catalogo vendido.
  - `Total a preco tabela R$` = quantidade x preco de tabela. E um valor
    TEORICO ("quanto teria dado ao preco do site"), nao o faturado; o faturado
    real fica na coluna ao lado e nas abas `Por dia`/`Diagnostico`, que seguem
    sempre com o preco praticado.

O filtro de canal casa por SUBSTRING normalizada (minusculas, sem acento) em
`ecommerce.nome` — a chave de agrupamento confiavel; `ecommerce.canalVenda` vem
vazio em quase tudo. O Diagnostico lista TODAS as integracoes da janela com a
contagem, marcando as que casaram, para nome errado nunca passar em silencio.

Uso:
    python relatorio_produtos_canal.py <tenant> --inicio 2026-09-04 --fim 2026-09-09 \
        --canal "shopify" --top 50
    python relatorio_produtos_canal.py <tenant> --canal shopee --top 20
    python relatorio_produtos_canal.py <tenant> --canal "" --top 50   # todos os canais
    python relatorio_produtos_canal.py <tenant> --situacoes 1,3,4,5,6,7
    python relatorio_produtos_canal.py <tenant> --canal shopify \
        --precos-xlsx tabela_site.xlsx --precos-coluna preco
"""

import sys
import argparse
from collections import Counter, defaultdict
from datetime import datetime, date, timedelta
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import _bootstrap  # noqa: E402  (efeito colateral: sys.path)

import pandas as pd  # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

import etl_db  # noqa: E402
from core import regras_vendas as rv  # noqa: E402

_REPORTS_DIR = _bootstrap.REPORTS_DIR / "pipeline_etl"

_DIAS_SEM = ["seg", "ter", "qua", "qui", "sex", "sab", "dom"]


def dias_do_periodo(inicio: date, fim: date) -> list[date]:
    return [inicio + timedelta(days=i) for i in range((fim - inicio).days + 1)]


def rotulo_dia(d: date) -> str:
    return f"{d:%d/%m} ({_DIAS_SEM[d.weekday()]})"


# ------------------------------------------------------------------ #
#  Leitura do warehouse                                               #
# ------------------------------------------------------------------ #

_SQL_PEDIDOS = """
SELECT id_pedido, data_pedido, situacao, valor_total_pedido, cliente_nome,
       ecommerce_nome, tem_detalhe, natureza_nome, valor_total_produtos,
       valor_desconto, valor_frete
  FROM pedido
 WHERE tenant = %s AND data_pedido BETWEEN %s AND %s
"""

_SQL_ITENS = """
SELECT i.id_pedido, i.sku, i.id_produto, i.descricao_item, i.quantidade,
       i.valor_unitario
  FROM pedido_item i
  JOIN pedido p ON p.tenant = i.tenant AND p.id_pedido = i.id_pedido
 WHERE i.tenant = %s AND p.data_pedido BETWEEN %s AND %s
   AND i.quantidade > 0
 ORDER BY i.id_pedido, i.seq
"""

_SQL_CATALOGO = """
SELECT DISTINCT d.id_produto, d.tipo_cadastro, d.descricao, d.categoria_caminho
  FROM dim_produto d
  JOIN pedido_item i ON i.tenant = d.tenant AND i.id_produto = d.id_produto
  JOIN pedido p ON p.tenant = i.tenant AND p.id_pedido = i.id_pedido
 WHERE d.tenant = %s AND p.data_pedido BETWEEN %s AND %s
   AND d.tipo_cadastro IS NOT NULL
"""


def carregar_do_warehouse(conn, tenant: str, inicio: date, fim: date) -> dict:
    """Monta, a partir do banco, a mesma estrutura que `agregar` consome.

    pedidos: {id: {data, sit, valor, cliente, ecom_nome}}
    det:     {id: {nat, vprod, vtp, desc, frete, itens: [[sku, id, nome, qtd, vunit]]}}
             so para pedido com detalhe (tem_detalhe = 1)
    catalogo:{id_produto: {tipo, nome, cat}} -- produto sem tipo_cadastro (404
             na fase 3) fica de fora e aparece como "fora do catalogo".
    """
    f = rv.para_float
    pedidos, det = {}, {}
    for r in etl_db.consultar(conn, _SQL_PEDIDOS, (tenant, inicio, fim)):
        pid = str(r["id_pedido"])
        pedidos[pid] = {
            "data": r["data_pedido"].isoformat(),
            "sit": -1 if r["situacao"] is None else int(r["situacao"]),
            "valor": f(r["valor_total_pedido"]),
            "cliente": r["cliente_nome"] or "",
            "ecom_nome": r["ecommerce_nome"] or "",
        }
        if r["tem_detalhe"]:
            det[pid] = {"nat": r["natureza_nome"] or "",
                        "vprod": f(r["valor_total_produtos"]),
                        "vtp": f(r["valor_total_pedido"]),
                        "desc": f(r["valor_desconto"]),
                        "frete": f(r["valor_frete"]),
                        "itens": []}
    for r in etl_db.consultar(conn, _SQL_ITENS, (tenant, inicio, fim)):
        info = det.get(str(r["id_pedido"]))
        if info is not None:
            info["itens"].append([
                r["sku"] or "",
                "" if r["id_produto"] is None else str(r["id_produto"]),
                r["descricao_item"] or "", f(r["quantidade"]),
                f(r["valor_unitario"])])
    catalogo = {
        str(r["id_produto"]): {"tipo": r["tipo_cadastro"] or "",
                               "nome": r["descricao"] or "",
                               "cat": r["categoria_caminho"] or ""}
        for r in etl_db.consultar(conn, _SQL_CATALOGO, (tenant, inicio, fim))}
    return {"pedidos": pedidos, "det": det, "catalogo": catalogo}


# ------------------------------------------------------------------ #
#  Tabela de precos externa (join por SKU)                            #
# ------------------------------------------------------------------ #

def carregar_precos(caminho: Path, col_preco: str,
                    col_sku: str = "sku") -> tuple[dict, dict]:
    """{sku: preco} da planilha externa + diagnostico do join.

    SKU repetido: vence `product_status = ACTIVE` quando a coluna existe;
    senao, a primeira ocorrencia. Conflito de preco entre duplicados e
    reportado — nao silenciado.
    """
    df = pd.read_excel(caminho)
    if col_sku not in df.columns or col_preco not in df.columns:
        raise SystemExit(
            f"  A planilha {caminho.name} nao tem a(s) coluna(s) pedida(s). "
            f"Colunas disponiveis: {list(df.columns)}")
    df[col_sku] = df[col_sku].astype(str).str.strip()
    df = df[df[col_sku].notna() & (df[col_sku] != "") & (df[col_sku] != "nan")]

    tem_status = "product_status" in df.columns
    precos: dict[str, float] = {}
    nomes: dict[str, str] = {}
    dg = {"linhas": len(df), "dups": 0, "conflitos": [], "sem_preco": 0,
          "arquivo": caminho.name, "coluna": col_preco}
    for sku, g in df.groupby(col_sku, sort=False):
        vals = [rv.para_float(v) for v in g[col_preco] if pd.notna(v)]
        vals = [v for v in vals if v > 0]
        if not vals:
            dg["sem_preco"] += 1
            continue
        if len(g) > 1:
            dg["dups"] += 1
            if len(set(round(v, 2) for v in vals)) > 1:
                dg["conflitos"].append((sku, sorted(set(round(v, 2) for v in vals))))
        escolhida = g
        if tem_status:
            ativos = g[g["product_status"].astype(str).str.upper() == "ACTIVE"]
            if not ativos.empty:
                escolhida = ativos
        preco = rv.para_float(escolhida.iloc[0][col_preco])
        if preco <= 0:
            preco = vals[0]
        precos[sku] = preco
        if "produto" in escolhida.columns:
            nomes[sku] = str(escolhida.iloc[0]["produto"])
    dg["skus"] = len(precos)
    return precos, {"tabela": dg, "nomes": nomes}


# ------------------------------------------------------------------ #
#  Agregacao (offline)                                                #
# ------------------------------------------------------------------ #

def _novo_prod() -> dict:
    return {"nome": "", "tipo": "", "cat": "", "qtd": 0.0, "valor": 0.0,
            "pedidos": set(), "precos": Counter(), "dias": defaultdict(float)}


def agregar(tenant: str, cache: dict, inicio: date, fim: date, canal: str,
            situacoes: set, filtrar_natureza: bool,
            sem_brinde: bool = False) -> dict:
    """Aplica os filtros na ordem e agrega por SKU.

    Cada pedido descartado cai em EXATAMENTE um motivo, para o censo do
    Diagnostico fechar com o numero de pedidos do canal.
    """
    pedidos = cache.get("pedidos") or {}
    det = cache.get("det") or {}
    cat = cache.get("catalogo") or {}
    excluidos = rv.clientes_excluidos(tenant)
    alvo = rv.norm(canal)
    dias = dias_do_periodo(inicio, fim)

    ag = {
        "tenant": tenant, "canal": canal,
        "produtos": defaultdict(_novo_prod),
        "por_dia": {d.isoformat(): {"pedidos": 0, "unid": 0.0, "fat": 0.0,
                                    "skus": set()} for d in dias},
        "integracoes_janela": Counter(), "integracoes_casadas": Counter(),
        "sit_n": Counter(), "sit_v": Counter(),
        "desc_sit": {"n": 0, "v": 0.0},
        "transf": {"n": 0, "v": 0.0},
        "sem_detalhe": 0,
        "nat_fora_n": Counter(), "nat_fora_v": Counter(),
        "nat_dentro": Counter(),
        "integridade": {"n": 0, "div": 0, "delta": 0.0},
        "somas": {"bruto": 0.0, "vprod": 0.0, "desc": 0.0, "frete": 0.0,
                  "vtp": 0.0},
        "kits": Counter(), "mp": Counter(),
        # Brinde e SEMPRE medido; sem_brinde decide se e descartado ou somado.
        "sem_brinde": sem_brinde,
        "brinde": {"linhas": 0, "unid": 0, "skus": set()},
        "ped_so_brinde": 0,
        "sem_sku": 0, "sem_catalogo": set(),
        "n_periodo": 0, "n_canal": 0, "n_contados": 0, "linhas": 0,
        "fora_periodo": (cache.get("descartes_lista") or {}).get("fora do periodo", 0),
    }

    for pid, ped in pedidos.items():
        try:
            dt = datetime.strptime(ped["data"], "%Y-%m-%d").date()
        except (KeyError, ValueError):
            continue
        if not (inicio <= dt <= fim):
            continue
        ag["n_periodo"] += 1

        enome = (ped.get("ecom_nome") or "").strip()
        ag["integracoes_janela"][enome or "(vazio)"] += 1

        # (1) canal — antes do censo de situacao, para o censo ser DO CANAL
        if alvo and alvo not in rv.norm(enome):
            continue
        ag["n_canal"] += 1
        ag["integracoes_casadas"][enome or "(vazio)"] += 1

        sit = ped.get("sit", -1)
        valor = rv.para_float(ped.get("valor"))
        ag["sit_n"][sit] += 1
        ag["sit_v"][sit] += valor

        # (2) situacao
        if sit not in situacoes:
            ag["desc_sit"]["n"] += 1
            ag["desc_sit"]["v"] += valor
            continue

        # (3) transferencia interna do grupo (inclui a propria razao social)
        if rv.eh_excluido(ped.get("cliente", ""), excluidos):
            ag["transf"]["n"] += 1
            ag["transf"]["v"] += valor
            continue

        # (4) detalhe ausente — sem ele nao existem itens
        info = det.get(pid)
        if info is None:
            ag["sem_detalhe"] += 1
            continue

        # (5) natureza de operacao
        nat = info.get("nat") or ""
        if filtrar_natureza and not rv.perfil(tenant, "produtos_canal").eh_natureza_venda(nat):
            ag["nat_fora_n"][nat or "(sem natureza)"] += 1
            ag["nat_fora_v"][nat or "(sem natureza)"] += valor
            continue
        ag["nat_dentro"][nat or "(sem natureza)"] += 1

        # (6) contado
        ag["n_contados"] += 1
        chave_dia = dt.isoformat()
        bucket = ag["por_dia"].setdefault(
            chave_dia, {"pedidos": 0, "unid": 0.0, "fat": 0.0, "skus": set()})
        bucket["pedidos"] += 1

        bruto_ped = 0.0
        bucket_itens_antes = ag["linhas"]
        for l in (info.get("itens") or []):
            sku, prod_id, nome_item = l[0], l[1], l[2]
            qtd, vunit = rv.para_float(l[3]), rv.para_float(l[4])
            if not sku:
                ag["sem_sku"] += 1
                sku = prod_id or "(sem sku)"
            if vunit <= 0:
                ag["brinde"]["linhas"] += 1
                ag["brinde"]["unid"] += qtd
                ag["brinde"]["skus"].add(sku)
                if sem_brinde:
                    # Corte por LINHA: o SKU que tambem sai vendido sobrevive
                    # com as unidades pagas e o preco de verdade.
                    continue
            cinfo = cat.get(prod_id) if prod_id else None
            if prod_id and cinfo is None:
                ag["sem_catalogo"].add(prod_id)
            tipo = (cinfo or {}).get("tipo", "")
            nome = (cinfo or {}).get("nome") or nome_item
            categoria = (cinfo or {}).get("cat", "")

            p = ag["produtos"][sku]
            p["nome"] = nome or p["nome"]
            p["tipo"] = tipo or p["tipo"]
            p["cat"] = categoria or p["cat"]
            p["qtd"] += qtd
            p["valor"] += qtd * vunit
            p["pedidos"].add(pid)
            p["precos"][round(vunit, 2)] += qtd
            p["dias"][chave_dia] += qtd

            if tipo == "K":
                ag["kits"][sku] += qtd
            elif tipo == "M":
                ag["mp"][sku] += qtd

            ag["linhas"] += 1
            bruto_ped += qtd * vunit
            bucket["unid"] += qtd
            bucket["skus"].add(sku)

        if sem_brinde and bucket_itens_antes == ag["linhas"]:
            # Pedido 100% brinde: continua sendo pedido contado (existiu e foi
            # enviado), mas nao contribui com nenhuma linha de produto.
            ag["ped_so_brinde"] += 1
        bucket["fat"] += bruto_ped
        vprod = rv.para_float(info.get("vprod"))
        ag["somas"]["bruto"] += bruto_ped
        ag["somas"]["vprod"] += vprod
        ag["somas"]["desc"] += rv.para_float(info.get("desc"))
        ag["somas"]["frete"] += rv.para_float(info.get("frete"))
        ag["somas"]["vtp"] += rv.para_float(info.get("vtp")) or valor
        ag["integridade"]["n"] += 1
        if abs(bruto_ped - vprod) > 0.01:
            ag["integridade"]["div"] += 1
        ag["integridade"]["delta"] += bruto_ped - vprod

    return ag


def tabela_produtos(ag: dict, precos_tab: dict | None = None,
                    so_com_preco: bool = False) -> pd.DataFrame:
    """Uma linha por SKU, ordenada por quantidade desc (empate: faturamento).

    Com `precos_tab` (join por SKU), acrescenta o preco de tabela, o total
    teorico a esse preco e o delta contra o praticado. `% do faturamento` passa
    a ser a fatia do total DE TABELA — a base que foi pedida — e o faturado real
    continua na coluna ao lado.
    """
    usa_tab = bool(precos_tab)
    itens = list(ag["produtos"].items())
    if usa_tab and so_com_preco:
        itens = [(sku, p) for sku, p in itens if sku in precos_tab]

    total_fat = sum(p["valor"] for _, p in itens) or 1.0
    total_un = sum(p["qtd"] for _, p in itens) or 1.0
    total_tab = sum((precos_tab.get(sku, p["valor"] / p["qtd"] if p["qtd"] else 0.0)
                     * p["qtd"]) for sku, p in itens) if usa_tab else 0.0
    total_tab = total_tab or 1.0

    linhas = []
    for sku, p in itens:
        precos = [pr for pr in p["precos"] if pr > 0]
        medio = round(p["valor"] / p["qtd"], 2) if p["qtd"] else 0.0
        reg = {
            "SKU": sku,
            "Produto": p["nome"],
            "Tipo": p["tipo"],
        }
        if usa_tab:
            na_tab = sku in precos_tab
            preco_tab = round(precos_tab[sku], 2) if na_tab else medio
            total_t = round(preco_tab * p["qtd"], 2)
            reg.update({
                "Preco tabela": preco_tab,
                "Qtd vendida": p["qtd"],
                "Total a preco tabela R$": total_t,
                "% do faturamento": round(100.0 * total_t / total_tab, 2),
                "Fonte do preco": "planilha" if na_tab else "pedido (fallback)",
                "Preco praticado medio": medio,
                "Total faturado R$": round(p["valor"], 2),
                "Delta R$": round(total_t - p["valor"], 2),
                "Delta %": (round(100.0 * (total_t - p["valor"]) / p["valor"], 2)
                            if p["valor"] else 0.0),
            })
        else:
            reg.update({
                "Preco unit. medio": medio,
                "Qtd vendida": p["qtd"],
                "Total faturado R$": round(p["valor"], 2),
                "% do faturamento": round(100.0 * p["valor"] / total_fat, 2),
            })
        reg.update({
            "% das unidades": round(100.0 * p["qtd"] / total_un, 2),
            "Preco min": min(precos) if precos else 0.0,
            "Preco max": max(precos) if precos else 0.0,
            "Precos dist.": len(p["precos"]),
            "Pedidos": len(p["pedidos"]),
            "Categoria": p["cat"],
        })
        linhas.append(reg)

    df = pd.DataFrame(linhas)
    if df.empty:
        return df
    col_val = "Total a preco tabela R$" if usa_tab else "Total faturado R$"
    df = df.sort_values(["Qtd vendida", col_val],
                        ascending=[False, False]).reset_index(drop=True)
    df.insert(0, "Rank", range(1, len(df) + 1))
    return df


# ------------------------------------------------------------------ #
#  Saida — console                                                    #
# ------------------------------------------------------------------ #

def imprimir(ag: dict, df: pd.DataFrame, top: int, inicio: date, fim: date) -> None:
    nome_emp = rv.NOMES.get(ag["tenant"], ag["tenant"])
    print(f"\n{'=' * 118}")
    print(f"  {nome_emp} | canal: {ag['canal'] or '(todos)'} | "
          f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}")
    print(f"{'=' * 118}")
    unid = sum(p["qtd"] for p in ag["produtos"].values())
    print(f"  Pedidos no periodo (todas as integracoes): {ag['n_periodo']}")
    print(f"  Pedidos do canal                         : {ag['n_canal']}")
    print(f"  Pedidos contados (pos-filtros)           : {ag['n_contados']}")
    print(f"  Linhas / unidades / SKUs distintos       : "
          f"{ag['linhas']} / {unid:,.0f} / {len(ag['produtos'])}")
    print(f"  Faturamento bruto                        : "
          f"R$ {ag['somas']['bruto']:,.2f}")
    if df.empty:
        print("\n  Nenhum produto no filtro.")
        return
    # Com preco de tabela as colunas mudam: preco e total passam a ser os da
    # planilha, e o praticado vira coluna de comparacao ao lado.
    usa_tab = "Preco tabela" in df.columns
    col_pr = "Preco tabela" if usa_tab else "Preco unit. medio"
    col_tot = "Total a preco tabela R$" if usa_tab else "Total faturado R$"
    ln = 42 if usa_tab else 52
    n = min(top, len(df))
    print(f"\n  TOP {n} por quantidade vendida"
          f"{' — preco de tabela' if usa_tab else ''}:\n")
    cab = (f"  {'#':>3}  {'SKU':<14} {'Produto':<{ln}} "
           f"{('Pr.tab' if usa_tab else 'Pr.un'):>8} {'Qtd':>6} "
           f"{'Total R$':>12} {'%fat':>6}")
    sep = (f"  {'-' * 3}  {'-' * 14} {'-' * ln} {'-' * 8} {'-' * 6} "
           f"{'-' * 12} {'-' * 6}")
    if usa_tab:
        cab += f" {'Pr.prat':>8} {'Delta R$':>10} {'Fonte':>8}"
        sep += f" {'-' * 8} {'-' * 10} {'-' * 8}"
    print(cab)
    print(sep)
    for _, r in df.head(n).iterrows():
        linha = (f"  {int(r['Rank']):>3}  {str(r['SKU'])[:14]:<14} "
                 f"{str(r['Produto'])[:ln]:<{ln}} "
                 f"{r[col_pr]:>8,.2f} {r['Qtd vendida']:>6,.0f} "
                 f"{r[col_tot]:>12,.2f} {r['% do faturamento']:>6.2f}")
        if usa_tab:
            fonte = "planilha" if r["Fonte do preco"] == "planilha" else "PEDIDO"
            linha += (f" {r['Preco praticado medio']:>8,.2f} "
                      f"{r['Delta R$']:>10,.2f} {fonte:>8}")
        print(linha)
    sub = df.head(n)
    print(sep)
    print(f"  {'':>3}  {'TOP ' + str(n):<14} {'':<{ln}} {'':>8} "
          f"{sub['Qtd vendida'].sum():>6,.0f} {sub[col_tot].sum():>12,.2f} "
          f"{sub['% do faturamento'].sum():>6.2f}")
    print(f"  {'':>3}  {'TOTAL do canal':<14} {'':<{ln}} {'':>8} "
          f"{df['Qtd vendida'].sum():>6,.0f} {df[col_tot].sum():>12,.2f} "
          f"{100.0:>6.2f}")
    if usa_tab:
        semp = df[df["Fonte do preco"] != "planilha"]
        un = df["Qtd vendida"].sum() or 1.0
        real = df["Total faturado R$"].sum() or 1.0
        print(f"\n  Join por SKU: {len(df) - len(semp)} de {len(df)} SKUs com "
              f"preco na planilha. {len(semp)} sem preco "
              f"({semp['Qtd vendida'].sum():,.0f} un, "
              f"{100.0 * semp['Qtd vendida'].sum() / un:.1f}% das unidades) "
              f"entraram pelo preco praticado.")
        print(f"  Total a preco de tabela R$ {df[col_tot].sum():,.2f} | "
              f"faturado real R$ {real:,.2f} | delta R$ "
              f"{df['Delta R$'].sum():,.2f} "
              f"({100.0 * df['Delta R$'].sum() / real:+.1f}%)")


def imprimir_diagnostico(ag: dict, situacoes: set, filtrar_natureza: bool) -> None:
    print("\n  --- Diagnostico ---")
    print("  Integracoes na janela (o filtro casou com as marcadas com *):")
    for nome, n in ag["integracoes_janela"].most_common():
        marca = "*" if nome in ag["integracoes_casadas"] else " "
        print(f"    {marca} {n:5d}  {nome}")
    print("  Censo de situacao DO CANAL (antes do filtro de situacao):")
    for s in sorted(ag["sit_n"]):
        dentro = "OK " if s in situacoes else "OUT"
        print(f"    [{dentro}] {s} {rv.SITUACOES.get(s, '?'):<18} "
              f"{ag['sit_n'][s]:5d} ped  R$ {ag['sit_v'][s]:12,.2f}")
    print(f"  Descartes: situacao {ag['desc_sit']['n']} "
          f"(R$ {ag['desc_sit']['v']:,.2f}) | transferencia interna "
          f"{ag['transf']['n']} (R$ {ag['transf']['v']:,.2f}) | "
          f"sem detalhe {ag['sem_detalhe']}")
    if filtrar_natureza and ag["nat_fora_n"]:
        print("  Natureza fora do alvo:")
        for nat, n in ag["nat_fora_n"].most_common():
            print(f"      {n:5d}  {nat}  (R$ {ag['nat_fora_v'][nat]:,.2f})")
    else:
        print("  Natureza fora do alvo: 0 pedido")
    print("  Naturezas contadas: " +
          "; ".join(f"{k} ({v})" for k, v in ag["nat_dentro"].most_common()))
    censo = (ag["desc_sit"]["n"] + ag["transf"]["n"] + ag["sem_detalhe"]
             + sum(ag["nat_fora_n"].values()) + ag["n_contados"])
    print(f"  Censo fecha? canal={ag['n_canal']} vs contados+descartes={censo} "
          f"-> {'OK' if censo == ag['n_canal'] else 'DIVERGENTE'}")
    i = ag["integridade"]
    print(f"  Integridade soma(qtd x valorUnitario) vs valorTotalProdutos: "
          f"{i['div']} divergencia(s) em {i['n']} pedidos, delta acumulado "
          f"R$ {i['delta']:,.2f}")
    print(f"  Bruto R$ {ag['somas']['bruto']:,.2f} | vlrProdutos R$ "
          f"{ag['somas']['vprod']:,.2f} | desconto R$ {ag['somas']['desc']:,.2f} | "
          f"frete R$ {ag['somas']['frete']:,.2f} | total pedido R$ "
          f"{ag['somas']['vtp']:,.2f}")
    if ag["kits"]:
        print(f"  [ATENCAO] {len(ag['kits'])} SKU(s) tipo=K (kit) como linha de "
              f"pedido — {sum(ag['kits'].values()):,.0f} un: " +
              "; ".join(f"{k} ({v:,.0f})" for k, v in ag["kits"].most_common()))
    else:
        print("  Kits (tipo=K) como linha de pedido: 0 (esperado)")
    br = ag["brinde"]
    print(f"  Brinde (valorUnitario = 0): {br['linhas']} linha(s) | "
          f"{br['unid']:,.0f} un | {len(br['skus'])} SKU(s) -> "
          f"{'DESCARTADO (--sem-brinde)' if ag['sem_brinde'] else 'incluido no ranking'}")
    if ag["sem_brinde"] and ag["ped_so_brinde"]:
        print(f"  [ATENCAO] {ag['ped_so_brinde']} pedido(s) ficaram sem nenhuma "
              f"linha (eram 100% brinde) — seguem contados como pedido.")
    if ag["mp"]:
        print(f"  Materia-prima (tipo=M) vendida: {len(ag['mp'])} SKU(s), "
              f"{sum(ag['mp'].values()):,.0f} un")
    if ag["sem_sku"]:
        print(f"  [ATENCAO] {ag['sem_sku']} linha(s) sem SKU (usou o id do produto)")
    if ag["sem_catalogo"]:
        print(f"  [ATENCAO] {len(ag['sem_catalogo'])} produto(s) fora do catalogo "
              f"(fase 3 incompleta -> coluna Tipo vazia)")
    print("  Por dia:")
    for d, b in sorted(ag["por_dia"].items()):
        dt = datetime.strptime(d, "%Y-%m-%d").date()
        print(f"      {rotulo_dia(dt)}: {b['pedidos']:4d} ped | "
              f"{b['unid']:8,.0f} un | R$ {b['fat']:11,.2f} | "
              f"{len(b['skus']):4d} SKUs")


# ------------------------------------------------------------------ #
#  Saida — Excel                                                      #
# ------------------------------------------------------------------ #

_AZUL = PatternFill("solid", fgColor="1F4E78")

_LARG = {"Rank": 6, "SKU": 18, "Produto": 58, "Tipo": 6,
         "Preco unit. medio": 16, "Qtd vendida": 12, "Total faturado R$": 18,
         "% do faturamento": 16, "% das unidades": 14, "Preco min": 12,
         "Preco max": 12, "Precos dist.": 12, "Pedidos": 10, "Categoria": 40,
         "Preco tabela": 14, "Total a preco tabela R$": 20,
         "Fonte do preco": 18, "Preco praticado medio": 18,
         "Delta R$": 14, "Delta %": 11}


def _formatar(ws, df: pd.DataFrame, larguras: dict) -> None:
    for c in range(1, len(df.columns) + 1):
        cel = ws.cell(row=1, column=c)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = _AZUL
        cel.alignment = Alignment(horizontal="center", vertical="center",
                                  wrap_text=True)
    for i, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = larguras.get(col, 14)
        fmt = None
        if col in ("Total faturado R$", "Preco unit. medio", "Preco min",
                   "Preco max", "Faturamento R$", "Preco tabela",
                   "Total a preco tabela R$", "Preco praticado medio",
                   "Delta R$"):
            fmt = 'R$ #,##0.00'
        elif col in ("Qtd vendida", "Unidades"):
            fmt = '#,##0'
        elif col in ("% do faturamento", "% das unidades", "Delta %"):
            fmt = '0.00"%"'
        if fmt:
            for r in range(2, len(df) + 2):
                ws.cell(row=r, column=i).number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def gerar_excel(ag: dict, df: pd.DataFrame, top: int, inicio: date, fim: date,
                situacoes: set, filtrar_natureza: bool, caminho: Path,
                info_precos: dict | None = None) -> None:
    nome_emp = rv.NOMES.get(ag["tenant"], ag["tenant"])
    with pd.ExcelWriter(caminho, engine="openpyxl") as xl:
        n = min(top, len(df))
        df_top = df.head(n)
        aba_top = f"Top {n}"
        df_top.to_excel(xl, sheet_name=aba_top, index=False)
        _formatar(xl.sheets[aba_top], df_top, _LARG)

        # Segunda ordenacao, por faturamento: no site uma parte grande das
        # unidades e brinde a R$ 0,00 (sache, tigela, pincel), entao o ranking
        # por quantidade e o ranking por dinheiro contam historias diferentes.
        col_tot = ("Total a preco tabela R$" if "Preco tabela" in df.columns
                   else "Total faturado R$")
        df_fat = (df.drop(columns=["Rank"])
                    .sort_values([col_tot, "Qtd vendida"],
                                 ascending=[False, False])
                    .reset_index(drop=True))
        df_fat.insert(0, "Rank", range(1, len(df_fat) + 1))
        df_fat_top = df_fat.head(n)
        aba_fat = f"Top {n} faturamento"
        df_fat_top.to_excel(xl, sheet_name=aba_fat, index=False)
        _formatar(xl.sheets[aba_fat], df_fat_top, _LARG)

        df.to_excel(xl, sheet_name="Produtos (todos)", index=False)
        _formatar(xl.sheets["Produtos (todos)"], df, _LARG)

        # SKU sem preco na planilha nao desaparece: vira aba propria, para a
        # decisao de completar a tabela de precos ser tomada com a lista na mao.
        if "Fonte do preco" in df.columns:
            semp = df[df["Fonte do preco"] != "planilha"].copy()
            if not semp.empty:
                semp.to_excel(xl, sheet_name="Sem preco na planilha",
                              index=False)
                _formatar(xl.sheets["Sem preco na planilha"], semp, _LARG)

        linhas_dia = []
        for d, b in sorted(ag["por_dia"].items()):
            dt = datetime.strptime(d, "%Y-%m-%d").date()
            linhas_dia.append({"Dia": rotulo_dia(dt), "Data": d,
                               "Pedidos": b["pedidos"], "Unidades": b["unid"],
                               "Faturamento R$": round(b["fat"], 2),
                               "SKUs distintos": len(b["skus"])})
        if linhas_dia:
            ddia = pd.DataFrame(linhas_dia)
            ddia.to_excel(xl, sheet_name="Por dia", index=False)
            _formatar(xl.sheets["Por dia"], ddia,
                      {"Dia": 14, "Data": 12, "Pedidos": 10, "Unidades": 12,
                       "Faturamento R$": 18, "SKUs distintos": 14})

        fl = [
            ("Empresa (tenant)", f"{nome_emp} ({ag['tenant']})"),
            ("Canal / integracao", ag["canal"] or "(todos)"),
            ("Match do canal", "substring normalizada em ecommerce.nome "
                               "(minusculas, sem acento)"),
            ("Integracoes que casaram",
             "; ".join(f"{k} ({v})"
                       for k, v in ag["integracoes_casadas"].most_common())),
            ("Periodo", f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y} "
                        f"({len(dias_do_periodo(inicio, fim))} dias)"),
            ("Situacoes consideradas",
             ", ".join(f"{s}={rv.SITUACOES.get(s, '?')}"
                       for s in sorted(situacoes))),
            ("Situacoes excluidas",
             ", ".join(f"{s}={rv.SITUACOES.get(s, '?')}"
                       for s in sorted(set(rv.SITUACOES) - situacoes))),
            ("Natureza de operacao",
             "so naturezas de venda (" + rv.perfil(ag["tenant"], "produtos_canal")
             .descrever_naturezas() + ")" if filtrar_natureza
             else "SEM FILTRO"),
            ("Transferencia interna",
             "excluida (clientes_excluidos + razao_social do proprio tenant)"),
            ("Kits", "NAO explodidos — o Tiny entrega o kit destrinchado em "
                     "componentes; linha de kit fica rotulada Tipo=K"),
            ("Valor", "BRUTO. 'Total faturado R$' = soma(quantidade x "
                      "valorUnitario). NAO existe desconto em nivel de item na "
                      "API; desconto e frete sao do pedido inteiro e nao foram "
                      "rateados."),
            ("Preco unit. medio", "media PONDERADA = Total faturado / Qtd. "
                                  "Ver Preco min/max e Precos dist. para a "
                                  "dispersao."),
            ("Brinde (valorUnitario = 0)",
             (f"DESCARTADO (--sem-brinde): {ag['brinde']['linhas']} linhas, "
              f"{ag['brinde']['unid']:,.0f} un, {len(ag['brinde']['skus'])} SKUs "
              f"fora do ranking. Corte por LINHA, nao por SKU: o SKU que tambem "
              f"sai vendido fica com as unidades pagas e o preco real. O "
              f"faturamento do canal NAO muda — linha de valor zero soma zero."
              if ag["sem_brinde"] else
              f"INCLUIDO: {ag['brinde']['linhas']} linhas, "
              f"{ag['brinde']['unid']:,.0f} un, {len(ag['brinde']['skus'])} SKUs "
              f"a R$ 0,00 contam como unidade vendida com receita zero.")),
            ("Ordenacao", "quantidade vendida desc (empate: faturamento desc)"),
            ("Fonte", "warehouse tiny_dw (pedido, pedido_item, dim_produto)"),
        ]
        if info_precos and "Fonte do preco" in df.columns:
            t = info_precos["tabela"]
            semp = df[df["Fonte do preco"] != "planilha"]
            un = df["Qtd vendida"].sum() or 1.0
            real = df["Total faturado R$"].sum() or 1.0
            fl += [
                ("PRECO DE TABELA — arquivo", t["arquivo"]),
                ("PRECO DE TABELA — coluna", t["coluna"]),
                ("Join", f"por SKU (codigo exato). {t['skus']} SKUs com preco "
                         f"na planilha, de {t['linhas']} linhas; "
                         f"{t['dups']} SKU(s) repetido(s) resolvido(s) por "
                         f"product_status=ACTIVE"
                         + (f"; CONFLITO de preco em {len(t['conflitos'])} "
                            f"SKU(s): {t['conflitos']}"
                            if t["conflitos"] else " (sem conflito de preco)")),
                ("Cobertura do join",
                 f"{len(df) - len(semp)} de {len(df)} SKUs do ranking casaram. "
                 f"{len(semp)} sem preco ({semp['Qtd vendida'].sum():,.0f} un, "
                 f"{100.0 * semp['Qtd vendida'].sum() / un:.1f}% das unidades) "
                 f"entraram pelo PRECO PRATICADO (Fonte do preco = 'pedido "
                 f"(fallback)') e estao na aba 'Sem preco na planilha'. "
                 f"--so-com-preco descarta esses SKUs."),
                ("Total a preco tabela R$",
                 f"{df['Total a preco tabela R$'].sum():,.2f} — valor TEORICO "
                 f"(quantidade x preco de tabela), NAO o faturado"),
                ("Total faturado real R$",
                 f"{real:,.2f} — soma(quantidade x valorUnitario do pedido). "
                 f"Delta R$ {df['Delta R$'].sum():,.2f} "
                 f"({100.0 * df['Delta R$'].sum() / real:+.1f}%)"),
                ("Abas Por dia / Diagnostico",
                 "seguem sempre com o PRECO PRATICADO (faturamento real); o "
                 "preco de tabela vale so nas abas de produto"),
                ("% do faturamento",
                 "fatia do total A PRECO DE TABELA (a base pedida), nao do "
                 "faturado real"),
            ]
        fl += [
            ("Gerado em", f"{datetime.now():%d/%m/%Y %H:%M:%S}"),
        ]
        dfl = pd.DataFrame(fl, columns=["Filtro", "Valor"])
        dfl.to_excel(xl, sheet_name="Filtros", index=False)
        ws = xl.sheets["Filtros"]
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 96
        for c in (1, 2):
            cel = ws.cell(row=1, column=c)
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = _AZUL
        for r in range(2, len(dfl) + 2):
            ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True,
                                                           vertical="top")

        i = ag["integridade"]
        censo = (ag["desc_sit"]["n"] + ag["transf"]["n"] + ag["sem_detalhe"]
                 + sum(ag["nat_fora_n"].values()) + ag["n_contados"])
        dg = [("1. Cobertura", "Pedidos no periodo (todas as integracoes)",
               ag["n_periodo"]),
              ("1. Cobertura", "Pedidos do canal", ag["n_canal"]),
              ("1. Cobertura", "Pedidos contados", ag["n_contados"]),
              ("1. Cobertura", "Pedidos sem detalhe (deve ser 0)",
               ag["sem_detalhe"]),
              ("1. Cobertura", "Censo fecha (canal = contados + descartes)",
               "OK" if censo == ag["n_canal"] else f"DIVERGENTE ({censo})"),
              ("1. Cobertura", "Fora do periodo na listagem",
               ag["fora_periodo"])]
        for nome, n in ag["integracoes_janela"].most_common():
            marca = " (CASOU)" if nome in ag["integracoes_casadas"] else ""
            dg.append(("2. Integracoes na janela", f"{nome}{marca}", n))
        for s in sorted(ag["sit_n"]):
            dg.append(("3. Situacao do canal (censo)",
                       f"{s} {rv.SITUACOES.get(s, '?')}"
                       f"{'' if s in situacoes else ' [EXCLUIDA]'}",
                       f"{ag['sit_n'][s]} ped | R$ {ag['sit_v'][s]:,.2f}"))
        dg += [("4. Descartes", "Situacao fora do conjunto",
                f"{ag['desc_sit']['n']} ped | R$ {ag['desc_sit']['v']:,.2f}"),
               ("4. Descartes", "Transferencia interna (grupo)",
                f"{ag['transf']['n']} ped | R$ {ag['transf']['v']:,.2f}")]
        for nat, n in ag["nat_fora_n"].most_common():
            dg.append(("4. Descartes", f"Natureza fora do alvo: {nat}",
                       f"{n} ped | R$ {ag['nat_fora_v'][nat]:,.2f}"))
        for nat, n in ag["nat_dentro"].most_common():
            dg.append(("5. Naturezas contadas", nat, f"{n} ped"))
        dg += [("6. Integridade",
                "Pedidos com soma(qtd x valorUnitario) != valorTotalProdutos",
                f"{i['div']} de {i['n']}"),
               ("6. Integridade", "Delta acumulado", f"R$ {i['delta']:,.2f}"),
               ("6. Integridade", "Faturamento bruto (soma dos itens)",
                f"R$ {ag['somas']['bruto']:,.2f}"),
               ("6. Integridade", "valorTotalProdutos somado",
                f"R$ {ag['somas']['vprod']:,.2f}"),
               ("6. Integridade", "Desconto do pedido (NAO rateado)",
                f"R$ {ag['somas']['desc']:,.2f}"),
               ("6. Integridade", "Frete do pedido (NAO rateado)",
                f"R$ {ag['somas']['frete']:,.2f}"),
               ("6. Integridade", "valorTotalPedido somado",
                f"R$ {ag['somas']['vtp']:,.2f}"),
               ("7. Kits", "SKUs tipo=K como linha de pedido",
                f"{len(ag['kits'])} SKU(s) | {sum(ag['kits'].values()):,.0f} un"
                if ag["kits"] else "0 (esperado)")]
        for k, v in ag["kits"].most_common():
            dg.append(("7. Kits", f"kit {k}", f"{v:,.0f} un"))
        dg += [("8. Brinde (valorUnitario = 0)",
                "Tratamento",
                "DESCARTADO (--sem-brinde)" if ag["sem_brinde"]
                else "incluido no ranking"),
               ("8. Brinde (valorUnitario = 0)", "Linhas / unidades / SKUs",
                f"{ag['brinde']['linhas']} linhas | "
                f"{ag['brinde']['unid']:,.0f} un | "
                f"{len(ag['brinde']['skus'])} SKUs"),
               ("8. Brinde (valorUnitario = 0)",
                "Pedidos que ficaram sem nenhuma linha (100% brinde)",
                ag["ped_so_brinde"] if ag["sem_brinde"] else "n/a"),
               ("8. Materia-prima", "SKUs tipo=M vendidos",
                f"{len(ag['mp'])} SKU(s) | {sum(ag['mp'].values()):,.0f} un"),
               ("9. Qualidade", "Linhas sem SKU", ag["sem_sku"]),
               ("9. Qualidade", "Produtos fora do catalogo (fase 3)",
                len(ag["sem_catalogo"]))]
        dfd = pd.DataFrame(dg, columns=["Secao", "Item", "Valor"])
        dfd.to_excel(xl, sheet_name="Diagnostico", index=False)
        ws = xl.sheets["Diagnostico"]
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 62
        ws.column_dimensions["C"].width = 32
        for c in (1, 2, 3):
            cel = ws.cell(row=1, column=c)
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = _AZUL
        ws.freeze_panes = "A2"
    print(f"\n  Planilha: {caminho}")


# ------------------------------------------------------------------ #
#  main                                                               #
# ------------------------------------------------------------------ #

def main():
    ap = argparse.ArgumentParser(
        description="Produtos mais vendidos por canal/integracao (le o "
                    "warehouse tiny_dw).")
    ap.add_argument("tenant", help="tenant_id de data/tenants.json")
    ap.add_argument("--inicio", type=rv.parse_data, required=True,
                    help="data inicial YYYY-MM-DD ou DD/MM/YYYY")
    ap.add_argument("--fim", type=rv.parse_data, required=True,
                    help="data final YYYY-MM-DD ou DD/MM/YYYY")
    ap.add_argument("--canal", required=True,
                    help="substring do ecommerce.nome; '' = todos os canais")
    ap.add_argument("--top", type=int, default=50, help="tamanho do ranking")
    ap.add_argument("--situacoes", default=None,
                    help="sobrescreve o conjunto, ex.: 1,3,4,5,6,7")
    ap.add_argument("--sem-filtro-natureza", action="store_true",
                    help="nao filtra por natureza de operacao")
    ap.add_argument("--sem-brinde", action="store_true",
                    help="descarta as linhas a valorUnitario 0 (corte por "
                         "linha, nao por SKU)")
    ap.add_argument("--precos-xlsx", default=None,
                    help="planilha externa de precos; o preco unitario passa a "
                         "vir dela, casado por SKU")
    ap.add_argument("--precos-coluna", default="preco",
                    help="coluna de preco na planilha externa")
    ap.add_argument("--precos-sku-coluna", default="sku",
                    help="coluna de SKU na planilha externa")
    ap.add_argument("--so-com-preco", action="store_true",
                    help="descarta os SKUs sem preco na planilha externa "
                         "(padrao: entram pelo preco praticado, sinalizados)")
    ap.add_argument("--sem-excel", action="store_true",
                    help="so imprime no console")
    args = ap.parse_args()

    tenant = args.tenant.lower()
    if tenant not in rv.NOMES:
        print(f"  Tenant desconhecido: {tenant!r}. Use: {list(rv.NOMES)}")
        sys.exit(1)
    if args.inicio > args.fim:
        print("  --inicio nao pode ser depois de --fim.")
        sys.exit(1)
    situacoes = (set(int(x) for x in args.situacoes.replace(" ", "").split(",") if x)
                 if args.situacoes
                 else set(rv.perfil(tenant, "produtos_canal").situacoes_validas))
    filtrar_natureza = not args.sem_filtro_natureza

    cfg = etl_db.carregar_config()
    conn = etl_db.conectar(cfg)
    try:
        cache = carregar_do_warehouse(conn, tenant, args.inicio, args.fim)
    finally:
        conn.close()
    if not cache["pedidos"]:
        print(f"  Nenhum pedido de {tenant} entre {args.inicio} e {args.fim} "
              f"no warehouse.")
        print(f"  Carregue a janela: pipeline_pedidos.py {tenant} "
              f"--inicio {args.inicio} --fim {args.fim}")
        sys.exit(1)

    print("=== Produtos por canal — warehouse ===")
    print(f"  Banco    : {etl_db.descrever(cfg)}")
    print(f"  Periodo  : {args.inicio:%d/%m/%Y} a {args.fim:%d/%m/%Y}")
    print(f"  Canal    : {args.canal or '(todos)'}")
    print(f"  Situacoes: {sorted(situacoes)}")
    print(f"  Natureza : "
          f"{'filtrada (so vendas)' if filtrar_natureza else 'SEM FILTRO'}")
    print(f"  Brinde   : "
          f"{'DESCARTADO (linhas a R$ 0,00)' if args.sem_brinde else 'incluido'}")

    precos_tab, info_precos = None, None
    if args.precos_xlsx:
        arq = Path(args.precos_xlsx)
        if not arq.is_absolute():
            arq = (_bootstrap.RAIZ / args.precos_xlsx)
        if not arq.exists():
            print(f"  Planilha de precos nao encontrada: {arq}")
            sys.exit(1)
        precos_tab, info_precos = carregar_precos(
            arq, args.precos_coluna, args.precos_sku_coluna)
        t = info_precos["tabela"]
        print(f"  Precos   : {t['arquivo']} coluna '{t['coluna']}' -> "
              f"{t['skus']} SKUs ({t['linhas']} linhas, {t['dups']} SKU(s) "
              f"repetido(s)"
              + (f", {len(t['conflitos'])} CONFLITO(S) de preco"
                 if t["conflitos"] else ", sem conflito de preco") + ")")
        if t["conflitos"]:
            for sku, precos in t["conflitos"][:10]:
                print(f"      [CONFLITO] {sku}: {precos}")
        print(f"  SKU sem preco na planilha: "
              f"{'DESCARTADO (--so-com-preco)' if args.so_com_preco else 'entra pelo preco praticado, sinalizado'}")

    ag = agregar(tenant, cache, args.inicio, args.fim, args.canal, situacoes,
                 filtrar_natureza, args.sem_brinde)
    df = tabela_produtos(ag, precos_tab, args.so_com_preco)
    imprimir(ag, df, args.top, args.inicio, args.fim)
    imprimir_diagnostico(ag, situacoes, filtrar_natureza)

    if not args.sem_excel and not df.empty:
        slug = rv.norm(args.canal).replace(" ", "_") or "todos_canais"
        if args.sem_brinde:
            slug += "_sem_brinde"
        if precos_tab:
            slug += "_preco_tabela"
        nome = (f"produtos_canal_{slug}_{tenant}_{args.inicio:%Y%m%d}_a_"
                f"{args.fim:%Y%m%d}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        caminho = _REPORTS_DIR / nome
        caminho.parent.mkdir(parents=True, exist_ok=True)
        try:
            gerar_excel(ag, df, args.top, args.inicio, args.fim, situacoes,
                        filtrar_natureza, caminho, info_precos)
        except PermissionError:
            alt = caminho.with_name(f"{caminho.stem}_b.xlsx")
            print(f"  arquivo bloqueado; salvando como {alt.name}")
            gerar_excel(ag, df, args.top, args.inicio, args.fim, situacoes,
                        filtrar_natureza, alt, info_precos)

    print("\n=== Concluido ===")


if __name__ == "__main__":
    main()
