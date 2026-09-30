"""
Relatorio de vendas por item e clientes por LINHA DE PRODUTO -- lido 100% do
warehouse tiny_dw (nenhum request a API).

As linhas vem de linhas_produto.json (TINY_APPS_CONFIG ou data/; modelo em
data/linhas_produto.exemplo.json): cada uma e reconhecida pelo prefixo de
`categoria_caminho` do cadastro, mais SKUs extras que o cadastro classifica
fora dela e menos SKUs excluidos (acessorio, kit que mistura linhas...).

Saida em reports/pipeline_etl/ (CSV `;` + decimal `,` + UTF-8 com BOM, para
abrir direto no Excel pt-BR):

  relatorio_linhas_<ini>_<fim>.xlsx                TUDO abaixo, uma aba por
                                                   tabela + aba Notas (numero
                                                   gravado como numero)
  resumo_mensal_<ini>_<fim>.csv                    totais por mes e empresa
  vendas_itens_mensal_consolidado_<ini>_<fim>.csv  mes x SKU, empresas somadas
  vendas_itens_mensal_<ini>_<fim>.csv              empresa x mes x SKU
  resumo_clientes_linhas_<ini>_<fim>.csv           quantos clientes por linha
  clientes_<chave>_<ini>_<fim>.csv                 um arquivo por linha,
                                                   um cliente por linha

DECISOES:
  - Base = v_venda_item: situacao valida, natureza de venda, sem transferencia
    entre as empresas do grupo. Nenhuma regra e reescrita aqui.
  - Faturamento = soma de qtd x valorUnitario (BRUTO -- a API nao tem desconto
    por item; desconto de cabecalho e frete ficam de fora). Preco medio =
    faturamento / quantidade, com brinde (R$ 0,00) incluido na quantidade, como
    no default da view; a coluna preco_medio_sem_brinde tira o brinde.
  - As tabelas de vendas por item trazem TODOS os SKUs vendidos; o recorte por
    linha vale para as listas de clientes.
  - "Comprou" = linha com valor_unitario > 0 e quantidade > 0. Quem so recebeu
    sache/mini de brinde da linha nao conta como comprador.
  - Cliente e unificado entre as empresas pelo CPF/CNPJ (so digitos); sem
    documento, a chave e empresa + id do contato. Contato vem do cadastro
    (dim_cliente), com fallback para o snapshot gravado no pedido.
  - O relatorio usa SO o banco: venda que nao passou pelo Tiny nao aparece.

Uso:
    python relatorio_linhas.py                                  # 01/01 do ano ate o fim do mes passado
    python relatorio_linhas.py --inicio 2026-01-01 --fim 2026-08-31
    python relatorio_linhas.py --config outra_config.json
"""

import argparse
import csv
import io
import json
import re
import sys
from collections import defaultdict
from datetime import date
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import _bootstrap
import etl_db
from core.auth_tiny import get_all_tenant_ids

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SAIDA_DIR = _bootstrap.REPORTS_DIR / "pipeline_etl"

from core import config as _config  # noqa: E402

CONFIG_PADRAO = _config.caminho("linhas_produto.json")


def carregar_linhas(caminho) -> tuple[dict, list]:
    """{chave: (rotulo, prefixo de categoria, SKUs extras, SKUs excluidos)} + notas."""
    if not caminho.exists():
        raise SystemExit(f"  [ERRO] {caminho} nao existe. Copie "
                         f"data/linhas_produto.exemplo.json e ajuste as linhas.")
    cfg = json.loads(caminho.read_text(encoding="utf-8"))
    linhas = {l["chave"]: (l["rotulo"], l["categoria"],
                           tuple(l.get("skus_extras", ())),
                           tuple(l.get("skus_excluidos", ())))
              for l in cfg["linhas"]}
    return linhas, list(cfg.get("notas", []))


def _lista_sql(skus) -> str:
    return ",".join("'" + str(s).replace("'", "''") + "'" for s in skus)


def _sql_linha(linhas: dict) -> str:
    """CASE que resolve a linha de cada item (NULL = fora de todas)."""
    partes = []
    for chave, (_, prefixo, extras, excluidos) in linhas.items():
        prefixo = prefixo.replace("'", "''").replace("%", "%%")
        cond = f"categoria_caminho LIKE '{prefixo}%%'"  # %% = % literal no PyMySQL
        if extras:
            cond += f" OR sku IN ({_lista_sql(extras)})"
        if excluidos:
            cond = f"({cond}) AND sku NOT IN ({_lista_sql(excluidos)})"
        partes.append(f"WHEN {cond} THEN '{chave}'")
    return "CASE " + " ".join(partes) + " END"


SQL_VENDAS = """
SELECT tenant, ano_mes, sku,
       MAX(produto) AS produto,
       MAX(categoria_caminho) AS categoria,
       SUM(quantidade) AS quantidade,
       SUM(CASE WHEN eh_brinde = 1 THEN quantidade ELSE 0 END) AS qtd_brinde,
       SUM(valor_bruto) AS faturamento,
       COUNT(DISTINCT id_pedido) AS pedidos,
       MIN(NULLIF(valor_unitario, 0)) AS preco_min,
       MAX(valor_unitario) AS preco_max
FROM v_venda_item
WHERE data_pedido BETWEEN %s AND %s AND qtd_zero = 0
GROUP BY tenant, ano_mes, sku
"""

SQL_CLIENTES = """
SELECT v.tenant, v.id_cliente, v.id_pedido, v.data_pedido, v.sku,
       v.quantidade, v.valor_bruto, v.ecommerce_nome,
       {caso_linha} AS linha,
       COALESCE(NULLIF(c.nome, ''), p.cliente_nome) AS nome,
       COALESCE(NULLIF(c.cpf_cnpj, ''), p.cliente_cpf_cnpj) AS cpf_cnpj,
       COALESCE(NULLIF(c.email, ''), p.cliente_email) AS email,
       COALESCE(NULLIF(c.telefone, ''), NULLIF(c.celular, ''),
                NULLIF(p.cliente_telefone, ''), p.cliente_celular) AS telefone,
       COALESCE(NULLIF(c.municipio, ''), p.cliente_municipio) AS municipio,
       COALESCE(NULLIF(c.uf, ''), p.cliente_uf) AS uf
FROM v_venda_item v
JOIN pedido p ON p.tenant = v.tenant AND p.id_pedido = v.id_pedido
LEFT JOIN dim_cliente c ON c.tenant = v.tenant AND c.id_cliente = v.id_cliente
WHERE v.data_pedido BETWEEN %s AND %s
  AND v.qtd_zero = 0 AND v.eh_brinde = 0
HAVING linha IS NOT NULL
"""

DOIS = Decimal("0.01")


def _r2(v):
    return None if v is None else Decimal(v).quantize(DOIS)


def _qtd(v):
    d = Decimal(v)
    return int(d) if d == d.to_integral_value() else d


def _celula_csv(v) -> str:
    """Numero em pt-BR sem separador de milhar (o Excel le como numero)."""
    if v is None:
        return ""
    if isinstance(v, Decimal):
        return str(v).replace(".", ",")
    if isinstance(v, date):
        return v.isoformat()
    return str(v)


def _formato(coluna: str):
    if coluna in ("faturamento", "ticket_medio") or coluna.startswith("preco"):
        return "#,##0.00"
    if coluna in ("quantidade", "qtd_brinde", "pedidos", "unidades", "clientes",
                  "skus") or coluna.startswith(("so_", "nas_", "com_")):
        return "#,##0"
    if coluna.endswith("_compra"):
        return "dd/mm/yyyy"
    return None


class Saida:
    """Acumula as tabelas: cada uma vira um CSV na hora e uma aba no xlsx no fim.
    As linhas guardam Decimal/int/date crus; cada formato formata na escrita."""

    def __init__(self, sufixo: str):
        self.sufixo = sufixo
        self.abas = []  # (nome_aba, cabecalho, linhas)

    def tabela(self, arquivo: str, aba: str, cab: list, linhas: list) -> None:
        caminho = SAIDA_DIR / f"{arquivo}_{self.sufixo}.csv"
        with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(cab)
            w.writerows([_celula_csv(v) for v in l] for l in linhas)
        print(f"  {caminho.relative_to(_bootstrap.RAIZ)}  ({len(linhas)} linhas)")
        self.abas.append((aba, cab, linhas))

    def xlsx(self, notas: list) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Notas"
        for texto in notas:
            ws.append([texto])
        ws.column_dimensions["A"].width = 120
        ws.cell(1, 1).font = Font(bold=True, size=12)
        for aba, cab, linhas in self.abas:
            ws = wb.create_sheet(aba[:31])
            ws.append(cab)
            for l in linhas:
                ws.append([float(v) if isinstance(v, Decimal) else v for v in l])
            for i, nome in enumerate(cab, 1):
                larg = max([len(nome)] + [len(_celula_csv(l[i - 1])) for l in linhas[:500]])
                ws.column_dimensions[get_column_letter(i)].width = min(max(larg + 2, 9), 60)
                ws.cell(1, i).font = Font(bold=True, color="FFFFFF")
                ws.cell(1, i).fill = PatternFill("solid", fgColor="305496")
                fmt = _formato(nome)
                if fmt:
                    for (c,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                        c.number_format = fmt
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        caminho = SAIDA_DIR / f"relatorio_linhas_{self.sufixo}.xlsx"
        wb.save(caminho)
        print(f"  {caminho.relative_to(_bootstrap.RAIZ)}  ({len(wb.sheetnames)} abas)")


def _linha_venda(chave, reg) -> list:
    qtd, brinde, fat = reg["quantidade"], reg["qtd_brinde"], reg["faturamento"]
    pago = qtd - brinde
    return [*chave, reg["produto"], reg["categoria"] or "",
            _qtd(qtd), _qtd(brinde), _r2(fat), reg["pedidos"],
            _r2(fat / qtd) if qtd else None,
            _r2(fat / pago) if pago else None,
            _r2(reg["preco_min"]), _r2(reg["preco_max"])]


CAB_VENDAS = ["produto", "categoria", "quantidade", "qtd_brinde", "faturamento",
              "pedidos", "preco_medio", "preco_medio_sem_brinde",
              "preco_unit_min", "preco_unit_max"]


def relatorio_vendas(conn, ini, fim, saida: Saida):
    regs = etl_db.consultar(conn, SQL_VENDAS, (ini, fim))
    regs.sort(key=lambda r: (r["tenant"], r["ano_mes"], -r["faturamento"]))

    # Resumo por mes: consolidado e cada empresa, para a primeira aba.
    mes = defaultdict(lambda: defaultdict(lambda: [Decimal(0), Decimal(0), Decimal(0), set()]))
    for r in regs:
        for emp in ("consolidado", r["tenant"]):
            m = mes[r["ano_mes"]][emp]
            m[0] += r["quantidade"]
            m[1] += r["qtd_brinde"]
            m[2] += r["faturamento"]
            m[3].add(r["sku"])
    resumo = []
    empresas = ["consolidado", *_tenants({r["tenant"] for r in regs})]
    for am in sorted(mes):
        for emp in empresas:
            q, b, f, skus = mes[am][emp]
            resumo.append([am, emp, _qtd(q), _qtd(b), _r2(f), len(skus),
                           _r2(f / q) if q else None,
                           _r2(f / (q - b)) if q - b else None])
    saida.tabela("resumo_mensal", "Resumo mensal",
                 ["mes", "empresa", "quantidade", "qtd_brinde", "faturamento", "skus",
                  "preco_medio", "preco_medio_sem_brinde"], resumo)

    # Consolidado: soma as empresas por mes x SKU. pedidos soma direto --
    # pedido nao se repete entre empresas (a transferencia ja saiu na view).
    cons = {}
    for r in regs:
        k = (r["ano_mes"], r["sku"])
        a = cons.setdefault(k, {"produto": r["produto"], "categoria": r["categoria"],
                                "quantidade": 0, "qtd_brinde": 0, "faturamento": 0,
                                "pedidos": 0, "preco_min": None, "preco_max": None})
        for campo in ("quantidade", "qtd_brinde", "faturamento", "pedidos"):
            a[campo] += r[campo]
        a["categoria"] = a["categoria"] or r["categoria"]
        if r["preco_min"] is not None:
            a["preco_min"] = min(filter(None, (a["preco_min"], r["preco_min"])))
        a["preco_max"] = max(filter(lambda x: x is not None, (a["preco_max"], r["preco_max"])))
    ordem = sorted(cons.items(), key=lambda kv: (kv[0][0], -kv[1]["faturamento"]))
    saida.tabela("vendas_itens_mensal_consolidado", "Vendas itens (consolidado)",
                 ["mes", "sku"] + CAB_VENDAS, [_linha_venda(k, a) for k, a in ordem])
    saida.tabela("vendas_itens_mensal", "Vendas itens por empresa",
                 ["empresa", "mes", "sku"] + CAB_VENDAS,
                 [_linha_venda((r["tenant"], r["ano_mes"], r["sku"]), r) for r in regs])

    tot_q = sum(r["quantidade"] for r in regs)
    tot_f = sum(r["faturamento"] for r in regs)
    print(f"  total: {tot_q:,.0f} un / R$ {tot_f:,.2f} / "
          f"{len({r['sku'] for r in regs})} SKUs")


def _doc(v) -> str:
    return re.sub(r"\D", "", v or "")


def _tenants(presentes: set) -> list:
    """Tenants na ordem de data/tenants.json, mais algum que so exista no banco."""
    ordem = get_all_tenant_ids()
    return [t for t in ordem if t in presentes] + sorted(presentes - set(ordem))


def relatorio_clientes(conn, ini, fim, saida: Saida, linhas: dict):
    regs = etl_db.consultar(
        conn, SQL_CLIENTES.format(caso_linha=_sql_linha(linhas)), (ini, fim))
    tenants = _tenants({r["tenant"] for r in regs})
    por_linha = defaultdict(dict)
    for r in regs:
        doc = _doc(r["cpf_cnpj"])
        chave = doc or f"{r['tenant']}:{r['id_cliente']}"
        c = por_linha[r["linha"]].get(chave)
        if c is None:
            c = por_linha[r["linha"]][chave] = {
                "nome": r["nome"] or "", "cpf_cnpj": r["cpf_cnpj"] or "",
                "email": "", "telefone": "", "municipio": "", "uf": "",
                "empresas": set(), "pedidos": set(), "skus": set(), "canais": set(),
                "unidades": Decimal(0), "faturamento": Decimal(0),
                "primeira": r["data_pedido"], "ultima": r["data_pedido"]}
        for campo in ("email", "telefone", "municipio", "uf"):
            c[campo] = c[campo] or (r[campo] or "")
        c["nome"] = c["nome"] or (r["nome"] or "")
        c["empresas"].add(r["tenant"])
        c["pedidos"].add((r["tenant"], r["id_pedido"]))
        c["skus"].add(r["sku"])
        c["canais"].add(r["ecommerce_nome"] or "(sem integracao)")
        c["unidades"] += r["quantidade"]
        c["faturamento"] += r["valor_bruto"]
        c["primeira"] = min(c["primeira"], r["data_pedido"])
        c["ultima"] = max(c["ultima"], r["data_pedido"])

    cab = ["nome", "cpf_cnpj", "email", "telefone", "municipio", "uf", "empresas",
           "canais", "pedidos", "unidades", "faturamento", "primeira_compra",
           "ultima_compra", "skus_comprados"]
    resumo, tabelas = [], []
    for chave, (rotulo, *_) in linhas.items():
        clientes = sorted(por_linha[chave].values(), key=lambda c: -c["faturamento"])
        tabelas.append((f"clientes_{chave}", f"Clientes {rotulo}", [
            [c["nome"], c["cpf_cnpj"], c["email"], c["telefone"], c["municipio"],
             c["uf"], ", ".join(sorted(c["empresas"])), ", ".join(sorted(c["canais"])),
             len(c["pedidos"]), _qtd(c["unidades"]), _r2(c["faturamento"]),
             c["primeira"], c["ultima"], ", ".join(sorted(c["skus"]))]
            for c in clientes]))
        so = lambda t: sum(1 for c in clientes if c["empresas"] == {t})
        pedidos = sum(len(c["pedidos"]) for c in clientes)
        fat = sum((c["faturamento"] for c in clientes), Decimal(0))
        resumo.append([rotulo, len(clientes), *(so(t) for t in tenants),
                       sum(1 for c in clientes if len(c["empresas"]) > 1),
                       sum(1 for c in clientes if c["email"]),
                       sum(1 for c in clientes if c["telefone"]),
                       pedidos, _qtd(sum(c["unidades"] for c in clientes)),
                       _r2(fat), _r2(fat / pedidos) if pedidos else None])
    # Resumo antes das listas, para a contagem vir primeiro no xlsx.
    saida.tabela("resumo_clientes_linhas", "Clientes por linha (resumo)",
                 ["linha", "clientes", *(f"so_{t}" for t in tenants),
                  "nas_duas" if len(tenants) == 2 else "em_mais_de_uma",
                  "com_email", "com_telefone", "pedidos", "unidades", "faturamento",
                  "ticket_medio"], resumo)
    for arquivo, aba, linhas in tabelas:
        saida.tabela(arquivo, aba, cab, linhas)
    for r in resumo:
        print(f"  {r[0]:<13} {r[1]:>6} clientes")


def _notas(ini, fim, linhas: dict, extras: list) -> list:
    """Aba Notas do xlsx: as decisoes da docstring, para quem so abre a planilha."""
    return [
        f"Relatorio de vendas e clientes por linha -- {ini} a {fim}",
        "",
        "Fonte: warehouse tiny_dw (pipeline_etl), view v_venda_item. Nenhum dado de fora do banco.",
        "Vendas validas: situacao valida, natureza de venda, sem vendas entre as empresas do grupo.",
        "Faturamento = quantidade x preco unitario do item (valor BRUTO). Desconto do pedido e frete ficam de fora.",
        "Preco medio = faturamento / quantidade, com brinde (R$ 0,00) na quantidade; a coluna 'sem brinde' o exclui.",
        "Linha do produto = prefixo da categoria do cadastro, mais os SKUs extras e menos os excluidos da configuracao:",
        *(f"  {rot}: categoria '{pref}*'"
          + (f", + {', '.join(ext)}" if ext else "")
          + (f", - {', '.join(exc)}" if exc else "")
          for rot, pref, ext, exc in linhas.values()),
        "Cliente = quem comprou item da linha com preco > 0; quem so recebeu brinde nao conta.",
        "Cliente que comprou em mais de uma empresa aparece uma vez so (unificado por CPF/CNPJ).",
        "Comprador de marketplace nao tem email/telefone no Tiny -- limitacao da origem.",
        *extras,
    ]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    hoje = date.today()
    fim_mes_passado = date.fromordinal(date(hoje.year, hoje.month, 1).toordinal() - 1)
    ap.add_argument("--inicio", default=f"{fim_mes_passado.year}-01-01")
    ap.add_argument("--fim", default=fim_mes_passado.isoformat())
    ap.add_argument("--config", default=str(CONFIG_PADRAO),
                    help="JSON com as linhas de produto (ver data/linhas_produto.exemplo.json)")
    a = ap.parse_args()
    from pathlib import Path
    linhas, notas_extra = carregar_linhas(Path(a.config))
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    saida = Saida(f"{a.inicio.replace('-', '')}_{a.fim.replace('-', '')}")
    conn = etl_db.conectar()
    print(f"Periodo {a.inicio} a {a.fim} -- {etl_db.descrever(etl_db.carregar_config())}")
    print("[1] vendas por item x mes")
    relatorio_vendas(conn, a.inicio, a.fim, saida)
    print("[2] clientes por linha")
    relatorio_clientes(conn, a.inicio, a.fim, saida, linhas)
    print("[3] planilha consolidada")
    saida.xlsx(_notas(a.inicio, a.fim, linhas, notas_extra))


if __name__ == "__main__":
    main()
