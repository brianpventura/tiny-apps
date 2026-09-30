"""
Recorrencia de compra por produto -- quais SKUs sao comprados em mais e em
menos pedidos numa janela, com linha, categoria, estoque atual, quantidade,
faturamento, ultima venda e canal. Vendas lidas do warehouse tiny_dw; so o
ESTOQUE vem da API, porque o banco nao guarda estoque.

Saida em reports/pipeline_etl/pedidos-de-venda/:
  recorrencia_produtos_<ini>_<fim>.xlsx   abas Notas, Resumo, Mais recorrentes,
                                          Menos recorrentes, Base completa,
                                          SKU x Canal, Sem venda no periodo

Fases:
  1 estoque (online, so leitura): lista /produtos de cada empresa (~16 req) para
    casar SKU -> id vivo, e GET /estoque/{id} de cada SKU do relatorio
    (~1 req por SKU por empresa). Snapshot em data/estoque_snapshot_<tenant>.json;
    resumivel -- rodar de novo so busca o que falta (--renovar-estoque refaz).
  2 planilha (offline, NAO valida token): SQL no banco + snapshot de estoque.

DECISOES:
  - Venda = perfil de regras da empresa (core/regras_vendas.py) com o recorte
    'recorrencia' (regras.json -> relatorios.recorrencia), aplicado direto
    sobre `pedido`: situacoes validas e naturezas de venda do perfil, por
    tenant. Sem recorte, vale o mesmo perfil do warehouse (v_venda_item).
  - Continua valendo a regra do repo de excluir venda entre as empresas do
    grupo (cliente_excluido), senao transferencia de estoque vira "recompra".
    Exclusoes EXTRAS por nome (cliente OU vendedor) e por CNPJ vem de
    relatorios.recorrencia.excluir_nomes / excluir_cnpjs; sem elas, nada alem
    do cliente_excluido.
  - Recorrencia = numero de PEDIDOS distintos em que o SKU foi comprado com
    preco > 0. Brinde (R$ 0,00) nao e compra: entra em quantidade (saiu do
    estoque) e em qtd_brinde, mas nao conta pedido. SKU que so saiu como brinde
    fica na Base completa com recorrencia 0 e fora dos rankings.
  - Cliente distinto = CPF/CNPJ (so digitos), senao empresa + id do contato.
    clientes_recompra = clientes com 2+ pedidos do SKU no periodo.
  - SKU unificado entre as empresas (mesmo SKU = mesmo produto). Estoque sai
    por empresa e somado. Estoque = `saldo` de GET /estoque/{id}, que o Tiny ja
    calcula ignorando depositos marcados "desconsiderar" (Full, avariados...).
  - Linha = 1o nivel de categoria_caminho do cadastro, categoria = o resto.
    Sem ajuste manual (ao contrario de relatorio_linhas.py).
  - Faturamento = soma de qtd x valorUnitario (BRUTO; desconto de cabecalho e
    frete ficam de fora, a API nao tem desconto por item).
  - Canal = grupo de canal do pedido (Site / Marketplace / Vendas internas /
    Feiras/Eventos / Outros, classificar_canal do repo); canal principal = o de
    mais pedidos do SKU. A integracao crua sai na aba SKU x Canal.
  - Menos recorrentes = SKUs com >= 1 pedido pago, do menor para o maior.
    Produto que nao vendeu nada na janela nao cabe nesse ranking: vai para
    "Sem venda no periodo" (vendido antes do inicio da janela e hoje ativo).

Uso:
    python relatorio_recorrencia_produtos.py                     # fases 1,2 (01/01 ate hoje)
    python relatorio_recorrencia_produtos.py --fases 2           # so a planilha
    python relatorio_recorrencia_produtos.py --inicio 2026-01-01 --fim 2026-09-28 --top 50
    python relatorio_recorrencia_produtos.py --fases 1 --renovar-estoque
"""

import argparse
import io
import json
import re
import sys
import threading
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import _bootstrap
import etl_db
from core import regras_vendas as rv

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS = get_all_tenant_ids()  # chaves de data/tenants.json
SAIDA_DIR = _bootstrap.REPORTS_DIR / "pipeline_etl" / "pedidos-de-venda"

# Politica de venda: perfil de regras (core/regras_vendas.py) com o recorte
# 'recorrencia' -- regras.json -> relatorios.recorrencia pode trocar situacoes,
# naturezas e acrescentar exclusoes extras:
#   "excluir_nomes": [...]  cliente OU vendedor com esse nome sai do relatorio
#   "excluir_cnpjs": [...]  cliente com esse CPF/CNPJ sai do relatorio
# (alem do cliente_excluido de data/tenants.json). Serve para empresa ligada ao
# grupo que nao e um tenant, e para cobrir tambem o VENDEDOR.
RELATORIO = "recorrencia"
# `NOT IN ()` e erro de sintaxe: lista vazia vira um valor que nunca casa.
_NUNCA = "\x00"

_MAX_WORKERS = 4  # o rate limiter controla a taxa, nao o pool

# ------------------------------------------------------------------ #
#  SQL                                                                 #
# ------------------------------------------------------------------ #

def _in(valores) -> str:
    return ",".join(["%s"] * (len(valores) or 1))


def _p(valores) -> tuple:
    """Parametros de um IN: a propria lista, ou o sentinela se vazia."""
    return tuple(valores) or (_NUNCA,)


def _exclusoes() -> tuple[list, list]:
    extras = rv.perfil(None, RELATORIO).extras
    nomes = list(extras.get("excluir_nomes", []))
    cnpjs = [re.sub(r"\D", "", c) for c in extras.get("excluir_cnpjs", [])]
    return nomes, cnpjs


_FILTRO_CACHE: dict = {}


def filtro(conn) -> tuple[str, tuple]:
    """(trecho SQL sobre `p`, parametros) com a politica de venda do relatorio.

    Um bloco por tenant, porque cada empresa pode ter situacoes e naturezas
    proprias. As naturezas aceitas sao os nomes EXISTENTES no banco que o
    perfil aprova -- assim o filtro continua sendo IN por nome exato, e o modo
    'prefixo' do perfil funciona igual ao modo 'lista'.
    """
    if "f" in _FILTRO_CACHE:
        return _FILTRO_CACHE["f"]
    blocos, params = [], []
    for t in TENANTS:
        pf = rv.perfil(t, RELATORIO)
        nomes = [r["n"] for r in etl_db.consultar(
            conn, "SELECT DISTINCT natureza_nome AS n FROM pedido "
                  "WHERE tenant = %s AND natureza_nome <> ''", (t,))
            if pf.eh_natureza_venda(r["n"])]
        sits = sorted(pf.situacoes_validas)
        blocos.append(f"(p.tenant = %s AND p.situacao IN ({_in(sits)}) "
                      f"AND p.natureza_nome IN ({_in(nomes)}))")
        params += [t, *_p(sits), *_p(nomes)]
    nomes_exc, cnpjs_exc = _exclusoes()
    sql = (f"({' OR '.join(blocos) or '1=0'})"
           f"\n  AND p.cliente_excluido = 0"
           f"\n  AND COALESCE(p.cliente_nome, '') NOT IN ({_in(nomes_exc)})"
           f"\n  AND COALESCE(p.vendedor_nome, '') NOT IN ({_in(nomes_exc)})"
           f"\n  AND REGEXP_REPLACE(COALESCE(p.cliente_cpf_cnpj, ''), '[^0-9]', '')"
           f" NOT IN ({_in(cnpjs_exc)})"
           f"\n  AND p.tem_detalhe = 1")
    params += [*_p(nomes_exc), *_p(nomes_exc), *_p(cnpjs_exc)]
    _FILTRO_CACHE["f"] = (sql, tuple(params))
    return _FILTRO_CACHE["f"]


def consultar(conn, sql_modelo: str, antes: tuple) -> list:
    """Executa um SQL_* com o filtro de venda: `antes` = parametros das datas."""
    f_sql, f_params = filtro(conn)
    return etl_db.consultar(conn, sql_modelo.format(filtro=f_sql), (*antes, *f_params))


SQL_ITENS = """
SELECT p.tenant, p.id_pedido, p.data_pedido, p.canal,
       COALESCE(NULLIF(p.ecommerce_nome, ''), '(sem integracao)') AS integracao,
       p.id_cliente, p.cliente_cpf_cnpj,
       i.id_produto, i.sku,
       COALESCE(NULLIF(NULLIF(d.descricao, ''), '(nao informado)'), i.descricao_item) AS produto,
       d.categoria_caminho, d.tipo_cadastro,
       i.quantidade, i.valor_unitario, i.valor_bruto
FROM pedido p
JOIN pedido_item i ON i.tenant = p.tenant AND i.id_pedido = p.id_pedido
LEFT JOIN dim_produto d ON d.tenant = i.tenant AND d.id_produto = i.id_produto
WHERE p.data_pedido BETWEEN %s AND %s AND {filtro}
  AND i.quantidade > 0
"""

# Ultima venda de cada SKU ANTES da janela, pelas mesmas regras -- alimenta a
# aba "Sem venda no periodo".
SQL_HISTORICO = """
SELECT i.sku,
       MAX(p.data_pedido) AS ultima,
       COUNT(DISTINCT p.tenant, p.id_pedido) AS pedidos,
       SUM(i.quantidade) AS quantidade,
       SUM(i.valor_bruto) AS faturamento,
       MAX(COALESCE(NULLIF(NULLIF(d.descricao, ''), '(nao informado)'), i.descricao_item)) AS produto,
       MAX(d.categoria_caminho) AS categoria_caminho,
       MAX(d.tipo_cadastro) AS tipo_cadastro,
       GROUP_CONCAT(DISTINCT p.tenant) AS empresas
FROM pedido p
JOIN pedido_item i ON i.tenant = p.tenant AND i.id_pedido = p.id_pedido
LEFT JOIN dim_produto d ON d.tenant = i.tenant AND d.id_produto = i.id_produto
WHERE p.data_pedido < %s AND i.sku <> '' AND i.valor_unitario > 0 AND {filtro}
  AND i.quantidade > 0
GROUP BY i.sku
"""

SQL_HEADER = """
SELECT COUNT(*) AS pedidos, SUM(valor_total_produtos) AS vprod
FROM pedido p
WHERE p.data_pedido BETWEEN %s AND %s AND {filtro}
"""

SQL_DIM_PRODUTO = """
SELECT tenant, id_produto, sku, situacao FROM dim_produto WHERE sku <> ''
"""

# ------------------------------------------------------------------ #
#  Fase 1 -- estoque (online)                                          #
# ------------------------------------------------------------------ #

def _arq_snapshot(tenant):
    return _bootstrap.DATA_DIR / f"estoque_snapshot_{tenant}.json"


def _ler_snapshot(tenant) -> dict:
    try:
        return json.loads(_arq_snapshot(tenant).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}


def _salvar_snapshot(tenant, snap) -> None:
    arq = _arq_snapshot(tenant)
    tmp = arq.with_suffix(".tmp")
    tmp.write_text(json.dumps(snap, ensure_ascii=False), encoding="utf-8")
    tmp.replace(arq)


def _num(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _skus_do_relatorio(conn, ini, fim) -> set:
    skus = {r["sku"] for r in consultar(conn, SQL_ITENS, (ini, fim))
            if r["sku"]}
    skus |= {r["sku"] for r in consultar(conn, SQL_HISTORICO, (ini,))}
    return skus


def fase_estoque(conn, ini, fim, renovar: bool) -> None:
    from core.auth_tiny import validate_tenants
    from core.tiny_client import TinyClient
    from core.produtos_utils import desempacotar_data

    if not validate_tenants(TENANTS, interactive=False):
        print("  token invalido -- rodar src/apps/auth/_reauth.py <tenant>")
        sys.exit(1)
    skus = _skus_do_relatorio(conn, ini, fim)
    dim = defaultdict(dict)  # tenant -> sku -> id (fallback: variacao pode faltar na listagem)
    for r in etl_db.consultar(conn, SQL_DIM_PRODUTO):
        if r["situacao"] != "E":
            dim[r["tenant"]].setdefault(r["sku"], r["id_produto"])
    print(f"  {len(skus)} SKUs no relatorio")

    for tenant in TENANTS:
        client = TinyClient(tenant)
        snap = {} if renovar else _ler_snapshot(tenant)
        if not snap.get("cadastro"):
            cadastro = {}
            for it in client.produtos_todos():
                sku = str(it.get("sku") or "").strip()
                sit = str(it.get("situacao") or "").strip().upper()
                if not sku or sit == "E":
                    continue
                atual = cadastro.get(sku)
                # SKU com dois vivos: fica o ativo.
                if atual is None or (atual["situacao"] != "A" and sit == "A"):
                    cadastro[sku] = {"id": it.get("id"), "situacao": sit,
                                     "descricao": it.get("descricao") or ""}
            snap = {"cadastro": cadastro, "cadastro_em": datetime.now().isoformat(timespec="seconds"),
                    "estoque": {}}
            _salvar_snapshot(tenant, snap)
            print(f"  [{tenant}] cadastro: {len(cadastro)} SKUs vivos")

        alvo = {}
        for sku in skus:
            pid = (snap["cadastro"].get(sku) or {}).get("id") or dim[tenant].get(sku)
            if pid and str(pid) not in snap["estoque"]:
                alvo[str(pid)] = sku
        print(f"  [{tenant}] estoque: {len(alvo)} a buscar ({len(snap['estoque'])} no snapshot)"
              f" ~{len(alvo) / 2 / 60:.0f} min")
        lock, cont = threading.Lock(), {"n": 0, "erros": 0}

        def _est(pid):
            try:
                e = desempacotar_data(client.obter_estoque(pid))
            except Exception as exc:  # 404 de variacao/cadastro removido: fica sem estoque
                return pid, None, str(exc)[:120]
            return pid, {"saldo": _num(e.get("saldo")), "reservado": _num(e.get("reservado")),
                         "disponivel": _num(e.get("disponivel")),
                         "em": datetime.now().isoformat(timespec="seconds")}, None

        with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as ex:
            for fut in as_completed([ex.submit(_est, pid) for pid in alvo]):
                pid, info, erro = fut.result()
                with lock:
                    if info is None:
                        cont["erros"] += 1
                        snap.setdefault("erros", {})[pid] = erro
                    else:
                        snap["estoque"][pid] = info
                    cont["n"] += 1
                    if cont["n"] % 100 == 0:
                        _salvar_snapshot(tenant, snap)
                        print(f"    {cont['n']}/{len(alvo)}")
        _salvar_snapshot(tenant, snap)
        print(f"  [{tenant}] ok: {len(snap['estoque'])} saldos, {cont['erros']} erros")


# ------------------------------------------------------------------ #
#  Fase 2 -- planilha (offline)                                        #
# ------------------------------------------------------------------ #

DOIS = Decimal("0.01")


def _r2(v):
    return None if v is None else Decimal(v).quantize(DOIS)


def _qtd(v):
    d = Decimal(v)
    return int(d) if d == d.to_integral_value() else float(d)


def _doc(v) -> str:
    return re.sub(r"\D", "", v or "")


def _linha_categoria(caminho):
    partes = [p.strip() for p in (caminho or "").split("->") if p.strip()]
    if not partes:
        return "(sem categoria)", ""
    return partes[0], " -> ".join(partes[1:])


def _tipo(t):
    return {"K": "Kit", "V": "Variacao", "S": "Simples", "M": "Materia-prima",
            "F": "Fabricado"}.get(t or "", t or "")


class Estoque:
    """sku -> saldo por empresa, lido dos snapshots da fase 1."""

    def __init__(self, conn):
        self.snaps = {t: _ler_snapshot(t) for t in TENANTS}
        self.dim = defaultdict(dict)
        for r in etl_db.consultar(conn, SQL_DIM_PRODUTO):
            if r["situacao"] != "E":
                self.dim[r["tenant"]].setdefault(r["sku"], r["id_produto"])
        self.data = max((s.get("cadastro_em", "") for s in self.snaps.values()), default="")

    def _info(self, tenant, sku):
        snap = self.snaps[tenant]
        cad = (snap.get("cadastro") or {}).get(sku) or {}
        pid = cad.get("id") or self.dim[tenant].get(sku)
        return cad, (snap.get("estoque") or {}).get(str(pid)) if pid else None

    def situacao(self, sku) -> str:
        sits = {self._info(t, sku)[0].get("situacao") for t in TENANTS} - {None}
        if "A" in sits:
            return "A"
        # sem cadastro vivo = SKU excluido (situacao E) em todas as empresas
        return "/".join(sorted(sits)) or "Excluido"

    def linha(self, sku) -> list:
        """[estoque_<tenant> para cada tenant, estoque_total, disponivel_total]"""
        saldos, disp = [], []
        for t in TENANTS:
            e = self._info(t, sku)[1]
            saldos.append(_qtd(Decimal(str(e["saldo"]))) if e else None)
            disp.append(e["disponivel"] if e else None)
        conhecidos = [s for s in saldos if s is not None]
        return [*saldos,
                _qtd(Decimal(str(sum(conhecidos)))) if conhecidos else None,
                _qtd(Decimal(str(sum(d for d in disp if d is not None))))
                if any(d is not None for d in disp) else None]


def _chave(r) -> str:
    """SKU unificado; item sem SKU (id_produto 0, pedido digitado) agrupa pela descricao."""
    return r["sku"] or f"(sem SKU) {r['produto']}"


def _agregar(regs):
    """Uma entrada por SKU unificado (SKU vazio: empresa + id do produto)."""
    skus = {}
    for r in regs:
        chave = _chave(r)
        a = skus.get(chave)
        if a is None:
            a = skus[chave] = {
                "sku": chave, "sku_real": bool(r["sku"]), "produto": r["produto"], "caminho": r["categoria_caminho"],
                "tipo": r["tipo_cadastro"], "data_nome": r["data_pedido"],
                "empresas": set(), "pedidos": set(), "pedidos_todos": set(),
                "clientes": Counter(), "meses": set(), "qtd": Decimal(0),
                "qtd_brinde": Decimal(0), "fat": Decimal(0), "primeira": r["data_pedido"],
                "ultima": r["data_pedido"], "ultima_compra": None,
                "canal_ped": defaultdict(set), "canal_ped_todos": defaultdict(set),
                "precos": []}
        # nome e categoria do cadastro mais recente vencem
        if r["data_pedido"] >= a["data_nome"]:
            a["data_nome"] = r["data_pedido"]
            a["produto"] = r["produto"] or a["produto"]
        a["caminho"] = a["caminho"] or r["categoria_caminho"]
        a["tipo"] = a["tipo"] or r["tipo_cadastro"]
        ped = (r["tenant"], r["id_pedido"])
        a["empresas"].add(r["tenant"])
        a["pedidos_todos"].add(ped)
        a["canal_ped_todos"][r["canal"] or "Outros"].add(ped)
        a["qtd"] += r["quantidade"]
        a["primeira"] = min(a["primeira"], r["data_pedido"])
        a["ultima"] = max(a["ultima"], r["data_pedido"])
        if r["valor_unitario"] > 0:
            a["pedidos"].add(ped)
            a["canal_ped"][r["canal"] or "Outros"].add(ped)
            cli = _doc(r["cliente_cpf_cnpj"]) or f"{r['tenant']}:{r['id_cliente']}"
            a["clientes"][cli] += 1  # 1 por linha; pedidos distintos contados abaixo
            a["meses"].add(r["data_pedido"].strftime("%Y-%m"))
            a["fat"] += r["valor_bruto"]
            a["ultima_compra"] = max(filter(None, (a["ultima_compra"], r["data_pedido"])))
            a["precos"].append(r["valor_unitario"])
        else:
            a["qtd_brinde"] += r["quantidade"]
    return skus


def _clientes_por_pedido(regs):
    """cliente -> pedidos distintos, por SKU (o Counter de _agregar conta linhas)."""
    out = defaultdict(lambda: defaultdict(set))
    for r in regs:
        if r["valor_unitario"] > 0:
            chave = _chave(r)
            cli = _doc(r["cliente_cpf_cnpj"]) or f"{r['tenant']}:{r['id_cliente']}"
            out[chave][cli].add((r["tenant"], r["id_pedido"]))
    return out


def _canais(a):
    base = a["canal_ped"] if a["pedidos"] else a["canal_ped_todos"]
    total = sum(len(p) for p in base.values())
    ordem = sorted(base.items(), key=lambda kv: -len(kv[1]))
    principal = ordem[0][0] if ordem else ""
    dist = " | ".join(f"{c} {len(p) / total:.0%}" for c, p in ordem) if total else ""
    return principal, dist


_CAB_ESTOQUE = [f"estoque_{t}" for t in TENANTS]
CAB = ["ranking", "sku", "produto", "linha", "categoria", "tipo", "situacao_cadastro",
       *_CAB_ESTOQUE, "estoque_total", "estoque_disponivel",
       "pedidos", "clientes", "clientes_recompra", "pct_recompra", "meses_com_venda",
       "quantidade", "qtd_brinde", "faturamento", "preco_medio", "primeira_venda",
       "ultima_venda", "dias_sem_venda", "canal_principal", "canais", "empresas"]


def _formato(col):
    if col in ("faturamento", "preco_medio"):
        return "#,##0.00"
    if col == "pct_recompra":
        return "0.0%"
    if col.endswith("_venda") and col != "dias_sem_venda" and col != "meses_com_venda":
        return "dd/mm/yyyy"
    if col in ("ranking", "pedidos", "clientes", "clientes_recompra", "meses_com_venda",
               "quantidade", "qtd_brinde", "dias_sem_venda") or col.startswith("estoque"):
        return "#,##0"
    return None


def _escrever_aba(wb, nome, cab, linhas, larguras=None):
    ws = wb.create_sheet(nome[:31])
    ws.append(cab)
    for l in linhas:
        ws.append([float(v) if isinstance(v, Decimal) else v for v in l])
    for i, col in enumerate(cab, 1):
        amostra = [len(str(l[i - 1])) for l in linhas[:300] if l[i - 1] is not None]
        larg = (larguras or {}).get(col) or min(max([len(col)] + amostra) + 2, 50)
        ws.column_dimensions[get_column_letter(i)].width = max(larg, 9)
        ws.cell(1, i).font = Font(bold=True, color="FFFFFF")
        ws.cell(1, i).fill = PatternFill("solid", fgColor="305496")
        fmt = _formato(col)
        if fmt:
            for (c,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                c.number_format = fmt
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def fase_planilha(conn, ini, fim, top: int) -> None:
    fim_d = date.fromisoformat(fim)
    regs = consultar(conn, SQL_ITENS, (ini, fim))
    skus = _agregar(regs)
    cli_ped = _clientes_por_pedido(regs)
    est = Estoque(conn)
    if not est.data:
        print("  AVISO: sem snapshot de estoque -- rodar --fases 1; colunas de estoque vazias")

    linhas = []
    for chave, a in skus.items():
        linha, cat = _linha_categoria(a["caminho"])
        cp = cli_ped.get(chave, {})
        recompra = sum(1 for p in cp.values() if len(p) >= 2)
        principal, dist = _canais(a)
        qtd_paga = a["qtd"] - a["qtd_brinde"]
        linhas.append([
            None, chave, a["produto"], linha, cat, _tipo(a["tipo"]), "(sem SKU)" if not a["sku_real"] else est.situacao(chave),
            *est.linha(chave),
            len(a["pedidos"]), len(cp), recompra,
            round(recompra / len(cp), 4) if cp else None, len(a["meses"]),
            _qtd(a["qtd"]), _qtd(a["qtd_brinde"]), _r2(a["fat"]),
            _r2(a["fat"] / qtd_paga) if qtd_paga else None,
            a["primeira"], a["ultima"], (fim_d - a["ultima"]).days,
            principal, dist, ", ".join(sorted(a["empresas"]))])
    ip, ifat = CAB.index("pedidos"), CAB.index("faturamento")
    linhas.sort(key=lambda l: (-l[ip], -(l[ifat] or 0), l[1]))
    for n, l in enumerate(linhas, 1):
        l[0] = n if l[ip] else None
    compradas = [l for l in linhas if l[ip]]
    mais = compradas[:top]
    menos = sorted(compradas, key=lambda l: (l[ip], l[ifat] or 0, l[1]))[:top]

    # SKU x canal x integracao
    sc = defaultdict(lambda: {"ped": set(), "qtd": Decimal(0), "fat": Decimal(0), "ult": None})
    for r in regs:
        chave = _chave(r)
        s = sc[(chave, r["canal"] or "Outros", r["integracao"])]
        if r["valor_unitario"] > 0:
            s["ped"].add((r["tenant"], r["id_pedido"]))
            s["fat"] += r["valor_bruto"]
        s["qtd"] += r["quantidade"]
        s["ult"] = max(filter(None, (s["ult"], r["data_pedido"])))
    canal_linhas = sorted(
        ([k[0], skus[k[0]]["produto"], k[1], k[2], len(v["ped"]), _qtd(v["qtd"]),
          _r2(v["fat"]), v["ult"]] for k, v in sc.items()),
        key=lambda l: (l[0], -l[4], -l[5]))

    # Sem venda no periodo: vendido antes (mesmas regras), nenhuma linha na janela,
    # e ativo no cadastro de alguma empresa.
    hist = consultar(conn, SQL_HISTORICO, (ini,))
    sem_venda = []
    for h in hist:
        if h["sku"] in skus:
            continue
        sit = est.situacao(h["sku"])
        if est.data and sit != "A":
            continue
        linha, cat = _linha_categoria(h["categoria_caminho"])
        sem_venda.append([h["sku"], h["produto"], linha, cat, _tipo(h["tipo_cadastro"]), sit,
                          *est.linha(h["sku"]), h["ultima"], (fim_d - h["ultima"]).days,
                          h["pedidos"], _qtd(h["quantidade"]), _r2(h["faturamento"]),
                          h["empresas"].replace(",", ", ")])
    sem_venda.sort(key=lambda l: (-(l[8] or 0), l[10]))  # mais estoque parado primeiro

    hdr = consultar(conn, SQL_HEADER, (ini, fim))[0]
    tot_fat = sum((a["fat"] for a in skus.values()), Decimal(0))
    tot_qtd = sum((a["qtd"] for a in skus.values()), Decimal(0))
    ped_total = len({(r["tenant"], r["id_pedido"]) for r in regs})
    por_emp = defaultdict(lambda: [set(), Decimal(0), Decimal(0)])
    for r in regs:
        e = por_emp[r["tenant"]]
        e[0].add(r["id_pedido"])
        e[1] += r["quantidade"]
        e[2] += r["valor_bruto"]
    faixas = [("1 pedido", 1, 1), ("2 a 5", 2, 5), ("6 a 20", 6, 20), ("21 a 100", 21, 100),
              ("101 a 500", 101, 500), ("mais de 500", 501, 10 ** 9)]
    resumo = [["Periodo", f"{date.fromisoformat(ini):%d/%m/%Y} a {fim_d:%d/%m/%Y}"],
              ["Pedidos", ped_total],
              *([f"  {t}", len(por_emp[t][0])] for t in TENANTS),
              ["Unidades (inclui brinde)", _qtd(tot_qtd)],
              ["Faturamento bruto dos itens (R$)", _r2(tot_fat)],
              *([f"  {t} (R$)", _r2(por_emp[t][2])] for t in TENANTS),
              ["Conferencia: valorTotalProdutos dos cabecalhos (R$)", _r2(hdr["vprod"])],
              ["SKUs com venda", len(skus)],
              ["  comprados (>= 1 pedido pago)", len(compradas)],
              ["  so como brinde", len(skus) - len(compradas)],
              ["SKUs ativos sem venda no periodo (vendidos antes)", len(sem_venda)],
              ["Estoque: snapshot da API em", est.data.replace("T", " ") or "(sem snapshot)"],
              [],
              ["Distribuicao da recorrencia (pedidos por SKU)", "SKUs"]]
    for rot, lo, hi in faixas:
        resumo.append([f"  {rot}", sum(1 for l in compradas if lo <= l[ip] <= hi)])

    notas = [
        f"Recorrencia de compra por produto -- {date.fromisoformat(ini):%d/%m/%Y} a {fim_d:%d/%m/%Y}",
        "",
        "Fonte das vendas: warehouse tiny_dw (pipeline_etl). Estoque: GET /estoque/{id} da API, snapshot na data indicada no Resumo.",
        *(f"Regras de venda [{t}]: situacoes "
          + ", ".join(rv.SITUACOES.get(x, str(x))
                      for x in sorted(rv.perfil(t, RELATORIO).situacoes_validas))
          + "; naturezas: " + rv.perfil(t, RELATORIO).descrever_naturezas() + "."
          for t in TENANTS),
        "Vendas entre as empresas do grupo (incluindo a propria razao social) estao excluidas.",
        *(["Excluidos tambem os pedidos com cliente OU vendedor: " + "; ".join(_exclusoes()[0])
           + (f" (CNPJ {', '.join(_exclusoes()[1])})" if _exclusoes()[1] else "") + "."]
          if any(_exclusoes()) else []),
        "RECORRENCIA = numero de pedidos distintos em que o SKU foi comprado com preco > 0. Brinde (R$ 0,00) nao conta pedido.",
        "clientes = compradores distintos (por CPF/CNPJ); clientes_recompra = os que compraram o SKU em 2+ pedidos; pct_recompra = recompra / clientes.",
        "meses_com_venda = em quantos meses da janela houve pelo menos um pedido pago do SKU (regularidade).",
        "quantidade inclui brinde (saiu do estoque); qtd_brinde mostra a parte gratis. preco_medio = faturamento / unidades pagas.",
        "Faturamento = qtd x preco unitario do item (BRUTO). Desconto do pedido e frete nao entram.",
        "SKU unificado entre as empresas. Estoque por empresa (saldo, ja sem depositos 'desconsiderar') e somado; disponivel = saldo - reservado.",
        "Estoque vazio = SKU sem cadastro vivo naquela empresa. situacao_cadastro 'Excluido' = SKU vendido na janela mas hoje excluido do cadastro em todas as empresas (foi substituido/descontinuado).",
        "Linha = 1o nivel da categoria do cadastro; categoria = subnivel. '(sem categoria)' = produto sem categoria no Tiny.",
        "Kits do site ja vem explodidos em componentes no pedido; a linha com tipo Kit e o kit vendido como item proprio.",
        "canal_principal = canal com mais pedidos do SKU; 'canais' traz a participacao de cada canal em pedidos. Integracao crua na aba SKU x Canal.",
        "dias_sem_venda = dias entre a ultima venda e o fim do periodo.",
        f"Mais/Menos recorrentes = {top} primeiros/ultimos entre os SKUs com >= 1 pedido pago. A Base completa traz todos.",
        "Sem venda no periodo = SKU ativo no cadastro, vendido antes do inicio da janela e sem nenhuma venda nela. Ordenado por estoque.",
        "Venda feita fora do Tiny (ex.: marketplace operado por outro integrador) nao esta no banco.",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Notas"
    for t in notas:
        ws.append([t])
    ws.column_dimensions["A"].width = 140
    ws.cell(1, 1).font = Font(bold=True, size=12)

    ws = wb.create_sheet("Resumo")
    for l in resumo:
        ws.append([float(v) if isinstance(v, Decimal) else v for v in l])
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 26
    for (c,) in ws.iter_rows(min_col=2, max_col=2):
        if isinstance(c.value, float):
            c.number_format = "#,##0.00"
        elif isinstance(c.value, int):
            c.number_format = "#,##0"

    larg = {"produto": 48, "canais": 44}
    _escrever_aba(wb, f"Mais recorrentes (top {top})", CAB, mais, larg)
    _escrever_aba(wb, f"Menos recorrentes ({top})", CAB, menos, larg)
    _escrever_aba(wb, "Base completa", CAB, linhas, larg)
    _escrever_aba(wb, "SKU x Canal", ["sku", "produto", "canal", "integracao", "pedidos",
                                      "quantidade", "faturamento", "ultima_venda"],
                  canal_linhas, larg)
    cab_sv = ["sku", "produto", "linha", "categoria", "tipo", "situacao_cadastro",
              *_CAB_ESTOQUE, "estoque_total", "estoque_disponivel",
              "ultima_venda", "dias_sem_venda", "pedidos", "quantidade", "faturamento",
              "empresas"]
    ws = _escrever_aba(wb, "Sem venda no periodo", cab_sv, sem_venda, larg)
    ws.cell(1, cab_sv.index("pedidos") + 1).value = "pedidos_antes_da_janela"

    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    arq = SAIDA_DIR / f"recorrencia_produtos_{ini.replace('-', '')}_{fim.replace('-', '')}.xlsx"
    wb.save(arq)
    print(f"  {ped_total} pedidos / {_qtd(tot_qtd):,} un / R$ {tot_fat:,.2f} "
          f"(cabecalhos R$ {hdr['vprod']:,.2f}) / {len(skus)} SKUs ({len(compradas)} comprados)"
          f" / {len(sem_venda)} sem venda")
    print(f"  {arq.relative_to(_bootstrap.RAIZ)}")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--inicio", default=f"{date.today().year}-01-01")
    ap.add_argument("--fim", default=date.today().isoformat())
    ap.add_argument("--fases", default="1,2")
    ap.add_argument("--top", type=int, default=50)
    ap.add_argument("--renovar-estoque", action="store_true",
                    help="descarta o snapshot de estoque e busca tudo de novo")
    a = ap.parse_args()
    fases = {int(f) for f in a.fases.split(",")}
    conn = etl_db.conectar()
    print(f"Periodo {a.inicio} a {a.fim} -- {etl_db.descrever(etl_db.carregar_config())}")
    if 1 in fases:
        print("[1] estoque (API)")
        fase_estoque(conn, a.inicio, a.fim, a.renovar_estoque)
    if 2 in fases:
        print("[2] planilha")
        fase_planilha(conn, a.inicio, a.fim, a.top)


if __name__ == "__main__":
    main()
