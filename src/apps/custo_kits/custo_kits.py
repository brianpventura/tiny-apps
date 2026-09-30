"""
Preco de CUSTO dos kits no Tiny = soma do custo dos componentes (custo x quantidade).

Por empresa:
    1. Lista TODOS os produtos (~15 requests: a listagem de /produtos ja traz 'tipo',
       'situacao' e o bloco 'precos' com precoCusto). Kit = tipo 'K', situacao != 'E'.
    2. GET /produtos/{kit} de cada kit -> array 'kit' (componentes e quantidades).
    3. Custo do componente: o precoCusto da listagem. Excecoes que vao ao GET direto:
       componente que nao esta na listagem, e variacao (tipoVariacao 'V'), cujo custo
       confiavel so vem do GET da propria variacao. Componente que e kit entra pelo
       custo CALCULADO do sub-kit (recursivo, com limite de profundidade).
    4. Preview -> confirmacao -> PUT /produtos/{kit} alterando SO precos.precoCusto ->
       releitura conferindo que o custo persistiu e que a composicao nao mudou.
    5. Planilha xlsx (abas Kits e Componentes) e log CSV em reports/.

Uso:
    python custo_kits.py [sku1 sku2 ...] [--tenants A,B] [--commit]
                         [--limite N] [--aceitar-incompleto] [--sim]

    sem SKUs            processa todos os kits da empresa
    --tenants A,B       default: todos os de data/tenants.json
    --commit            grava de verdade. SEM ele nada e escrito (dry-run e o padrao).
    --limite N          so os N primeiros kits (teste inicial)
    --aceitar-incompleto grava tambem kit com componente de custo zero/ausente
                        (default: NAO grava — o custo sairia subestimado em silencio)
    --sim               dispensa a confirmacao interativa do --commit

Decisoes:
    * Kit com algum componente sem custo (0 ou ausente) fica 'custo_incompleto' e NAO e
      gravado por padrao: um custo parcial parece valido no ERP e distorce margem.
    * Kit sem componentes cadastrados fica 'sem_componentes' e nao e gravado.
    * Kit cujo custo atual ja e igual ao calculado (diferenca < R$ 0,005) fica 'igual'.
    * Kit que e VARIACAO (tipoVariacao 'V' — os 52 colox* de cada empresa) grava por
      PUT /produtos/{pai}/variacoes/{id}; PUT /produtos/{id} da 404 para variacao.
    * O payload do PUT e o GET completo passado por core.preparar_payload_put, que
      PRESERVA os fornecedores (sanitizar_payload_produto os apagaria). O bloco 'kit'
      vai junto e a releitura confere que o numero de componentes nao mudou.
"""

import csv
import os
import sys
from datetime import datetime
from pathlib import Path

_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from core import TinyClient, validate_tenants, desempacotar_data
from core.produtos_utils import preparar_payload_put

from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS_PADRAO = get_all_tenant_ids()  # chaves de data/tenants.json
REPORTS_DIR = _SRC.parent / "reports"
PROFUNDIDADE_MAX = 5
TOLERANCIA = 0.005


def _f(v) -> float | None:
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _custo(d: dict) -> float | None:
    return _f((d.get("precos") or {}).get("precoCusto"))


# ------------------------------------------------------------------ #
#  Coleta e calculo                                                   #
# ------------------------------------------------------------------ #

def listar_catalogo(client: TinyClient) -> dict:
    """id -> item da listagem (todos os produtos, inclusive inativos e 'E')."""
    catalogo = {}
    for it in client.produtos_todos(on_total=lambda t: print(f"  catalogo: {t} produtos")):
        catalogo[it.get("id")] = it
    print(f"  catalogo lido: {len(catalogo)} produtos")
    return catalogo


class Calculadora:
    def __init__(self, client: TinyClient, catalogo: dict):
        self.client = client
        self.catalogo = catalogo
        self.detalhe: dict = {}        # id -> GET completo (cache em memoria)
        self.resultado: dict = {}      # kit_id -> dict do kit calculado

    def _get(self, pid) -> dict:
        if pid not in self.detalhe:
            self.detalhe[pid] = desempacotar_data(self.client.obter_produto(pid))
        return self.detalhe[pid]

    def _custo_componente(self, pid, profundidade: int) -> tuple[float | None, str, str, str, str]:
        """(custo unitario, origem, sku, descricao, tipo)."""
        item = self.catalogo.get(pid)
        if item is None or (item.get("tipoVariacao") or "").upper() == "V":
            full = self._get(pid)
            origem = "GET variacao" if item is not None else "GET (fora da listagem)"
            item_ref = full
        else:
            full = None
            origem = "listagem"
            item_ref = item
        tipo = str(item_ref.get("tipo") or "").upper()
        sku = str(item_ref.get("sku") or "")
        desc = str(item_ref.get("descricao") or "")
        if tipo == "K":
            sub = self.calcular(pid, profundidade + 1)
            if sub["status"] in ("ok", "igual"):
                return sub["custo_calculado"], "sub-kit", sku, desc, tipo
            return None, f"sub-kit {sub['status']}", sku, desc, tipo
        return _custo(item_ref), origem, sku, desc, tipo

    def calcular(self, kit_id, profundidade: int = 0) -> dict:
        if kit_id in self.resultado:
            return self.resultado[kit_id]
        if profundidade > PROFUNDIDADE_MAX:
            return {"status": "ciclo_ou_profundo", "custo_calculado": None}
        full = self._get(kit_id)
        r = {
            "id": kit_id, "sku": str(full.get("sku") or ""),
            "descricao": str(full.get("descricao") or "").strip(),
            "situacao": full.get("situacao"), "custo_atual": _custo(full),
            "custo_calculado": None, "n_componentes": 0,
            "componentes_sem_custo": "", "componentes": [], "status": "",
            "full": full,
        }
        comps = full.get("kit") or []
        r["n_componentes"] = len(comps)
        if not comps:
            r["status"] = "sem_componentes"
            self.resultado[kit_id] = r
            return r
        soma, sem_custo = 0.0, []
        for comp in comps:
            prod = comp.get("produto") or {}
            pid = prod.get("id")
            qtd = _f(comp.get("quantidade")) or 0.0
            try:
                cu, origem, sku, desc, tipo = self._custo_componente(pid, profundidade)
            except Exception as e:
                cu, origem, sku, desc, tipo = None, f"erro: {e}", "", "", ""
            sku = sku or str(prod.get("sku") or "")
            desc = desc or str(prod.get("descricao") or "")
            if not cu or cu <= 0:
                sem_custo.append(sku or str(pid))
            sub = (cu or 0.0) * qtd
            soma += sub
            r["componentes"].append({
                "kit_sku": r["sku"], "kit_descricao": r["descricao"],
                "componente_id": pid, "componente_sku": sku,
                "componente_descricao": desc, "componente_tipo": tipo,
                "quantidade": qtd, "custo_unitario": cu if cu is not None else "",
                "subtotal": round(sub, 4), "origem_custo": origem,
            })
        r["custo_calculado"] = round(soma, 2)
        r["componentes_sem_custo"] = ", ".join(sem_custo)
        if sem_custo:
            r["status"] = "custo_incompleto"
        elif r["custo_atual"] is not None and abs(r["custo_atual"] - soma) < TOLERANCIA:
            r["status"] = "igual"
        else:
            r["status"] = "ok"
        self.resultado[kit_id] = r
        return r


def pendentes_de_gravacao(kits: list[dict], aceitar_incompleto: bool) -> list[dict]:
    alvo = {"ok"} | ({"custo_incompleto"} if aceitar_incompleto else set())
    return [k for k in kits if k["status"] in alvo and (k["custo_calculado"] or 0) > 0]


# ------------------------------------------------------------------ #
#  Escrita                                                            #
# ------------------------------------------------------------------ #

def gravar(client: TinyClient, kits: list[dict]) -> None:
    print(f"\n  Gravando precoCusto em {len(kits)} kit(s)...")
    for i, k in enumerate(kits, 1):
        novo = k["custo_calculado"]
        print(f"  [{i}/{len(kits)}] {k['sku']} -> R$ {novo:,.2f} ... ", end="", flush=True)
        try:
            full = desempacotar_data(client.obter_produto(k["id"]))   # estado fresco
            n_antes = len(full.get("kit") or [])
            payload = preparar_payload_put(full)
            precos = dict(payload.get("precos") or {})
            precos["precoCusto"] = novo
            payload["precos"] = precos
            if (full.get("tipoVariacao") or "").upper() == "V":
                # kit que e variacao (os colox*): PUT /produtos/{id} da 404
                pai = (full.get("produtoPai") or {}).get("id")
                if not pai:
                    raise RuntimeError("variacao sem produtoPai")
                client.atualizar_variacao(pai, k["id"], payload)
            else:
                client.atualizar_produto(k["id"], payload)
            apos = desempacotar_data(client.obter_produto(k["id"]))
            relido = _custo(apos) or 0.0
            n_apos = len(apos.get("kit") or [])
            k["custo_relido"] = relido
            if abs(relido - novo) > 0.01:
                k["resultado"] = f"nao_persistiu (relido {relido:.2f})"
            elif n_apos != n_antes:
                k["resultado"] = f"ALERTA composicao {n_antes}->{n_apos}"
            else:
                k["resultado"] = "atualizado"
            print(k["resultado"])
        except Exception as e:
            k["resultado"] = f"erro_put: {e}"
            print(k["resultado"])


# ------------------------------------------------------------------ #
#  Saida                                                              #
# ------------------------------------------------------------------ #

_BORDA = Border(left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin"))


def _formatar(ws, df, money_cols, larguras):
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for c, col in enumerate(df.columns, start=1):
        if col in money_cols:
            for r in range(2, len(df) + 2):
                ws.cell(row=r, column=c).number_format = 'R$ #,##0.00'
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = larguras.get(col, 14)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def gerar_saida(por_tenant: dict, commit: bool) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    modo = "commit" if commit else "dryrun"
    lk, lc = [], []
    for tenant, kits in por_tenant.items():
        for k in kits:
            lk.append({
                "Empresa": tenant, "SKU": k["sku"], "Descricao": k["descricao"],
                "Situacao": k["situacao"], "Custo atual": k["custo_atual"],
                "Custo calculado": k["custo_calculado"],
                "Componentes": k["n_componentes"],
                "Componentes sem custo": k["componentes_sem_custo"],
                "Status": k["status"], "Resultado": k.get("resultado", ""),
                "Custo relido": k.get("custo_relido", ""), "ID": k["id"],
            })
            for c in k["componentes"]:
                lc.append({"Empresa": tenant, **c})
    df_k = pd.DataFrame(lk)
    df_c = pd.DataFrame(lc).rename(columns={
        "kit_sku": "Kit SKU", "kit_descricao": "Kit descricao",
        "componente_id": "Componente ID", "componente_sku": "Componente SKU",
        "componente_descricao": "Componente descricao", "componente_tipo": "Tipo",
        "quantidade": "Qtd", "custo_unitario": "Custo unit.", "subtotal": "Subtotal",
        "origem_custo": "Origem do custo"})
    xlsx = REPORTS_DIR / f"custo_kits_{modo}_{stamp}.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        df_k.to_excel(w, index=False, sheet_name="Kits")
        _formatar(w.sheets["Kits"], df_k, {"Custo atual", "Custo calculado", "Custo relido"},
                  {"SKU": 16, "Descricao": 50, "Componentes sem custo": 30,
                   "Status": 18, "Resultado": 24})
        df_c.to_excel(w, index=False, sheet_name="Componentes")
        _formatar(w.sheets["Componentes"], df_c, {"Custo unit.", "Subtotal"},
                  {"Kit SKU": 16, "Kit descricao": 45, "Componente SKU": 16,
                   "Componente descricao": 45, "Origem do custo": 22})
    csv_path = REPORTS_DIR / f"custo_kits_{modo}_{stamp}.csv"
    df_k.to_csv(csv_path, index=False, sep=";", decimal=",", encoding="utf-8-sig",
                quoting=csv.QUOTE_MINIMAL)
    print(f"\n  Planilha: {xlsx}\n  Log CSV : {csv_path}")
    return xlsx


# ------------------------------------------------------------------ #
#  Main                                                               #
# ------------------------------------------------------------------ #

def processar_tenant(tenant: str, skus: list[str], limite: int | None,
                     commit: bool, aceitar_incompleto: bool, sim: bool) -> list[dict]:
    print(f"\n=== {tenant} ===")
    client = TinyClient(tenant)
    catalogo = listar_catalogo(client)
    kits_ids = [pid for pid, it in catalogo.items()
                if str(it.get("tipo") or "").upper() == "K"
                and str(it.get("situacao") or "").upper() != "E"]
    if skus:
        alvo = set(skus)
        kits_ids = [pid for pid in kits_ids if str(catalogo[pid].get("sku") or "") in alvo]
    kits_ids.sort(key=lambda pid: str(catalogo[pid].get("sku") or ""))
    if limite:
        kits_ids = kits_ids[:limite]
    print(f"  kits a processar: {len(kits_ids)}")

    calc = Calculadora(client, catalogo)
    kits = []
    for i, pid in enumerate(kits_ids, 1):
        try:
            k = calc.calcular(pid)
        except Exception as e:
            it = catalogo[pid]
            k = {"id": pid, "sku": it.get("sku"), "descricao": it.get("descricao"),
                 "situacao": it.get("situacao"), "custo_atual": _custo(it),
                 "custo_calculado": None, "n_componentes": 0,
                 "componentes_sem_custo": "", "componentes": [], "status": f"erro: {e}"}
        kits.append(k)
        if i % 25 == 0 or i == len(kits_ids):
            print(f"  [{i}/{len(kits_ids)}] calculados")

    cont: dict = {}
    for k in kits:
        cont[k["status"]] = cont.get(k["status"], 0) + 1
    print(f"  status: {cont}")
    com_custo_previo = [k for k in kits if (k["custo_atual"] or 0) > 0]
    print(f"  kits que ja tinham custo > 0: {len(com_custo_previo)}")

    pend = pendentes_de_gravacao(kits, aceitar_incompleto)
    print(f"  a gravar: {len(pend)} kit(s), soma R$ {sum(k['custo_calculado'] for k in pend):,.2f}")
    for k in pend[:15]:
        print(f"    {k['sku']:18s} {str(k['custo_atual'] or 0):>8s} -> {k['custo_calculado']:>9,.2f}  "
              f"{k['descricao'][:45]}")
    if len(pend) > 15:
        print(f"    ... (+{len(pend) - 15})")

    if commit and pend:
        if not sim:
            resp = input(f"\n  [{tenant}] Confirma gravar o custo de {len(pend)} kit(s)? (SIM) ")
            if resp.strip().upper() != "SIM":
                print("  cancelado.")
                return kits
        gravar(client, pend)
    elif not commit:
        print("  DRY-RUN: nada gravado. Repita com --commit para aplicar.")
    return kits


def main():
    args = sys.argv[1:]

    def _opt(nome, default=None):
        if nome in args:
            i = args.index(nome)
            v = args[i + 1]
            del args[i:i + 2]
            return v
        return default

    def _flag(nome):
        if nome in args:
            args.remove(nome)
            return True
        return False

    tenants = _opt("--tenants", ",".join(TENANTS_PADRAO)).split(",")
    limite = _opt("--limite")
    limite = int(limite) if limite else None
    commit = _flag("--commit")
    aceitar = _flag("--aceitar-incompleto")
    sim = _flag("--sim")
    skus = [a.strip() for a in args if a.strip()]

    print("=== Custo de kits = soma do custo dos componentes ===")
    print(f"  Empresas: {', '.join(tenants)}   Modo: {'COMMIT' if commit else 'DRY-RUN'}")
    if not validate_tenants(tenants, interactive=False):
        print("Autenticacao nao concluida. Rode src/apps/auth/_reauth.py <tenant>.")
        sys.exit(1)

    por_tenant = {}
    for t in tenants:
        por_tenant[t] = processar_tenant(t, skus, limite, commit, aceitar, sim)
    gerar_saida(por_tenant, commit)

    if commit:
        for t, kits in por_tenant.items():
            res: dict = {}
            for k in kits:
                if "resultado" in k:
                    res[k["resultado"]] = res.get(k["resultado"], 0) + 1
            print(f"  [{t}] resultado da gravacao: {res}")


if __name__ == "__main__":
    main()
