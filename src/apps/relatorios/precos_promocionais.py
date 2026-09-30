"""
Snapshot do preco promocional de TODOS os produtos cadastrados no Tiny.

Le a listagem paginada de /produtos — que ja traz `precos.preco` e
`precos.precoPromocional` prontos, sem custar o detalhe — e gera um xlsx com uma
aba por empresa mais uma aba de resumo. Sai por ~15 requests por empresa
(limit=100): e o jeito barato de auditar uma acao de preco depois de grava-la.

Uso:
    python precos_promocionais.py                  # todas as empresas
    python precos_promocionais.py <tenant>         # so uma
    python precos_promocionais.py --so-ativos      # descarta situacao 'E' e 'I'
    python precos_promocionais.py --saida reports/meu_arquivo.xlsx

Decisoes de negocio acordadas:
    - Traz TUDO por padrao: o gemeo com situacao='E' (exclusao de produto e logica,
      nao existe DELETE de produto usado em pedido), os inativos e as variacoes.
      Esconder linha por padrao faria o total da planilha nao bater com o total que
      a propria API informa. Quem quer so o cadastro vivo usa --so-ativos.
    - `precoPromocional = 0` quer dizer SEM promocao ativa — o produto vende pelo
      `precos.preco`. E esse o estado em que precos/atualizar_precos.py --zerar deixa
      o produto ao encerrar uma acao; zero aqui e normalidade, nao preco de brinde.
    - Desconto so e calculado quando ha promocional > 0 E preco > 0. Nos demais casos
      a celula fica VAZIA, nunca zero — zero seria lido como "sem desconto".
    - Promocional MAIOR que o preco de venda nao e erro de leitura: acontece, e sai
      marcado na coluna 'Anomalia' e contado no Resumo, para nao passar em silencio.
"""

import sys
from datetime import datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
for _p in Path(__file__).resolve().parents:
    if (_p / "core").is_dir():
        _SRC = _p
        sys.path.insert(0, str(_p))
        break

import pandas as pd

from core.auth_tiny import validate_tenants
from core.tiny_client import TinyClient
from core import desempacotar_data as _desempacotar

REPORTS_DIR = _SRC.parent / "reports"
from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS_PADRAO = get_all_tenant_ids()  # chaves de data/tenants.json

SITUACOES = {"A": "Ativo", "I": "Inativo", "E": "Excluido"}


# --------------------------------------------------------------------------- #
#  Coleta                                                                      #
# --------------------------------------------------------------------------- #

def coletar(client: TinyClient, tenant: str) -> list[dict]:
    """Pagina /produtos ate o fim. A listagem ja traz precos; o detalhe nao e preciso."""
    produtos: list[dict] = []
    offset = 0
    total = None

    while True:
        resp = _desempacotar(client.listar_produtos(offset=offset))
        itens = resp.get("itens") or []
        if total is None:
            total = (resp.get("paginacao") or {}).get("total")
            print(f"  [{tenant}] {total} produto(s) no cadastro; baixando...")
        produtos.extend(itens)
        if not itens or (total is not None and len(produtos) >= total):
            break
        offset += len(itens)

    print(f"  [{tenant}] {len(produtos)} linha(s) coletada(s).")
    return produtos


def _preco(produto: dict, campo: str) -> float | None:
    precos = produto.get("precos")
    if not isinstance(precos, dict):
        return None
    valor = precos.get(campo)
    try:
        return float(valor) if valor is not None else None
    except (TypeError, ValueError):
        return None


def montar_linhas(produtos: list[dict]) -> pd.DataFrame:
    linhas = []
    for p in produtos:
        preco = _preco(p, "preco")
        promo = _preco(p, "precoPromocional")
        situacao = str(p.get("situacao") or "").strip().upper()
        em_promo = bool(promo and promo > 0)

        desconto_rs = desconto_pct = None
        if em_promo and preco and preco > 0:
            desconto_rs = round(preco - promo, 2)
            desconto_pct = round((preco - promo) / preco * 100, 2)

        anomalia = ""
        if em_promo and preco and promo > preco:
            anomalia = "promocional > preco"
        elif em_promo and not preco:
            anomalia = "promocional sem preco de venda"

        linhas.append({
            "SKU": str(p.get("sku") or "").strip(),
            "Descricao": str(p.get("descricao") or "").strip(),
            "Tipo": str(p.get("tipo") or "").strip(),
            "Tipo variacao": str(p.get("tipoVariacao") or "").strip(),
            "Situacao": SITUACOES.get(situacao, situacao),
            "Unidade": str(p.get("unidade") or "").strip(),
            "GTIN": str(p.get("gtin") or "").strip(),
            "Preco de venda": preco,
            "Preco promocional": promo,
            "Em promocao": "sim" if em_promo else "nao",
            "Desconto (R$)": desconto_rs,
            "Desconto (%)": desconto_pct,
            "Anomalia": anomalia,
            "ID": p.get("id"),
            "Data alteracao": str(p.get("dataAlteracao") or "").strip(),
        })

    df = pd.DataFrame(linhas)
    return df.sort_values(["Situacao", "SKU"], kind="stable").reset_index(drop=True)


def resumir(df: pd.DataFrame, tenant: str) -> dict:
    vivos = df[df["Situacao"] == "Ativo"]
    em_promo = df[df["Em promocao"] == "sim"]
    return {
        "Empresa": tenant,
        "Produtos (todos)": len(df),
        "Ativos": len(vivos),
        "Inativos": int((df["Situacao"] == "Inativo").sum()),
        "Excluidos (logico)": int((df["Situacao"] == "Excluido").sum()),
        "Variacoes": int((df["Tipo variacao"] == "V").sum()),
        "Com promocional > 0": len(em_promo),
        "Com promocional > 0 (so ativos)": int((vivos["Em promocao"] == "sim").sum()),
        "Sem promocional": int((df["Em promocao"] == "nao").sum()),
        "Desconto medio (%)": (round(em_promo["Desconto (%)"].mean(), 2)
                               if not em_promo["Desconto (%)"].dropna().empty else None),
        "Anomalias": int((df["Anomalia"] != "").sum()),
    }


# --------------------------------------------------------------------------- #
#  Planilha                                                                    #
# --------------------------------------------------------------------------- #

def salvar(abas: dict[str, pd.DataFrame], resumo: list[dict], saida: Path) -> Path:
    saida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)
        for nome, df in abas.items():
            df.to_excel(writer, sheet_name=nome[:31], index=False)

        for _nome, ws in writer.sheets.items():
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                largura = max((len(str(c.value)) for c in col[:400] if c.value), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(largura + 2, 10), 52)
    return saida


# --------------------------------------------------------------------------- #
#  Main                                                                        #
# --------------------------------------------------------------------------- #

def main() -> None:
    argumentos = sys.argv[1:]
    flags = [a for a in argumentos if a.startswith("--")]
    so_ativos = "--so-ativos" in flags

    def valor_de(nome: str) -> str | None:
        for f in flags:
            if f.startswith(f"{nome}="):
                return f.split("=", 1)[1]
        if nome in flags:
            i = argumentos.index(nome)
            if i + 1 < len(argumentos) and not argumentos[i + 1].startswith("--"):
                return argumentos[i + 1]
        return None

    saida_txt = valor_de("--saida")
    posicionais = [a for a in argumentos if not a.startswith("--") and a != saida_txt]
    arg_tenant = (posicionais[0] if posicionais else "todas").strip().lower()
    tenants = (list(TENANTS_PADRAO) if arg_tenant in ("todas", "ambas", "all", "")
               else [t.strip() for t in arg_tenant.split(",") if t.strip()])

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    saida = Path(saida_txt) if saida_txt else (
        REPORTS_DIR / f"precos_promocionais_{'_'.join(tenants)}_{stamp}.xlsx")

    print(f"\n  Empresas: {', '.join(tenants)}")
    print("  Escopo  : " + ("so cadastro ativo" if so_ativos
                            else "TUDO (inclui inativo e excluido logico)"))

    if not validate_tenants(tenants, interactive=False):
        print("  Autenticacao nao concluida. Rode src/apps/auth/_reauth.py <tenant>.")
        sys.exit(1)

    abas: dict[str, pd.DataFrame] = {}
    resumo: list[dict] = []
    for tenant in tenants:
        df = montar_linhas(coletar(TinyClient(tenant), tenant))
        if so_ativos:
            df = df[df["Situacao"] == "Ativo"].reset_index(drop=True)
        resumo.append(resumir(df, tenant))
        abas[tenant] = df

    arquivo = salvar(abas, resumo, saida)

    print("\n" + "=" * 78)
    for r in resumo:
        print(f"  {r['Empresa']:<10} {r['Produtos (todos)']:>5} produtos | "
              f"{r['Com promocional > 0']:>4} com promocional "
              f"({r['Com promocional > 0 (so ativos)']} ativos) | "
              f"{r['Anomalias']} anomalia(s)")
    print("=" * 78)
    print(f"\n  Planilha: {arquivo}")


if __name__ == "__main__":
    main()
