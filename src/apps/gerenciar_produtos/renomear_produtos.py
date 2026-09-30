"""
Renomeia produtos no Tiny a partir de uma planilha Excel (SKU + nova descricao).

Uso:
    python renomear_produtos.py <caminho_planilha.xlsx> [tenant_id]

Fluxo:
    1. Le a planilha (colunas: 'sku' e 'descricao' - case-insensitive).
    2. Para cada linha: localiza o produto via SKU e faz GET completo (/produtos/{id}).
    3. Substitui o campo 'descricao' no JSON pelo valor da planilha.
    4. Mostra preview (atual -> novo) e pede confirmacao.
    5. Faz POST /produtos/{id} com o JSON inteiro atualizado.
    6. Salva log em reports/renomear_produtos_<tenant>_<timestamp>.csv.
"""

import csv
import sys
from datetime import datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import openpyxl

from core.auth_tiny import resolver_tenant

from core.auth_tiny import validate_tenants
from core.tiny_client import TinyClient
from core import desempacotar_data as _desempacotar
from core.produtos_utils import preparar_payload_put

REPORTS_DIR = _SRC.parent / "reports"


def ler_planilha(caminho: Path) -> list[dict]:
    """Le a planilha e retorna lista de {sku, descricao_nova}.

    Aceita colunas 'sku'/'codigo' e 'descricao'/'descricao_nova'/'nome' (case-insensitive).
    """
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip().lower() if c is not None else "" for c in next(rows)]

    col_sku = next((i for i, h in enumerate(header) if h in ("sku", "codigo", "código")), None)
    col_desc = next(
        (i for i, h in enumerate(header) if h in ("descricao", "descrição", "descricao_nova", "nome", "nova_descricao")),
        None,
    )

    if col_sku is None or col_desc is None:
        raise ValueError(
            f"Planilha precisa ter colunas SKU e Descricao. Cabecalho encontrado: {header}"
        )

    itens: list[dict] = []
    for i, row in enumerate(rows, start=2):
        if not row or all(c is None for c in row):
            continue
        sku = row[col_sku]
        desc = row[col_desc]
        if sku is None or desc is None:
            print(f"  [linha {i}] ignorada (sku ou descricao vazios)")
            continue
        itens.append({
            "linha": i,
            "sku": str(sku).strip(),
            "descricao_nova": str(desc).strip(),
        })

    return itens


def montar_preview(client: TinyClient, itens: list[dict]) -> list[dict]:
    """Para cada SKU: localiza id, busca JSON completo e prepara o payload atualizado.

    Variacoes (tipoVariacao=='V') sao tratadas via produto pai: o array 'variacoes'
    do pai e modificado e o PUT vai no pai. Pais sao cacheados, entao varias
    variacoes do mesmo pai compartilham o mesmo payload (um unico PUT depois).
    """
    plano: list[dict] = []
    parent_cache: dict[int, dict] = {}  # id_pai -> payload sanitizado e mutavel
    print(f"\n  Buscando {len(itens)} produto(s) no Tiny...")

    for idx, item in enumerate(itens, start=1):
        sku = item["sku"]
        nova = item["descricao_nova"]
        print(f"  [{idx}/{len(itens)}] SKU={sku} ... ", end="", flush=True)

        try:
            resumo = client.buscar_produto_por_sku(sku)
        except Exception as e:
            print(f"ERRO busca ({e})")
            plano.append({**item, "id": None, "atual": None, "payload": None,
                          "tipo": "?", "alvo_id": None, "status": f"erro_busca: {e}"})
            continue

        if not resumo:
            print("NAO ENCONTRADO")
            plano.append({**item, "id": None, "atual": None, "payload": None,
                          "tipo": "?", "alvo_id": None, "status": "nao_encontrado"})
            continue

        produto_id = resumo.get("id")

        # GET do JSON completo do produto (a busca por codigo retorna versao resumida)
        try:
            produto_full = _desempacotar(client.obter_produto(produto_id))
        except Exception as e:
            print(f"ERRO GET ({e})")
            plano.append({**item, "id": produto_id, "atual": None, "payload": None,
                          "tipo": "?", "alvo_id": None, "status": f"erro_get: {e}"})
            continue

        tipo_var = (produto_full.get("tipoVariacao") or "").upper()

        # ---------------- variacao (filha) ---------------------------------- #
        if tipo_var == "V":
            parent_id = (produto_full.get("produtoPai") or {}).get("id")
            if not parent_id:
                print("ERRO (variacao sem produtoPai)")
                plano.append({**item, "id": produto_id, "atual": None, "payload": None,
                              "tipo": "variacao", "alvo_id": None,
                              "status": "erro: variacao sem produtoPai"})
                continue

            if parent_id not in parent_cache:
                try:
                    pai_full = _desempacotar(client.obter_produto(parent_id))
                except Exception as e:
                    print(f"ERRO GET pai ({e})")
                    plano.append({**item, "id": produto_id, "atual": None, "payload": None,
                                  "tipo": "variacao", "alvo_id": parent_id,
                                  "status": f"erro_get_pai: {e}"})
                    continue
                parent_cache[parent_id] = preparar_payload_put(pai_full)

            pai_payload = parent_cache[parent_id]
            target = next((v for v in pai_payload.get("variacoes", [])
                           if v.get("id") == produto_id), None)
            if target is None:
                print("ERRO (variacao nao encontrada no pai)")
                plano.append({**item, "id": produto_id, "atual": None, "payload": None,
                              "tipo": "variacao", "alvo_id": parent_id,
                              "status": "erro: variacao ausente no array do pai"})
                continue

            atual = str(target.get("descricao", "")).strip()
            if atual == nova:
                print("OK (ja esta correto) [variacao]")
                plano.append({**item, "id": produto_id, "atual": atual, "payload": None,
                              "tipo": "variacao", "alvo_id": parent_id,
                              "status": "sem_mudanca"})
                continue

            target["descricao"] = nova  # muta o payload cacheado do pai
            print(f"OK [variacao, pai={parent_id}]")
            plano.append({**item, "id": produto_id, "atual": atual, "payload": pai_payload,
                          "tipo": "variacao", "alvo_id": parent_id, "status": "pendente"})
            continue

        # ---------------- produto comum ------------------------------------- #
        atual = str(produto_full.get("descricao", "")).strip()
        if atual == nova:
            print("OK (ja esta correto)")
            plano.append({**item, "id": produto_id, "atual": atual, "payload": None,
                          "tipo": "normal", "alvo_id": produto_id, "status": "sem_mudanca"})
            continue

        payload = preparar_payload_put(produto_full)
        payload["descricao"] = nova
        print("OK")
        plano.append({**item, "id": produto_id, "atual": atual, "payload": payload,
                      "tipo": "normal", "alvo_id": produto_id, "status": "pendente"})

    return plano


def imprimir_plano(plano: list[dict]) -> tuple[int, int, int]:
    """Mostra o plano e retorna (pendentes, sem_mudanca, erros)."""
    pendentes = [p for p in plano if p["status"] == "pendente"]
    sem_mudanca = [p for p in plano if p["status"] == "sem_mudanca"]
    erros = [p for p in plano if p["status"] not in ("pendente", "sem_mudanca")]

    print("\n" + "=" * 80)
    print("  PLANO DE MUDANCAS")
    print("=" * 80)

    if pendentes:
        print(f"\n  {len(pendentes)} produto(s) sera(o) renomeado(s):\n")
        for p in pendentes:
            extra = ""
            if p.get("tipo") == "variacao":
                extra = f"  [variacao do pai {p.get('alvo_id')}]"
            print(f"   SKU {p['sku']:>15}  (id={p['id']}){extra}")
            print(f"     atual: {p['atual']}")
            print(f"     novo : {p['descricao_nova']}")
            print()

    if sem_mudanca:
        print(f"  {len(sem_mudanca)} ja estao com a descricao correta (sera(o) ignorado(s))")

    if erros:
        print(f"\n  {len(erros)} com problemas:")
        for p in erros:
            print(f"   SKU {p['sku']}: {p['status']}")

    print("=" * 80)
    return len(pendentes), len(sem_mudanca), len(erros)


def aplicar_mudancas(client: TinyClient, plano: list[dict]) -> list[dict]:
    """Aplica os PUTs.

    Produtos normais: 1 PUT cada.
    Variacoes: 1 PUT por produto pai (todas as variacoes do mesmo pai vao juntas
    porque compartilham o mesmo payload em memoria).
    """
    pendentes = [p for p in plano if p["status"] == "pendente"]
    # Conta PUTs reais: 1 por produto normal + 1 por pai unico
    pais_pendentes = {p["alvo_id"] for p in pendentes if p.get("tipo") == "variacao"}
    normais = [p for p in pendentes if p.get("tipo") != "variacao"]
    total_puts = len(normais) + len(pais_pendentes)
    print(f"\n  Aplicando {len(pendentes)} mudanca(s) via PUT "
          f"({len(normais)} produto(s) + {len(pais_pendentes)} pai(s) de variacao)...")

    feito_por_pai: dict[int, str] = {}  # parent_id -> "atualizado" ou "erro: ..."
    contador = 0

    for p in pendentes:
        if p.get("tipo") == "variacao":
            parent_id = p["alvo_id"]
            if parent_id in feito_por_pai:
                # ja foi enviado junto com outra variacao do mesmo pai
                p["status"] = feito_por_pai[parent_id]
                continue
            contador += 1
            print(f"  [{contador}/{total_puts}] pai_id={parent_id} (variacoes) ... ",
                  end="", flush=True)
            try:
                client.atualizar_produto(parent_id, p["payload"])
                feito_por_pai[parent_id] = "atualizado"
                p["status"] = "atualizado"
                print("OK")
            except Exception as e:
                msg = f"erro_put: {e}"
                feito_por_pai[parent_id] = msg
                p["status"] = msg
                print(f"ERRO ({e})")
        else:
            contador += 1
            print(f"  [{contador}/{total_puts}] SKU={p['sku']} ... ", end="", flush=True)
            try:
                client.atualizar_produto(p["id"], p["payload"])
                p["status"] = "atualizado"
                print("OK")
            except Exception as e:
                p["status"] = f"erro_put: {e}"
                print(f"ERRO ({e})")

    return plano


def salvar_log(plano: list[dict], tenant_id: str) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"renomear_produtos_{tenant_id}_{stamp}.csv"

    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["linha", "sku", "id", "tipo", "alvo_id", "atual",
                        "descricao_nova", "status"],
        )
        writer.writeheader()
        for p in plano:
            writer.writerow({
                "linha": p.get("linha", ""),
                "sku": p.get("sku", ""),
                "id": p.get("id", ""),
                "tipo": p.get("tipo", ""),
                "alvo_id": p.get("alvo_id", ""),
                "atual": p.get("atual", ""),
                "descricao_nova": p.get("descricao_nova", ""),
                "status": p.get("status", ""),
            })

    return arquivo


def main():
    if len(sys.argv) < 2:
        print("Uso: python renomear_produtos.py <caminho_planilha.xlsx> [tenant_id]")
        sys.exit(1)

    caminho = Path(sys.argv[1])
    if not caminho.exists():
        print(f"Arquivo nao encontrado: {caminho}")
        sys.exit(1)

    tenant_id = resolver_tenant(sys.argv[2] if len(sys.argv) > 2 else None)

    if not validate_tenants([tenant_id]):
        print("Autenticacao nao concluida. Encerrando.")
        sys.exit(1)

    print(f"\n  Lendo planilha: {caminho.name}")
    itens = ler_planilha(caminho)
    print(f"  {len(itens)} linha(s) com dados validos.")

    if not itens:
        print("Nada a fazer.")
        return

    client = TinyClient(tenant_id)
    plano = montar_preview(client, itens)
    pendentes, _, _ = imprimir_plano(plano)

    if pendentes == 0:
        print("\n  Nenhuma mudanca a aplicar.")
        log = salvar_log(plano, tenant_id)
        print(f"  Log salvo em: {log}")
        return

    resp = input(f"\n  Confirma aplicar {pendentes} mudanca(s)? (s/N): ").strip().lower()
    if resp != "s":
        print("  Cancelado. Nenhuma mudanca aplicada.")
        log = salvar_log(plano, tenant_id)
        print(f"  Log do dry-run salvo em: {log}")
        return

    aplicar_mudancas(client, plano)
    log = salvar_log(plano, tenant_id)

    sucesso = sum(1 for p in plano if p["status"] == "atualizado")
    print(f"\n  Concluido: {sucesso}/{pendentes} atualizado(s).")
    print(f"  Log salvo em: {log}")


if __name__ == "__main__":
    main()
