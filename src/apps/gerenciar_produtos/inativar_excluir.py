"""
Inativa ou exclui produtos no Tiny a partir de uma planilha Excel com SKUs.

Uso:
    python inativar_excluir.py <planilha.xlsx> <inativar|excluir> [tenant_id]

Planilha:
    Precisa ter apenas a coluna 'sku' (ou 'codigo'). Linhas vazias sao ignoradas.

Comportamento:
    inativar  -> produto normal: PUT /produtos/{id} com situacao='I'
                 variacao      : marca situacao='I' dentro do array variacoes do pai
                                 (1 PUT por pai)
    excluir   -> produto normal: DELETE /produtos/{id}
                 variacao      : remove a entrada do array variacoes do pai
                                 (1 PUT por pai)

Em ambos os casos faz dry-run com preview e pede confirmacao antes de aplicar.
Salva log CSV em reports/.

Observacao: o Tiny bloqueia exclusao de produtos com vinculos (estoque, pedidos,
notas, kits). Nesses casos use 'inativar'.
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

# reports/ na raiz do projeto (raiz = pasta pai de src/)
REPORTS_DIR = _SRC.parent / "reports"

ACOES_VALIDAS = ("inativar", "excluir")


def ler_skus(caminho: Path) -> list[dict]:
    """Le a planilha e retorna [{linha, sku}, ...]."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip().lower() if c is not None else "" for c in next(rows)]

    col_sku = next((i for i, h in enumerate(header) if h in ("sku", "codigo", "código")), None)
    if col_sku is None:
        raise ValueError(f"Planilha precisa ter coluna SKU. Cabecalho: {header}")

    itens = []
    for i, row in enumerate(rows, start=2):
        if not row or all(c is None for c in row):
            continue
        sku = row[col_sku]
        if sku is None:
            continue
        itens.append({"linha": i, "sku": str(sku).strip()})
    return itens


# ====================================================================== #
#  Planejamento                                                            #
# ====================================================================== #

def montar_preview(client: TinyClient, itens: list[dict], acao: str) -> list[dict]:
    """Para cada SKU monta a operacao a executar.

    Variacoes editam o array do pai; pais sao cacheados em memoria para que
    todas as variacoes do mesmo pai produzam 1 unico PUT.
    """
    plano: list[dict] = []
    parent_cache: dict[int, dict] = {}
    print(f"\n  Buscando {len(itens)} produto(s) no Tiny...")

    for idx, item in enumerate(itens, start=1):
        sku = item["sku"]
        print(f"  [{idx}/{len(itens)}] SKU={sku} ... ", end="", flush=True)

        try:
            resumo = client.buscar_produto_por_sku(sku)
        except Exception as e:
            print(f"ERRO busca ({e})")
            plano.append({**item, "id": None, "tipo": "?", "alvo_id": None,
                          "descricao": None, "situacao_atual": None,
                          "payload": None, "status": f"erro_busca: {e}"})
            continue

        if not resumo:
            print("NAO ENCONTRADO")
            plano.append({**item, "id": None, "tipo": "?", "alvo_id": None,
                          "descricao": None, "situacao_atual": None,
                          "payload": None, "status": "nao_encontrado"})
            continue

        produto_id = resumo.get("id")

        try:
            full = _desempacotar(client.obter_produto(produto_id))
        except Exception as e:
            print(f"ERRO GET ({e})")
            plano.append({**item, "id": produto_id, "tipo": "?", "alvo_id": None,
                          "descricao": None, "situacao_atual": None,
                          "payload": None, "status": f"erro_get: {e}"})
            continue

        tipo_var = (full.get("tipoVariacao") or "").upper()
        descricao = str(full.get("descricao", "")).strip()
        situacao = str(full.get("situacao", "")).strip().upper()

        # -------------------- variacao ---------------------------------- #
        if tipo_var == "V":
            parent_id = (full.get("produtoPai") or {}).get("id")
            if not parent_id:
                print("ERRO (variacao sem produtoPai)")
                plano.append({**item, "id": produto_id, "tipo": "variacao",
                              "alvo_id": None, "descricao": descricao,
                              "situacao_atual": situacao, "payload": None,
                              "status": "erro: variacao sem produtoPai"})
                continue

            if parent_id not in parent_cache:
                try:
                    pai_full = _desempacotar(client.obter_produto(parent_id))
                except Exception as e:
                    print(f"ERRO GET pai ({e})")
                    plano.append({**item, "id": produto_id, "tipo": "variacao",
                                  "alvo_id": parent_id, "descricao": descricao,
                                  "situacao_atual": situacao, "payload": None,
                                  "status": f"erro_get_pai: {e}"})
                    continue
                parent_cache[parent_id] = preparar_payload_put(pai_full)

            pai_payload = parent_cache[parent_id]
            variacoes = pai_payload.get("variacoes") or []
            target = next((v for v in variacoes if v.get("id") == produto_id), None)
            if target is None:
                print("ERRO (variacao ausente no array do pai)")
                plano.append({**item, "id": produto_id, "tipo": "variacao",
                              "alvo_id": parent_id, "descricao": descricao,
                              "situacao_atual": situacao, "payload": None,
                              "status": "erro: variacao ausente no array do pai"})
                continue

            if acao == "inativar":
                if (target.get("situacao") or "").upper() == "I" or situacao == "I":
                    print("OK (ja inativa) [variacao]")
                    plano.append({**item, "id": produto_id, "tipo": "variacao",
                                  "alvo_id": parent_id, "descricao": descricao,
                                  "situacao_atual": situacao, "payload": None,
                                  "status": "sem_mudanca"})
                    continue
                target["situacao"] = "I"
            else:  # excluir
                pai_payload["variacoes"] = [v for v in variacoes if v.get("id") != produto_id]

            print(f"OK [variacao, pai={parent_id}]")
            plano.append({**item, "id": produto_id, "tipo": "variacao",
                          "alvo_id": parent_id, "descricao": descricao,
                          "situacao_atual": situacao, "payload": pai_payload,
                          "status": "pendente"})
            continue

        # -------------------- produto normal ---------------------------- #
        if acao == "inativar":
            if situacao == "I":
                print("OK (ja inativo)")
                plano.append({**item, "id": produto_id, "tipo": "normal",
                              "alvo_id": produto_id, "descricao": descricao,
                              "situacao_atual": situacao, "payload": None,
                              "status": "sem_mudanca"})
                continue
            payload = preparar_payload_put(full)
            payload["situacao"] = "I"
            print("OK")
            plano.append({**item, "id": produto_id, "tipo": "normal",
                          "alvo_id": produto_id, "descricao": descricao,
                          "situacao_atual": situacao, "payload": payload,
                          "status": "pendente"})
        else:  # excluir
            print("OK")
            plano.append({**item, "id": produto_id, "tipo": "normal",
                          "alvo_id": produto_id, "descricao": descricao,
                          "situacao_atual": situacao, "payload": None,
                          "status": "pendente"})

    return plano


# ====================================================================== #
#  Preview                                                                 #
# ====================================================================== #

def imprimir_plano(plano: list[dict], acao: str) -> int:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    sem_mudanca = [p for p in plano if p["status"] == "sem_mudanca"]
    erros = [p for p in plano if p["status"] not in ("pendente", "sem_mudanca")]

    print("\n" + "=" * 80)
    print(f"  PLANO DE {acao.upper()}")
    print("=" * 80)

    if pendentes:
        print(f"\n  {len(pendentes)} produto(s) sera(o) {acao.replace('r', 'do')}(s):\n")
        for p in pendentes:
            extra = f"  [variacao do pai {p['alvo_id']}]" if p.get("tipo") == "variacao" else ""
            print(f"   SKU {p['sku']:>15}  (id={p['id']}){extra}")
            print(f"     descricao: {p['descricao']}")
            if acao == "inativar":
                print(f"     situacao : {p['situacao_atual']} -> I")
            else:
                print(f"     acao     : EXCLUIR")
            print()

    if sem_mudanca:
        print(f"  {len(sem_mudanca)} ja esta(ao) no estado desejado.")

    if erros:
        print(f"\n  {len(erros)} com problemas:")
        for p in erros:
            print(f"   SKU {p['sku']}: {p['status']}")

    print("=" * 80)
    return len(pendentes)


# ====================================================================== #
#  Aplicacao                                                               #
# ====================================================================== #

def aplicar(client: TinyClient, plano: list[dict], acao: str) -> list[dict]:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    pais_pendentes = {p["alvo_id"] for p in pendentes if p.get("tipo") == "variacao"}
    normais = [p for p in pendentes if p.get("tipo") != "variacao"]
    total = len(normais) + len(pais_pendentes)
    print(f"\n  Aplicando {len(pendentes)} mudanca(s) "
          f"({len(normais)} produto(s) + {len(pais_pendentes)} pai(s) de variacao)...")

    feito_por_pai: dict[int, str] = {}
    contador = 0

    for p in pendentes:
        if p.get("tipo") == "variacao":
            parent_id = p["alvo_id"]
            if parent_id in feito_por_pai:
                p["status"] = feito_por_pai[parent_id]
                continue
            contador += 1
            print(f"  [{contador}/{total}] pai_id={parent_id} (variacoes) ... ",
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
            print(f"  [{contador}/{total}] SKU={p['sku']} ... ", end="", flush=True)
            try:
                if acao == "inativar":
                    resp = client.atualizar_produto(p["id"], p["payload"])
                    # Verifica se o Tiny realmente aplicou a mudanca
                    apos = _desempacotar(client.obter_produto(p["id"]))
                    situacao_nova = str(apos.get("situacao", "")).strip().upper()
                    if situacao_nova == "I":
                        p["status"] = "atualizado"
                        print("OK")
                    else:
                        p["status"] = (f"erro_silencioso: PUT aceito mas situacao "
                                       f"continua '{situacao_nova}' (esperado 'I'). "
                                       f"Resposta do PUT: {resp}")
                        print(f"FALHA SILENCIOSA (situacao apos={situacao_nova!r})")
                else:
                    client.excluir_produto(p["id"])
                    p["status"] = "excluido"
                    print("OK")
            except Exception as e:
                p["status"] = f"erro_{acao}: {e}"
                print(f"ERRO ({e})")

    return plano


# ====================================================================== #
#  Log                                                                     #
# ====================================================================== #

def salvar_log(plano: list[dict], tenant_id: str, acao: str) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"{acao}_produtos_{tenant_id}_{stamp}.csv"
    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "linha", "sku", "id", "tipo", "alvo_id", "descricao",
            "situacao_atual", "status",
        ])
        writer.writeheader()
        for p in plano:
            writer.writerow({k: p.get(k, "") for k in writer.fieldnames})
    return arquivo


# ====================================================================== #
#  Main                                                                    #
# ====================================================================== #

def main():
    if len(sys.argv) < 3:
        print("Uso: python gerenciar_produtos.py <planilha.xlsx> <inativar|excluir> [tenant_id]")
        sys.exit(1)

    caminho = Path(sys.argv[1])
    acao = sys.argv[2].strip().lower()
    tenant_id = resolver_tenant(sys.argv[3] if len(sys.argv) > 3 else None)

    if acao not in ACOES_VALIDAS:
        print(f"Acao invalida: {acao}. Use uma de: {', '.join(ACOES_VALIDAS)}")
        sys.exit(1)
    if not caminho.exists():
        print(f"Arquivo nao encontrado: {caminho}")
        sys.exit(1)

    if not validate_tenants([tenant_id]):
        print("Autenticacao nao concluida. Encerrando.")
        sys.exit(1)

    print(f"\n  Acao: {acao.upper()}  |  Tenant: {tenant_id}")
    print(f"  Lendo planilha: {caminho.name}")
    itens = ler_skus(caminho)
    print(f"  {len(itens)} SKU(s) na planilha.")

    if not itens:
        return

    client = TinyClient(tenant_id)
    plano = montar_preview(client, itens, acao)
    pendentes = imprimir_plano(plano, acao)

    if pendentes == 0:
        print("\n  Nada a aplicar.")
        log = salvar_log(plano, tenant_id, acao)
        print(f"  Log salvo em: {log}")
        return

    aviso = "EXCLUSAO E IRREVERSIVEL!" if acao == "excluir" else ""
    resp = input(f"\n  Confirma {acao} de {pendentes} produto(s)? {aviso} (s/N): ").strip().lower()
    if resp != "s":
        print("  Cancelado.")
        log = salvar_log(plano, tenant_id, acao)
        print(f"  Log do dry-run salvo em: {log}")
        return

    aplicar(client, plano, acao)
    log = salvar_log(plano, tenant_id, acao)

    sucesso = sum(1 for p in plano if p["status"] in ("atualizado", "excluido"))
    print(f"\n  Concluido: {sucesso}/{pendentes} processado(s).")
    print(f"  Log salvo em: {log}")


if __name__ == "__main__":
    main()
