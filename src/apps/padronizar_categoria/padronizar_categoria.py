"""
Padroniza a CATEGORIA dos produtos no Tiny a partir de uma planilha Excel.

Planilha (colunas, case-insensitive):
    sku          -> codigo do produto (obrigatorio)
    descricao    -> apenas informativo no preview/log (opcional)
    categoria    -> nome da categoria pai          (obrigatorio)
    subcategoria -> nome da subcategoria (filha)    (opcional)

Uso:
    python padronizar_categoria.py <planilha.xlsx> [tenant_id]

Como funciona:
    1. Busca a arvore de categorias do Tiny (GET /categorias/todas) e resolve
       cada par (categoria, subcategoria) da planilha para o id da categoria.
       - categoria sem subcategoria  -> categoria raiz com aquele nome
       - categoria + subcategoria    -> filha 'subcategoria' sob a pai 'categoria'
       Categoria inexistente e REPORTADA e PULADA (nada e criado).
    2. Para cada SKU: GET do produto completo, troca somente o campo 'categoria'
       por {id: <id resolvido>} e devolve via PUT /produtos/{id}.
       Variacoes (tipoVariacao='V') tem a categoria aplicada no PRODUTO PAI
       (1 PUT por pai; a categoria e atributo do produto, nao da variacao).
    3. Dry-run com preview + confirmacao. Verifica pos-PUT se a categoria.id
       realmente mudou. Log CSV em reports/.
"""

import csv
import sys
import time
from datetime import datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import openpyxl
import requests

from core.auth_tiny import resolver_tenant

from core import (
    validate_tenants,
    TinyClient,
    desempacotar_data,
    preparar_payload_put,
    achatar_categorias,
    ResolvedorCategoria,
)

REPORTS_DIR = _SRC.parent / "reports"


# ====================================================================== #
#  Planilha                                                                #
# ====================================================================== #

def ler_planilha(caminho: Path) -> list[dict]:
    """Le a planilha e retorna [{linha, sku, descricao, categoria, subcategoria}]."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip().lower() if c is not None else "" for c in next(rows)]

    def achar(*nomes):
        return next((i for i, h in enumerate(header) if h in nomes), None)

    col_sku = achar("sku", "codigo", "código")
    col_desc = achar("descricao", "descrição", "nome")
    col_cat = achar("categoria", "categoria_pai")
    col_sub = achar("subcategoria", "sub categoria", "sub-categoria")

    if col_sku is None or col_cat is None:
        raise ValueError(
            f"Planilha precisa ter pelo menos as colunas 'sku' e 'categoria'. "
            f"Cabecalho encontrado: {header}"
        )

    def val(row, idx):
        if idx is None or idx >= len(row) or row[idx] is None:
            return ""
        return str(row[idx]).strip()

    itens = []
    for i, row in enumerate(rows, start=2):
        if not row or all(c is None for c in row):
            continue
        sku = val(row, col_sku)
        if not sku:
            continue
        itens.append({
            "linha": i,
            "sku": sku,
            "descricao_plan": val(row, col_desc),
            "categoria": val(row, col_cat),
            "subcategoria": val(row, col_sub),
        })
    return itens


# ====================================================================== #
#  Criacao de categorias faltantes (com preview)                          #
# ====================================================================== #

def planejar_criacao(itens: list[dict], resolvedor: ResolvedorCategoria) -> dict:
    """Descobre quais categorias precisam ser criadas para cobrir a planilha.

    Retorna {'raizes': [nome,...], 'filhas': [(pai, filha), ...]} (sem duplicar e
    apenas o que ainda nao existe). Pares sem categoria sao ignorados aqui.
    """
    raizes_necessarias: dict[str, str] = {}   # norm -> nome original (para exibir)
    filhas_necessarias: dict[tuple, tuple] = {}  # (npai,nfilha) -> (pai, filha)

    from core import norm_cat
    for it in itens:
        cat, sub = it["categoria"].strip(), it["subcategoria"].strip()
        if not cat:
            continue
        raizes_necessarias.setdefault(norm_cat(cat), cat)
        if sub:
            filhas_necessarias.setdefault((norm_cat(cat), norm_cat(sub)), (cat, sub))

    # raizes que ainda nao existem
    raizes_criar = [nome for nkey, nome in raizes_necessarias.items()
                    if resolvedor.id_raiz(nome) is None]
    # filhas que ainda nao existem (sob o pai correspondente)
    filhas_criar = [(pai, filha) for (_pk, _fk), (pai, filha) in filhas_necessarias.items()
                    if not resolvedor.existe_filha(pai, filha)]

    return {"raizes": sorted(raizes_criar, key=str.casefold),
            "filhas": sorted(filhas_criar, key=lambda x: (x[0].casefold(), x[1].casefold()))}


def imprimir_plano_criacao(plano_cri: dict) -> int:
    raizes, filhas = plano_cri["raizes"], plano_cri["filhas"]
    total = len(raizes) + len(filhas)
    print("\n" + "=" * 80)
    print("  CATEGORIAS A CRIAR")
    print("=" * 80)
    if raizes:
        print(f"\n  {len(raizes)} categoria(s) RAIZ nova(s):")
        for nome in raizes:
            print(f"    + {nome}")
    if filhas:
        print(f"\n  {len(filhas)} SUBcategoria(s) nova(s):")
        for pai, filha in filhas:
            print(f"    + {pai} >> {filha}")
    print("\n" + "=" * 80)
    print(f"  Total a criar: {total}")
    print("=" * 80)
    return total


def criar_categorias(client: TinyClient, plano_cri: dict,
                     resolvedor: ResolvedorCategoria) -> list[dict]:
    """Cria raizes (primeiro) e depois filhas. Re-busca a arvore ao final para
    atualizar os ids no resolvedor. Retorna log das criacoes.

    Resiliencia a throttle: o Tiny pode devolver 403 transitorio ao criar muitas
    categorias em rajada. Como a permissao ja foi validada (criacoes anteriores
    funcionaram), tratamos 403 como temporario com backoff e retry.
    """
    log: list[dict] = []
    raizes, filhas = plano_cri["raizes"], plano_cri["filhas"]

    def _criar(descricao, pai_id=None) -> str:
        """Tenta criar com backoff em 403/throttle. Retorna 'criada' ou 'erro: ...'."""
        esperas = [15, 30, 60, 90]  # segundos entre tentativas
        for tentativa in range(len(esperas) + 1):
            try:
                client.criar_categoria(descricao, id_categoria_pai=pai_id)
                return "criada"
            except requests.exceptions.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else 0
                if status == 403 and tentativa < len(esperas):
                    espera = esperas[tentativa]
                    print(f"[403 throttle, aguardando {espera}s] ", end="", flush=True)
                    time.sleep(espera)
                    continue
                return f"erro: {exc}"
            except Exception as exc:
                return f"erro: {exc}"
        return "erro: 403 persistente apos retries"

    # 1) raizes
    print(f"\n  Criando {len(raizes)} categoria(s) raiz...")
    for i, nome in enumerate(raizes, 1):
        print(f"  [{i}/{len(raizes)}] {nome} ... ", end="", flush=True)
        status = _criar(nome)
        print("OK" if status == "criada" else f"ERRO ({status})")
        log.append({"tipo": "raiz", "nome": nome, "pai": "", "status": status})

    # re-busca a arvore para obter os ids das raizes recem-criadas
    resolvedor.atualizar(achatar_categorias(client.listar_categorias_arvore()))

    # 2) filhas (precisam do id do pai)
    print(f"\n  Criando {len(filhas)} subcategoria(s)...")
    for i, (pai, filha) in enumerate(filhas, 1):
        print(f"  [{i}/{len(filhas)}] {pai} >> {filha} ... ", end="", flush=True)
        pai_id = resolvedor.id_raiz(pai)
        if pai_id is None:
            print("ERRO (pai inexistente)")
            log.append({"tipo": "filha", "nome": filha, "pai": pai,
                        "status": "erro: pai nao encontrado apos criacao"})
            continue
        status = _criar(filha, pai_id=pai_id)
        print("OK" if status == "criada" else f"ERRO ({status})")
        log.append({"tipo": "filha", "nome": filha, "pai": pai, "status": status})

    # re-busca final para refletir todas as filhas criadas
    resolvedor.atualizar(achatar_categorias(client.listar_categorias_arvore()))
    return log


# ====================================================================== #
#  Planejamento                                                            #
# ====================================================================== #

def montar_preview(client: TinyClient, itens: list[dict],
                   resolvedor: ResolvedorCategoria) -> list[dict]:
    """Resolve categoria-alvo e prepara o payload para cada SKU."""
    plano: list[dict] = []
    parent_cache: dict[int, dict] = {}   # id_pai -> payload mutavel
    parent_target: dict[int, int] = {}   # id_pai -> id categoria ja planejado
    print(f"\n  Resolvendo {len(itens)} produto(s)...")

    for idx, item in enumerate(itens, start=1):
        sku = item["sku"]
        print(f"  [{idx}/{len(itens)}] SKU={sku} ... ", end="", flush=True)

        # 1) resolve a categoria-alvo (nome -> id)
        res = resolvedor.resolver(item["categoria"], item["subcategoria"])
        cat_id = res["id"]
        cat_caminho = res["caminho"]
        if cat_id is None:
            print(res["erro"])
            plano.append({**item, "id": None, "tipo": "?", "alvo_id": None,
                          "categoria_id": None, "categoria_caminho": None,
                          "categoria_atual": None, "payload": None,
                          "status": res["erro"]})
            continue

        # 2) localiza o produto e busca o JSON completo
        try:
            resumo = client.buscar_produto_por_sku(sku)
        except Exception as e:
            print(f"ERRO busca ({e})")
            plano.append({**item, "id": None, "tipo": "?", "alvo_id": None,
                          "categoria_id": cat_id, "categoria_caminho": cat_caminho,
                          "categoria_atual": None, "payload": None,
                          "status": f"erro_busca: {e}"})
            continue
        if not resumo:
            print("NAO ENCONTRADO")
            plano.append({**item, "id": None, "tipo": "?", "alvo_id": None,
                          "categoria_id": cat_id, "categoria_caminho": cat_caminho,
                          "categoria_atual": None, "payload": None,
                          "status": "produto_nao_encontrado"})
            continue

        produto_id = resumo.get("id")
        try:
            full = desempacotar_data(client.obter_produto(produto_id))
        except Exception as e:
            print(f"ERRO GET ({e})")
            plano.append({**item, "id": produto_id, "tipo": "?", "alvo_id": None,
                          "categoria_id": cat_id, "categoria_caminho": cat_caminho,
                          "categoria_atual": None, "payload": None,
                          "status": f"erro_get: {e}"})
            continue

        tipo_var = (full.get("tipoVariacao") or "").upper()

        # 3a) variacao -> categoria aplicada no PAI
        if tipo_var == "V":
            parent_id = (full.get("produtoPai") or {}).get("id")
            if not parent_id:
                print("ERRO (variacao sem produtoPai)")
                plano.append({**item, "id": produto_id, "tipo": "variacao",
                              "alvo_id": None, "categoria_id": cat_id,
                              "categoria_caminho": cat_caminho, "categoria_atual": None,
                              "payload": None, "status": "erro: variacao sem produtoPai"})
                continue

            # conflito: outra linha ja planejou categoria diferente para o mesmo pai
            if parent_id in parent_target and parent_target[parent_id] != cat_id:
                print(f"CONFLITO (pai {parent_id} ja planejado p/ outra categoria)")
                plano.append({**item, "id": produto_id, "tipo": "variacao",
                              "alvo_id": parent_id, "categoria_id": cat_id,
                              "categoria_caminho": cat_caminho, "categoria_atual": None,
                              "payload": None,
                              "status": f"erro_conflito: pai {parent_id} mapeado p/ "
                                        f"categoria {parent_target[parent_id]} e {cat_id}"})
                continue

            if parent_id not in parent_cache:
                try:
                    pai_full = desempacotar_data(client.obter_produto(parent_id))
                except Exception as e:
                    print(f"ERRO GET pai ({e})")
                    plano.append({**item, "id": produto_id, "tipo": "variacao",
                                  "alvo_id": parent_id, "categoria_id": cat_id,
                                  "categoria_caminho": cat_caminho, "categoria_atual": None,
                                  "payload": None, "status": f"erro_get_pai: {e}"})
                    continue
                parent_cache[parent_id] = preparar_payload_put(pai_full)

            pai_payload = parent_cache[parent_id]
            atual_id = (pai_payload.get("categoria") or {}).get("id")
            if atual_id == cat_id:
                print("OK (ja correta) [variacao]")
                plano.append({**item, "id": produto_id, "tipo": "variacao",
                              "alvo_id": parent_id, "categoria_id": cat_id,
                              "categoria_caminho": cat_caminho, "categoria_atual": atual_id,
                              "payload": None, "status": "sem_mudanca"})
                continue

            pai_payload["categoria"] = {"id": cat_id}
            parent_target[parent_id] = cat_id
            print(f"OK [variacao, pai={parent_id}]")
            plano.append({**item, "id": produto_id, "tipo": "variacao",
                          "alvo_id": parent_id, "categoria_id": cat_id,
                          "categoria_caminho": cat_caminho, "categoria_atual": atual_id,
                          "payload": pai_payload, "status": "pendente"})
            continue

        # 3b) produto normal
        atual_id = (full.get("categoria") or {}).get("id")
        if atual_id == cat_id:
            print("OK (ja correta)")
            plano.append({**item, "id": produto_id, "tipo": "normal",
                          "alvo_id": produto_id, "categoria_id": cat_id,
                          "categoria_caminho": cat_caminho, "categoria_atual": atual_id,
                          "payload": None, "status": "sem_mudanca"})
            continue

        payload = preparar_payload_put(full)
        payload["categoria"] = {"id": cat_id}
        print("OK")
        plano.append({**item, "id": produto_id, "tipo": "normal",
                      "alvo_id": produto_id, "categoria_id": cat_id,
                      "categoria_caminho": cat_caminho, "categoria_atual": atual_id,
                      "payload": payload, "status": "pendente"})

    return plano


# ====================================================================== #
#  Preview                                                                 #
# ====================================================================== #

def imprimir_plano(plano: list[dict]) -> int:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    sem_mudanca = [p for p in plano if p["status"] == "sem_mudanca"]
    erros = [p for p in plano if p["status"] not in ("pendente", "sem_mudanca")]

    print("\n" + "=" * 80)
    print("  PLANO DE CATEGORIZACAO")
    print("=" * 80)

    if pendentes:
        print(f"\n  {len(pendentes)} produto(s) tera(o) a categoria atualizada:\n")
        for p in pendentes:
            extra = f"  [variacao do pai {p['alvo_id']}]" if p.get("tipo") == "variacao" else ""
            print(f"   SKU {p['sku']:>18}  (id={p['id']}){extra}")
            if p.get("descricao_plan"):
                print(f"     produto  : {p['descricao_plan']}")
            print(f"     categoria: id {p['categoria_atual']} -> id {p['categoria_id']}  "
                  f"({p['categoria_caminho']})")
            print()

    if sem_mudanca:
        print(f"  {len(sem_mudanca)} ja esta(ao) na categoria correta.")

    if erros:
        print(f"\n  {len(erros)} com problemas:")
        for p in erros:
            print(f"   SKU {p['sku']}: {p['status']}")

    print("=" * 80)
    return len(pendentes)


# ====================================================================== #
#  Aplicacao                                                               #
# ====================================================================== #

def aplicar(client: TinyClient, plano: list[dict]) -> list[dict]:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    pais = {p["alvo_id"] for p in pendentes if p.get("tipo") == "variacao"}
    normais = [p for p in pendentes if p.get("tipo") != "variacao"]
    total = len(normais) + len(pais)
    print(f"\n  Aplicando via PUT ({len(normais)} produto(s) + {len(pais)} pai(s) de variacao)...")

    feito_por_pai: dict[int, str] = {}
    contador = 0

    def _verificar(alvo_id, esperado_id) -> bool:
        """Confirma pos-PUT que a categoria.id do alvo virou o esperado."""
        try:
            apos = desempacotar_data(client.obter_produto(alvo_id))
            return (apos.get("categoria") or {}).get("id") == esperado_id
        except Exception:
            return False

    for p in pendentes:
        if p.get("tipo") == "variacao":
            parent_id = p["alvo_id"]
            if parent_id in feito_por_pai:
                p["status"] = feito_por_pai[parent_id]
                continue
            contador += 1
            print(f"  [{contador}/{total}] pai_id={parent_id} (variacoes) ... ", end="", flush=True)
            try:
                client.atualizar_produto(parent_id, p["payload"])
                if _verificar(parent_id, p["categoria_id"]):
                    feito_por_pai[parent_id] = "atualizado"
                    print("OK")
                else:
                    feito_por_pai[parent_id] = "erro_silencioso: categoria nao persistiu"
                    print("FALHA SILENCIOSA")
                p["status"] = feito_por_pai[parent_id]
            except Exception as e:
                feito_por_pai[parent_id] = f"erro_put: {e}"
                p["status"] = feito_por_pai[parent_id]
                print(f"ERRO ({e})")
        else:
            contador += 1
            print(f"  [{contador}/{total}] SKU={p['sku']} ... ", end="", flush=True)
            try:
                client.atualizar_produto(p["id"], p["payload"])
                if _verificar(p["id"], p["categoria_id"]):
                    p["status"] = "atualizado"
                    print("OK")
                else:
                    p["status"] = "erro_silencioso: categoria nao persistiu"
                    print("FALHA SILENCIOSA")
            except Exception as e:
                p["status"] = f"erro_put: {e}"
                print(f"ERRO ({e})")

    return plano


# ====================================================================== #
#  Exclusao das categorias antigas (fora da taxonomia da planilha)         #
# ====================================================================== #

def _profundidade(no: dict) -> int:
    """Nivel do no na arvore (raiz=1), a partir do caminho 'A >> B >> C'."""
    return (no.get("caminho") or "").count(" >> ") + 1


def planejar_exclusao(itens: list[dict], flat: list[dict]) -> dict:
    """Decide quais categorias MANTER (as que a planilha define) e quais APAGAR.

    Mantem:
      - raiz cujo nome e uma 'categoria' da planilha
      - filha (pai, nome) que casa com um par (categoria, subcategoria) da planilha
    Apaga todo o resto. Ordena a exclusao por profundidade DECRESCENTE
    (folhas primeiro), para nao tentar apagar um pai antes das filhas.
    """
    from core import norm_cat
    R = {norm_cat(it["categoria"]) for it in itens if it["categoria"].strip()}
    L = {(norm_cat(it["categoria"]), norm_cat(it["subcategoria"]))
         for it in itens if it["subcategoria"].strip()}

    manter, apagar = [], []
    for no in flat:
        if no["pai_nome"] is None and norm_cat(no["nome"]) in R:
            manter.append(no)
        elif (norm_cat(no["pai_nome"] or ""), norm_cat(no["nome"])) in L:
            manter.append(no)
        else:
            apagar.append(no)

    apagar.sort(key=lambda n: _profundidade(n), reverse=True)
    return {"manter": manter, "apagar": apagar}


def imprimir_plano_exclusao(plano_exc: dict) -> int:
    manter, apagar = plano_exc["manter"], plano_exc["apagar"]
    print("\n" + "=" * 80)
    print("  CATEGORIAS ANTIGAS A APAGAR")
    print("=" * 80)
    print(f"\n  MANTER: {len(manter)} categoria(s) (as definidas pela planilha)")
    print(f"  APAGAR: {len(apagar)} categoria(s) que NAO estao na planilha:\n")
    for no in apagar:
        print(f"    - [id {no['id']}] {no['caminho']}")
    print("\n" + "=" * 80)
    print(f"  Total a apagar: {len(apagar)}")
    print("  (categorias com produtos vinculados — inclusive inativos — podem")
    print("   ser recusadas pelo Tiny; nesse caso sao puladas e reportadas.)")
    print("=" * 80)
    return len(apagar)


def excluir_categorias(client: TinyClient, plano_exc: dict) -> list[dict]:
    """Apaga as categorias do plano (folhas primeiro). Tolera falhas (vinculos)."""
    apagar = plano_exc["apagar"]
    print(f"\n  Apagando {len(apagar)} categoria(s) antiga(s)...")
    log: list[dict] = []
    for i, no in enumerate(apagar, 1):
        print(f"  [{i}/{len(apagar)}] [id {no['id']}] {no['caminho']} ... ", end="", flush=True)
        try:
            client.excluir_categoria(no["id"])
            print("OK")
            log.append({"id": no["id"], "caminho": no["caminho"], "status": "apagada"})
        except Exception as e:
            print(f"PULADA ({e})")
            log.append({"id": no["id"], "caminho": no["caminho"], "status": f"erro: {e}"})
    return log


# ====================================================================== #
#  Log                                                                     #
# ====================================================================== #

def salvar_log_exclusao(log_exc: list[dict], tenant_id: str) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"categorias_apagadas_{tenant_id}_{stamp}.csv"
    campos = ["id", "caminho", "status"]
    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        for r in log_exc:
            writer.writerow({k: r.get(k, "") for k in campos})
    print(f"  Log das exclusoes salvo em: {arquivo}")
    return arquivo


def salvar_log_criacao(log_cri: list[dict], tenant_id: str) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"categorias_criadas_{tenant_id}_{stamp}.csv"
    campos = ["tipo", "pai", "nome", "status"]
    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        for r in log_cri:
            writer.writerow({k: r.get(k, "") for k in campos})
    print(f"  Log das criacoes salvo em: {arquivo}")
    return arquivo


def salvar_log(plano: list[dict], tenant_id: str) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"padronizar_categoria_{tenant_id}_{stamp}.csv"
    campos = ["linha", "sku", "descricao_plan", "id", "tipo", "alvo_id",
              "categoria", "subcategoria", "categoria_caminho",
              "categoria_atual", "categoria_id", "status"]
    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        for p in plano:
            writer.writerow({k: p.get(k, "") for k in campos})
    return arquivo


# ====================================================================== #
#  Main                                                                    #
# ====================================================================== #

def main():
    if len(sys.argv) < 2:
        print("Uso: python padronizar_categoria.py <planilha.xlsx> [tenant_id]")
        sys.exit(1)

    caminho = Path(sys.argv[1])
    if not caminho.exists():
        print(f"Arquivo nao encontrado: {caminho}")
        sys.exit(1)
    tenant_id = resolver_tenant(sys.argv[2] if len(sys.argv) > 2 else None)

    if not validate_tenants([tenant_id]):
        print("Autenticacao nao concluida. Encerrando.")
        sys.exit(1)

    print(f"\n  Tenant: {tenant_id}")
    print(f"  Lendo planilha: {caminho.name}")
    itens = ler_planilha(caminho)
    print(f"  {len(itens)} linha(s) com dados validos.")
    if not itens:
        return

    client = TinyClient(tenant_id)

    print("  Carregando arvore de categorias do Tiny...")
    flat = achatar_categorias(client.listar_categorias_arvore())
    print(f"  {len(flat)} categoria(s) na conta.")
    resolvedor = ResolvedorCategoria(flat)

    # --- FASE 1: criar categorias faltantes (com preview + confirmacao) --- #
    plano_cri = planejar_criacao(itens, resolvedor)
    if plano_cri["raizes"] or plano_cri["filhas"]:
        total_cri = imprimir_plano_criacao(plano_cri)
        resp = input(f"\n  Criar essas {total_cri} categoria(s) no Tiny? (s/N): ").strip().lower()
        if resp == "s":
            log_cri = criar_categorias(client, plano_cri, resolvedor)
            salvar_log_criacao(log_cri, tenant_id)
            criadas = sum(1 for r in log_cri if r["status"] == "criada")
            print(f"\n  {criadas}/{len(log_cri)} categoria(s) criada(s).")
        else:
            print("  Criacao pulada — produtos cujas categorias nao existem serao reportados.")
    else:
        print("  Todas as categorias da planilha ja existem no Tiny.")

    # --- FASE 2: categorizar os produtos --- #
    plano = montar_preview(client, itens, resolvedor)
    pendentes = imprimir_plano(plano)

    if pendentes > 0:
        resp = input(f"\n  Confirma atualizar a categoria de {pendentes} produto(s)? (s/N): ").strip().lower()
        if resp != "s":
            print("  Cancelado na categorizacao. NAO vou apagar categorias antigas")
            print("  (apagar antes de reatribuir os produtos os deixaria sem categoria).")
            salvar_log(plano, tenant_id)
            return
        aplicar(client, plano)
        sucesso = sum(1 for p in plano if p["status"] == "atualizado")
        print(f"\n  Categorizacao: {sucesso}/{pendentes} atualizado(s).")
    else:
        print("\n  Produtos ja estao todos na categoria correta.")
    salvar_log(plano, tenant_id)

    # --- FASE 3: apagar as categorias antigas (fora da taxonomia da planilha) --- #
    plano_exc = planejar_exclusao(itens, resolvedor.flat)
    if not plano_exc["apagar"]:
        print("\n  Nenhuma categoria antiga para apagar — arvore ja esta limpa.")
        return

    total_exc = imprimir_plano_exclusao(plano_exc)
    print("\n  ATENCAO: apagar categorias e IRREVERSIVEL.")
    resp = input(f"  Apagar essas {total_exc} categoria(s) antiga(s)? (s/N): ").strip().lower()
    if resp != "s":
        print("  Exclusao pulada. As categorias antigas foram mantidas.")
        return

    log_exc = excluir_categorias(client, plano_exc)
    salvar_log_exclusao(log_exc, tenant_id)
    apagadas = sum(1 for r in log_exc if r["status"] == "apagada")
    puladas = len(log_exc) - apagadas
    print(f"\n  Exclusao: {apagadas} apagada(s), {puladas} pulada(s) (vinculo/erro).")
    print("  Concluido.")


if __name__ == "__main__":
    main()
