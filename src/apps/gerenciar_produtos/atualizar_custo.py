"""
Atualiza o PRECO DE CUSTO de produtos no Tiny a partir de uma planilha Excel.

A planilha precisa ter as colunas 'sku' e 'custo'. Por padrao SOMENTE o preco de
custo e alterado; qualquer coluna 'fornecedor' e ignorada.

Atualizar tambem o FORNECEDOR e opcional e fica DESLIGADO por padrao — habilite
com a flag --fornecedor. Nesse modo, o nome da coluna 'fornecedor' e resolvido
para o contato correspondente no Tiny e gravado como fornecedor padrao.
(Atencao: o cadastro de contatos pode ter nomes duplicados; veja a politica de
desempate/aviso em _aplicar_item.)

Uso:
    python atualizar_custo.py <caminho_planilha.xlsx> [tenant] [--fornecedor]

    tenant pode ser: um tenant_id de data/tenants.json | todas (default: todas)

Fluxo (por empresa):
    1. Le a planilha (colunas 'sku', 'custo' e — so com --fornecedor — 'fornecedor').
    2. Para cada SKU: localiza o produto e faz GET completo (/produtos/{id}).
    3. Altera precos.precoCusto (e o fornecedor, se --fornecedor). Nenhum outro
       campo do produto e tocado.
    4. Mostra preview (atual -> novo) e pede confirmacao.
    5. Faz PUT /produtos/{id} com o JSON inteiro.
    6. Salva log em reports/atualizar_custo_<tenant>_<timestamp>.csv.

Notas:
    * precos.precoCustoMedio e calculado pelo Tiny e NAO e tocado.
    * O historico de custo (GET /produtos/{id}/custos) e somente leitura na API:
      cada alteracao adiciona uma linha nova; nao da para apagar registros antigos.
    * O PUT do Tiny exige fornecedores[].padrao nao-nulo; por isso fornecedores
      reais (que o GET retorna sem 'padrao') tem o campo reinjetado, preservando
      o vinculo mesmo no modo so-custo.
"""

import copy
import csv
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import openpyxl

from core.auth_tiny import validate_tenants
from core.tiny_client import TinyClient
from core import desempacotar_data as _desempacotar

REPORTS_DIR = _SRC.parent / "reports"
from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS_PADRAO = get_all_tenant_ids()  # chaves de data/tenants.json

# Valores que aparecem na coluna SKU mas sao rotulos de status, nao codigos reais.
# Linhas com esses "SKUs" sao ignoradas na leitura da planilha.
_SKUS_IGNORAR = {"DESATIVADO", "NOVO"}

# Diferenca de custo abaixo da qual consideramos "sem mudanca" (apenas ruido de
# float). Pequena de proposito: o custo da planilha ja vem arredondado a 2 casas,
# entao qualquer valor armazenado diferente do alvo (ex.: 18.3175 vs 18.32) deve
# ser regravado.
_TOLERANCIA = 0.0001


# ------------------------------------------------------------------ #
#  Normalizacao / parsing                                             #
# ------------------------------------------------------------------ #

def _norm_nome(s) -> str:
    """Uppercase, sem acentos, espacos colapsados — para comparar nomes."""
    if s is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(s))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.upper().split())


def _parse_custo(valor) -> float | None:
    """Converte o valor da coluna 'custo' em float.

    Aceita numeros (7.02) e strings 'R$ 7,02', '7,02', '7.02', '1.234,56'
    (padrao brasileiro) e '1,234.56'.
    """
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip()
    if not s:
        return None
    s = s.replace("R$", "").replace("\xa0", "").replace(" ", "").strip()
    if not s:
        return None

    tem_virgula = "," in s
    tem_ponto = "." in s
    if tem_virgula and tem_ponto:
        if s.rfind(",") > s.rfind("."):  # virgula decimal -> ponto e milhar
            s = s.replace(".", "").replace(",", ".")
        else:                            # ponto decimal -> virgula e milhar
            s = s.replace(",", "")
    elif tem_virgula:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


# ------------------------------------------------------------------ #
#  Leitura da planilha                                                #
# ------------------------------------------------------------------ #

def ler_planilha(caminho: Path, ler_fornecedor: bool = False) -> list[dict]:
    """Le a planilha e retorna lista de {linha, sku, custo_novo, fornecedor_nome}.

    Colunas aceitas (case-insensitive):
      sku        : 'sku' / 'codigo'
      custo      : 'custo' / 'preco_custo' / 'precocusto' / 'novo_custo' / ...
      fornecedor : 'fornecedor' / 'fornecedores' / 'supplier'  (so se ler_fornecedor)

    Com ler_fornecedor=False (padrao) a coluna fornecedor e ignorada e
    'fornecedor_nome' fica vazio em todas as linhas (modo so-custo).
    """
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip().lower() if c is not None else "" for c in next(rows)]

    col_sku = next((i for i, h in enumerate(header) if h in ("sku", "codigo", "código")), None)
    col_custo = next(
        (i for i, h in enumerate(header)
         if h in ("custo", "preco_custo", "preço_custo", "precocusto", "preço de custo",
                  "preco de custo", "preco_de_custo", "custo_novo", "novo_custo")),
        None,
    )
    col_forn = None
    if ler_fornecedor:
        col_forn = next(
            (i for i, h in enumerate(header)
             if h in ("fornecedor", "fornecedores", "supplier", "nome_fornecedor",
                      "nome do fornecedor")),
            None,
        )

    if col_sku is None or col_custo is None:
        raise ValueError(
            f"Planilha precisa ter colunas 'sku' e 'custo'. Cabecalho encontrado: {header}"
        )
    if ler_fornecedor and col_forn is None:
        print("  [aviso] --fornecedor pedido, mas nenhuma coluna 'fornecedor' foi encontrada.")

    itens: list[dict] = []
    for i, row in enumerate(rows, start=2):
        if not row or all(c is None for c in row):
            continue
        sku = row[col_sku] if col_sku < len(row) else None
        custo_raw = row[col_custo] if col_custo < len(row) else None
        forn_raw = row[col_forn] if (col_forn is not None and col_forn < len(row)) else None
        if sku is None or (isinstance(sku, str) and not sku.strip()):
            continue
        if str(sku).strip().upper() in _SKUS_IGNORAR:
            print(f"  [linha {i}] ignorada (rotulo, nao e SKU: {str(sku).strip()!r})")
            continue
        custo = _parse_custo(custo_raw)
        if custo is None:
            print(f"  [linha {i}] ignorada (custo invalido: {custo_raw!r})")
            continue
        fornecedor_nome = str(forn_raw).strip() if forn_raw not in (None, "") else ""
        itens.append({
            "linha": i,
            "sku": str(sku).strip(),
            "custo_novo": round(custo, 2),  # custo monetario: 2 casas decimais
            "fornecedor_nome": fornecedor_nome,
        })

    return itens


# ------------------------------------------------------------------ #
#  Resolucao de fornecedor (nome -> contato no Tiny)                  #
# ------------------------------------------------------------------ #

class ResolvedorFornecedor:
    """Resolve nome de fornecedor -> {id, nome} no cadastro de contatos do Tiny.

    Faz cache por nome normalizado. Retorna (status, dado):
      ('ok', {id, nome})       -> resolvido sem ambiguidade
      ('nao_encontrado', None) -> nenhum contato bate
      ('ambiguo', [contatos])  -> mais de um candidato igualmente plausivel
      ('erro', mensagem)       -> falha na chamada
    """

    def __init__(self, client: TinyClient):
        self.client = client
        self._cache: dict[str, tuple] = {}

    def resolver(self, nome: str) -> tuple:
        chave = _norm_nome(nome)
        if not chave:
            return ("erro", "nome vazio")
        if chave in self._cache:
            return self._cache[chave]
        resultado = self._buscar(nome, chave)
        self._cache[chave] = resultado
        return resultado

    def _buscar(self, nome: str, chave: str) -> tuple:
        itens = []
        try:
            resp = _desempacotar(self.client.buscar_contatos(nome=nome, limit=50))
            itens = resp.get("itens", []) if isinstance(resp, dict) else (resp or [])
        except Exception as e:
            return ("erro", str(e))

        if not itens:
            return ("nao_encontrado", None)

        # match exato por nome ou nome fantasia (normalizados)
        exatos = [
            it for it in itens
            if _norm_nome(it.get("nome")) == chave or _norm_nome(it.get("fantasia")) == chave
        ]
        if len(exatos) == 1:
            return ("ok", {"id": exatos[0].get("id"), "nome": exatos[0].get("nome")})
        if len(exatos) > 1:
            return ("ambiguo", exatos)

        # sem exato: aceita apenas se a busca devolveu um unico candidato
        if len(itens) == 1:
            return ("ok", {"id": itens[0].get("id"), "nome": itens[0].get("nome")})
        return ("ambiguo", itens)


# ------------------------------------------------------------------ #
#  Helpers de payload                                                 #
# ------------------------------------------------------------------ #

def _custo_atual(precos: dict) -> float | None:
    if not isinstance(precos, dict):
        return None
    v = precos.get("precoCusto")
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _fornecedor_atual(d: dict) -> dict | None:
    """Retorna o fornecedor padrao (ou o primeiro) do objeto, ou None."""
    forn = d.get("fornecedores")
    if not isinstance(forn, list) or not forn:
        return None
    validos = [f for f in forn if isinstance(f, dict) and (f.get("id") or 0) > 0]
    if not validos:
        return None
    return next((f for f in validos if f.get("padrao") is True), validos[0])


def _corrigir_fornecedores(d: dict) -> None:
    """Deixa o array 'fornecedores' de `d` aceitavel para o PUT, preservando reais.

    O GET retorna fornecedores reais SEM 'padrao', mas o PUT exige 'padrao'
    nao-nulo (HTTP 400). Tambem remove o placeholder [{id:0, padrao:null}].
    Garante exatamente UM fornecedor com padrao=True; os demais padrao=False.
    """
    forn = d.get("fornecedores")
    if not isinstance(forn, list):
        return
    validos = [f for f in forn if isinstance(f, dict) and (f.get("id") or 0) > 0]
    if not validos:
        d.pop("fornecedores", None)
        return
    idx_padrao = next((i for i, f in enumerate(validos) if f.get("padrao") is True), 0)
    for i, f in enumerate(validos):
        f["padrao"] = (i == idx_padrao)
    d["fornecedores"] = validos


def _preparar_payload(produto_full: dict) -> dict:
    """Copia o JSON do GET e o torna PUT-safe SEM alterar dados de negocio.

    A unica transformacao alem de copiar e corrigir 'fornecedores' (no produto e
    em cada variacao). Preco de custo e fornecedor sao ajustados pelo chamador.
    """
    payload = copy.deepcopy(produto_full)
    _corrigir_fornecedores(payload)
    for var in payload.get("variacoes", []) or []:
        if isinstance(var, dict):
            _corrigir_fornecedores(var)
    return payload


def _definir_fornecedor(target: dict, contato: dict) -> None:
    """Grava `contato` ({id, nome}) como UNICO fornecedor padrao de `target`.

    Preserva 'codigoProdutoNoFornecedor' se o fornecedor ja era o mesmo id.
    """
    cid = contato["id"]
    anterior = _fornecedor_atual(target)
    codigo = anterior.get("codigoProdutoNoFornecedor") if (anterior and anterior.get("id") == cid) else None
    entry = {"id": cid, "nome": contato.get("nome"), "padrao": True}
    if codigo:
        entry["codigoProdutoNoFornecedor"] = codigo
    target["fornecedores"] = [entry]


# ------------------------------------------------------------------ #
#  Montagem do plano                                                  #
# ------------------------------------------------------------------ #

def _aplicar_item(target: dict, item: dict, resolvedor: ResolvedorFornecedor,
                  source: dict | None = None) -> dict:
    """Aplica custo e (opcional) fornecedor no dict `target` (produto ou variacao).

    `source` e de onde se LE o estado atual (custo/fornecedor); por padrao e o
    proprio `target`. Para variacoes isso difere: o array 'variacoes' do PAI traz
    precoCusto sempre 0 (stale), entao o custo atual real vem do GET direto da
    variacao (source=payload_da_variacao), mas a escrita vai no `target` (entrada
    da variacao dentro do pai, que sera enviada no PUT do pai).

    Politica de fornecedor (decidida com o usuario):
      * o CUSTO sempre e aplicado quando muda — problema de fornecedor nunca o bloqueia;
      * se o nome casar com varios contatos (duplicatas), desempata preferindo o
        fornecedor que o produto JA usa (quando o nome confere) — nesse caso e
        tratado como "sem mudanca de fornecedor";
      * se nao der para resolver de forma unica (e o atual nao serve), o fornecedor
        e deixado como esta e um aviso e registrado em 'forn_aviso'.

    Retorna {status, custo_atual, custo_mudou, forn_atual, forn_novo, forn_mudou,
    forn_aviso}. status='pendente' se custo e/ou fornecedor mudam; 'sem_mudanca'
    caso contrario (ainda pode carregar forn_aviso).
    """
    if source is None:
        source = target
    novo_custo = item["custo_novo"]
    forn_nome = item.get("fornecedor_nome") or ""

    custo_atual = _custo_atual(source.get("precos"))
    forn_obj_atual = _fornecedor_atual(source)
    forn_atual_nome = forn_obj_atual.get("nome") if forn_obj_atual else ""
    forn_atual_id = forn_obj_atual.get("id") if forn_obj_atual else None

    contato = None
    forn_novo_nome = forn_atual_nome
    forn_aviso = ""

    if forn_nome:
        status, dado = resolvedor.resolver(forn_nome)
        if status == "ok":
            contato = dado
            forn_novo_nome = dado.get("nome")
        elif status == "ambiguo":
            candidatos = dado or []
            cand_ids = {c.get("id") for c in candidatos}
            # desempate: o produto ja usa um dos contatos de mesmo nome -> mantem
            if (forn_atual_id in cand_ids
                    and _norm_nome(forn_atual_nome) == _norm_nome(forn_nome)):
                forn_novo_nome = forn_atual_nome  # mantem o atual; sem mudanca
            else:
                forn_aviso = f"fornecedor_ambiguo ({len(candidatos)} contatos): {forn_nome}"
        elif status == "nao_encontrado":
            forn_aviso = f"fornecedor_nao_encontrado: {forn_nome}"
        else:  # erro
            forn_aviso = f"fornecedor_erro: {dado}"

    custo_mudou = custo_atual is None or abs(custo_atual - novo_custo) >= _TOLERANCIA
    forn_mudou = bool(contato) and (contato["id"] != forn_atual_id)

    base = {"custo_atual": custo_atual, "custo_mudou": custo_mudou,
            "forn_atual": forn_atual_nome, "forn_novo": forn_novo_nome,
            "forn_mudou": forn_mudou, "forn_aviso": forn_aviso}

    if not custo_mudou and not forn_mudou:
        return {"status": "sem_mudanca", **{**base, "custo_mudou": False}}

    # aplica as mudancas no target
    target.setdefault("precos", {})["precoCusto"] = novo_custo
    if forn_mudou:
        _definir_fornecedor(target, contato)
        forn_novo_nome = contato.get("nome")
        base["forn_novo"] = forn_novo_nome

    return {"status": "pendente", **base}


def montar_preview(client: TinyClient, itens: list[dict]) -> list[dict]:
    """Para cada SKU localiza o produto, faz GET completo e prepara o payload.

    Produtos normais: altera custo/fornecedor e faz 1 PUT em /produtos/{id}.
    Variacoes (tipoVariacao == 'V'): usa o payload da propria variacao e faz 1 PUT
    em /produtos/{pai}/variacoes/{id} (unico caminho que persiste o custo).
    """
    plano: list[dict] = []
    resolvedor = ResolvedorFornecedor(client)
    print(f"\n  Buscando {len(itens)} produto(s) no Tiny...")

    for idx, item in enumerate(itens, start=1):
        sku = item["sku"]
        print(f"  [{idx}/{len(itens)}] SKU={sku} ... ", end="", flush=True)

        try:
            resumo = client.buscar_produto_por_sku(sku)
        except Exception as e:
            print(f"ERRO busca ({e})")
            plano.append({**item, "id": None, "payload": None, "tipo": "?",
                          "alvo_id": None, "status": f"erro_busca: {e}"})
            continue

        if not resumo:
            print("NAO ENCONTRADO")
            plano.append({**item, "id": None, "payload": None, "tipo": "?",
                          "alvo_id": None, "status": "nao_encontrado"})
            continue

        produto_id = resumo.get("id")
        try:
            produto_full = _desempacotar(client.obter_produto(produto_id))
        except Exception as e:
            print(f"ERRO GET ({e})")
            plano.append({**item, "id": produto_id, "payload": None, "tipo": "?",
                          "alvo_id": None, "status": f"erro_get: {e}"})
            continue

        tipo_var = (produto_full.get("tipoVariacao") or "").upper()

        # ---------------- variacao (filha) ---------------------------------- #
        # O custo de uma variacao SO persiste via PUT /produtos/{pai}/variacoes/{id}
        # (PUT no proprio id da variacao da 404; e o array 'variacoes' do pai ignora
        # o custo). Usamos o payload da PROPRIA variacao (produto_full), que ja traz
        # o precoCusto real, e enviamos no endpoint de variacao em aplicar_mudancas.
        if tipo_var == "V":
            parent_id = (produto_full.get("produtoPai") or {}).get("id")
            if not parent_id:
                print("ERRO (variacao sem produtoPai)")
                plano.append({**item, "id": produto_id, "payload": None,
                              "tipo": "variacao", "alvo_id": None,
                              "status": "erro: variacao sem produtoPai"})
                continue

            payload = _preparar_payload(produto_full)
            res = _aplicar_item(payload, item, resolvedor)
            print(_msg_item(res, f"[variacao, pai={parent_id}]"))
            plano.append({**item, "id": produto_id,
                          "payload": payload if res["status"] == "pendente" else None,
                          "tipo": "variacao", "alvo_id": parent_id, **res})
            continue

        # ---------------- produto comum ------------------------------------- #
        payload = _preparar_payload(produto_full)
        res = _aplicar_item(payload, item, resolvedor)
        print(_msg_item(res, ""))
        plano.append({**item, "id": produto_id,
                      "payload": payload if res["status"] == "pendente" else None,
                      "tipo": "normal", "alvo_id": produto_id, **res})

    return plano


def _msg_item(res: dict, extra: str) -> str:
    st = res["status"]
    aviso = f" [!{res['forn_aviso']}]" if res.get("forn_aviso") else ""
    if st == "pendente":
        partes = []
        if res.get("custo_mudou"):
            partes.append("custo")
        if res.get("forn_mudou"):
            partes.append("fornecedor")
        return f"OK ({'+'.join(partes)}) {extra}{aviso}".strip()
    if st == "sem_mudanca":
        return f"OK (ja esta correto) {extra}{aviso}".strip()
    return st  # mensagem de erro


# ------------------------------------------------------------------ #
#  Preview / aplicacao / log                                          #
# ------------------------------------------------------------------ #

def _fmt(v) -> str:
    return f"R$ {v:.2f}" if isinstance(v, (int, float)) else "?"


def imprimir_plano(plano: list[dict]) -> tuple[int, int, int]:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    sem_mudanca = [p for p in plano if p["status"] == "sem_mudanca"]
    erros = [p for p in plano if p["status"] not in ("pendente", "sem_mudanca")]

    print("\n" + "=" * 80)
    print("  PLANO DE MUDANCAS — CUSTO / FORNECEDOR")
    print("=" * 80)

    if pendentes:
        print(f"\n  {len(pendentes)} produto(s) sera(o) atualizado(s):\n")
        for p in pendentes:
            extra = f"  [variacao do pai {p.get('alvo_id')}]" if p.get("tipo") == "variacao" else ""
            print(f"   SKU {p['sku']:>15}  (id={p['id']}){extra}")
            if p.get("custo_mudou"):
                print(f"     custo      : {_fmt(p['custo_atual'])}  ->  {_fmt(p['custo_novo'])}")
            if p.get("forn_mudou"):
                print(f"     fornecedor : {p.get('forn_atual') or '(nenhum)'}  ->  {p.get('forn_novo')}")
            if p.get("forn_aviso"):
                print(f"     [aviso] {p['forn_aviso']} (fornecedor mantido)")
            print()

    if sem_mudanca:
        print(f"  {len(sem_mudanca)} ja estao corretos (sera(o) ignorado(s))")

    # avisos de fornecedor que nao impedem a atualizacao do custo
    avisos = [p for p in plano if p.get("forn_aviso") and p["status"] in ("pendente", "sem_mudanca")]
    if avisos:
        print(f"\n  {len(avisos)} aviso(s) de fornecedor (custo segue normalmente):")
        for p in avisos:
            print(f"   SKU {p['sku']}: {p['forn_aviso']}")

    if erros:
        print(f"\n  {len(erros)} com problemas:")
        for p in erros:
            print(f"   SKU {p['sku']}: {p['status']}")

    print("=" * 80)
    return len(pendentes), len(sem_mudanca), len(erros)


def aplicar_mudancas(client: TinyClient, plano: list[dict]) -> list[dict]:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    variacoes = [p for p in pendentes if p.get("tipo") == "variacao"]
    normais = [p for p in pendentes if p.get("tipo") != "variacao"]
    total_puts = len(pendentes)  # 1 PUT por item (normal ou variacao)
    print(f"\n  Aplicando {len(pendentes)} mudanca(s) via PUT "
          f"({len(normais)} produto(s) + {len(variacoes)} variacao(oes))...")

    contador = 0
    for p in pendentes:
        contador += 1
        print(f"  [{contador}/{total_puts}] SKU={p['sku']} ... ", end="", flush=True)
        try:
            if p.get("tipo") == "variacao":
                # PUT /produtos/{pai}/variacoes/{id} — unico caminho que grava o custo
                client.atualizar_variacao(p["alvo_id"], p["id"], p["payload"])
            else:
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
    arquivo = REPORTS_DIR / f"atualizar_custo_{tenant_id}_{stamp}.csv"

    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["linha", "sku", "id", "tipo", "alvo_id",
                        "custo_atual", "custo_novo",
                        "fornecedor_atual", "fornecedor_novo",
                        "fornecedor_aviso", "status"],
        )
        writer.writeheader()
        for p in plano:
            writer.writerow({
                "linha": p.get("linha", ""),
                "sku": p.get("sku", ""),
                "id": p.get("id", ""),
                "tipo": p.get("tipo", ""),
                "alvo_id": p.get("alvo_id", ""),
                "custo_atual": p.get("custo_atual", ""),
                "custo_novo": p.get("custo_novo", ""),
                "fornecedor_atual": p.get("forn_atual", ""),
                "fornecedor_novo": p.get("forn_novo", ""),
                "fornecedor_aviso": p.get("forn_aviso", ""),
                "status": p.get("status", ""),
            })

    return arquivo


def processar_tenant(tenant_id: str, itens: list[dict]) -> None:
    print("\n" + "#" * 80)
    print(f"#  EMPRESA: {tenant_id.upper()}")
    print("#" * 80)

    client = TinyClient(tenant_id)
    plano = montar_preview(client, itens)
    pendentes, _, _ = imprimir_plano(plano)

    if pendentes == 0:
        print(f"\n  [{tenant_id}] Nenhuma mudanca a aplicar.")
        log = salvar_log(plano, tenant_id)
        print(f"  Log salvo em: {log}")
        return

    resp = input(f"\n  [{tenant_id}] Confirma aplicar {pendentes} mudanca(s)? (s/N): ").strip().lower()
    if resp != "s":
        print(f"  [{tenant_id}] Cancelado. Nenhuma mudanca aplicada.")
        log = salvar_log(plano, tenant_id)
        print(f"  Log do dry-run salvo em: {log}")
        return

    aplicar_mudancas(client, plano)
    log = salvar_log(plano, tenant_id)

    sucesso = sum(1 for p in plano if p["status"] == "atualizado")
    print(f"\n  [{tenant_id}] Concluido: {sucesso}/{pendentes} atualizado(s).")
    print(f"  Log salvo em: {log}")


def main():
    # separa flags dos argumentos posicionais
    flags = {a.lower() for a in sys.argv[1:] if a.startswith("-")}
    posicionais = [a for a in sys.argv[1:] if not a.startswith("-")]
    atualizar_forn = "--fornecedor" in flags or "-f" in flags

    if not posicionais:
        print("Uso: python atualizar_custo.py <caminho_planilha.xlsx> [tenant] [--fornecedor]")
        print("     tenant: tenant_id de data/tenants.json | todas (default: todas)")
        print("     --fornecedor: tambem atualiza o fornecedor (padrao: so o custo)")
        sys.exit(1)

    caminho = Path(posicionais[0])
    if not caminho.exists():
        print(f"Arquivo nao encontrado: {caminho}")
        sys.exit(1)

    arg_tenant = (posicionais[1] if len(posicionais) > 1 else "todas").strip().lower()
    if arg_tenant in ("todas", "ambas", "all", ""):
        tenants = list(TENANTS_PADRAO)
    else:
        tenants = [t.strip() for t in arg_tenant.split(",") if t.strip()]

    if not validate_tenants(tenants):
        print("Autenticacao nao concluida. Encerrando.")
        sys.exit(1)

    modo = "CUSTO + FORNECEDOR" if atualizar_forn else "SOMENTE CUSTO"
    print(f"\n  Modo: {modo}")
    print(f"  Lendo planilha: {caminho.name}")
    itens = ler_planilha(caminho, ler_fornecedor=atualizar_forn)
    print(f"  {len(itens)} linha(s) com SKU + custo validos.")
    if not itens:
        print("Nada a fazer.")
        return

    print(f"  Empresas a processar: {', '.join(tenants)}")
    for tenant_id in tenants:
        processar_tenant(tenant_id, itens)

    print("\n  Tudo concluido.")


if __name__ == "__main__":
    main()
