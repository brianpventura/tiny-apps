"""
Aplica BALANCO DE ESTOQUE (tipo 'B') no Tiny a partir de uma planilha Excel.

A planilha pode vir em dois layouts:

  (a) uma quantidade por linha + coluna de empresa:
      SKU | Produto | Quant | Lote | Fabricacao | Validade | Fabricante | Empresa

  (b) uma coluna de quantidade POR EMPRESA (planilha unica de todas):
      SKU | Produto | ... | Lote | Fabricacao | Validade |
      Resultado <Empresa A> | Resultado <Empresa B>
      (<Empresa> = tenant_id ou `nome` de data/tenants.json)
      Cada linha gera um lancamento por empresa. Celula vazia = empresa fora
      daquela linha; um 0 explicito zera o estoque.

Para cada SKU: localiza o produto, le o saldo atual, e envia
POST /estoque/{idProduto} com tipo='B' e quantidade = coluna 'Quant'.
Lote e validade vao no campo 'observacoes', no formato:

    Balanço - Lote: 1234 | Valid: 27/03/2029 ; Lote: 5678 | Valid: 06/10/2029 | Saldo ant: 300

'Saldo ant' e o saldo lido no Tiny imediatamente antes do lancamento — o unico
registro que sobra dele, ja que o balanco sobrescreve o saldo.

O campo so guarda 100 CARACTERES: a API aceita textos maiores sem erro e corta
na gravacao, silenciosamente. O texto e montado priorizando o saldo anterior;
lotes que nao couberem viram aviso no preview e no log.

So entram na observacao os lotes COM saldo naquela empresa: se o SKU tem dois
lotes e um esta zerado para a empresa, so o outro e citado.

Uso:
    python balanco_estoque.py <planilha.xlsx> [tenant] [flags]

    tenant : tenant_id de data/tenants.json | auto   (default: auto = usa a coluna 'Empresa')

    Como cada empresa tem sua propria planilha, o normal e passar o tenant no
    comando. No modo 'auto' TODAS as linhas precisam da coluna 'Empresa'
    preenchida — o script aborta se faltar, em vez de adivinhar (o mesmo SKU
    pode existir em mais de uma empresa com ids diferentes).

    --dry-run        so mostra o plano, nunca escreve (nao pergunta nada)
    --limite N       processa apenas as N primeiras linhas (use para testar!)
    --deposito ID    lanca num deposito especifico (default: o padrao do Tiny)
    --preco-zero     envia precoUnitario=0 em vez do custo atual do produto
    --confirmar      pula o prompt de confirmacao (para execucao nao-interativa)
    --forcar         lanca mesmo onde o saldo ja esta no alvo (regravar obs)
    --regravar       quando o Tiny responde 204 (saldo ja e o alvo, nada
                     gravado), lanca alvo+1 e volta ao alvo — 2 movimentacoes,
                     saldo final identico, e a observacao fica gravada
    --somente-lancados
                     processa so os SKUs que ja tiveram lancamento no arquivo
                     de --saldo-anterior (coluna 'id_lancamento' preenchida)
    --divergencia=PCT
                     anexa 'Divergencia -95%' a observacao quando o saldo novo
                     difere do atual em PCT% ou mais (base: saldo lido agora,
                     por empresa; aviso com alvo = saldo nunca e marcado)
    --saldo-anterior ARQ.csv
                     le o saldo anterior REAL de um balanco ja aplicado e usa
                     esse valor em 'Saldo ant'. Sem isso, uma regravacao
                     gravaria o saldo pos-balanco, que nao e o saldo anterior.

Fluxo:
    1. Le a planilha e agrupa por (empresa, SKU).
    2. Para cada SKU: busca o produto, GET /produtos/{id} e GET /estoque/{id}.
    3. Monta o preview (saldo atual -> saldo alvo) e pede confirmacao.
    4. POST /estoque/{id} com tipo B.
    5. Le o estoque de novo (depois) e salva log CSV + snapshot JSON em reports/.

AVISOS IMPORTANTES
    * O lancamento e IRREVERSIVEL: a API nao tem DELETE de movimentacao.
      Um balanco errado so se corrige com outro lancamento, e ambos ficam
      no historico.
    * 'B' redefine o saldo (nao soma). Se o SKU aparece em varias linhas
      (lotes diferentes), as quantidades sao SOMADAS num unico lancamento —
      caso contrario o ultimo balanco apagaria os anteriores.
    * A API v3 NAO tem campos de lote/validade. O texto em 'observacoes' e
      apenas descritivo: nao gera saldo por lote nem alimenta rastreabilidade.
      Nao ha garantia de que o balanco preserve lotes existentes no ERP.
    * precoUnitario e obrigatorio na API e influencia o custo medio. Por
      padrao reenviamos o precoCusto atual do produto para manter o custo
      neutro. Com --preco-zero envia 0 (pode zerar o custo medio).
"""

import csv
import json
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import openpyxl

from core.auth_tiny import TENANT_NAMES, validate_tenants
from core.tiny_client import TinyClient
from core import desempacotar_data as _desempacotar

REPORTS_DIR = _SRC.parent / "reports"

# O Tiny grava so os 100 primeiros caracteres de 'observacoes' e nao avisa:
# a API aceita textos maiores sem erro (testado ate 1000 chars) e corta na
# gravacao. Manter o texto dentro do limite e responsabilidade do cliente.
MAX_OBSERVACOES = 100

# Colunas aceitas (normalizadas: minusculas, sem acento)
_COLS = {
    "sku": ("sku", "codigo"),
    "produto": ("produto", "descricao", "nome"),
    "quant": ("quant", "quantidade", "qtd", "qtde", "saldo"),
    "lote": ("lote", "n lote", "numero do lote", "num lote"),
    "fabricacao": ("fabricacao", "data de fabricacao", "dt fabricacao"),
    "validade": ("validade", "data de validade", "dt validade", "vencimento"),
    "fabricante": ("fabricante", "marca"),
    "empresa": ("empresa", "tenant"),
}


# ------------------------------------------------------------------ #
#  Normalizacao / parsing                                             #
# ------------------------------------------------------------------ #

def _norm(s) -> str:
    """Minusculas, sem acentos, espacos colapsados."""
    if s is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(s))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.lower().split())


def _parse_quant(valor) -> float | None:
    """Converte a coluna 'Quant' em float. Aceita 10, '10', '10,5', '1.234,5'."""
    if valor is None:
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip().replace("\xa0", "").replace(" ", "")
    if not s:
        return None

    tem_virgula, tem_ponto = "," in s, "." in s
    if tem_virgula and tem_ponto:
        if s.rfind(",") > s.rfind("."):      # virgula decimal
            s = s.replace(".", "").replace(",", ".")
        else:                                 # ponto decimal
            s = s.replace(",", "")
    elif tem_virgula:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


def _parse_data(valor) -> str:
    """Converte uma data da planilha para 'dd/mm/aaaa'. Retorna '' se nao der.

    Aceita datetime/date (openpyxl ja converte celulas de data), alem de
    strings 'dd/mm/aaaa', 'aaaa-mm-dd', 'dd-mm-aaaa' e 'mm/aaaa'.
    """
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")

    s = str(valor).strip()
    if not s:
        return ""

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass

    # mm/aaaa (validade de lote costuma vir assim) -> ultimo dia do mes
    for fmt in ("%m/%Y", "%m-%Y"):
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("01/%m/%Y")
        except ValueError:
            pass

    return s  # devolve como veio; o preview mostra e o usuario decide


def _resolver_tenant(valor) -> str | None:
    """Mapeia o texto da coluna 'Empresa' para um tenant_id conhecido.

    Retorna None se a celula estiver vazia ou nao casar com nenhum tenant.
    NAO ha fallback silencioso de proposito: o mesmo SKU pode existir em
    mais de uma empresa com ids diferentes, entao adivinhar a empresa
    lancaria o balanco no lugar errado sem nenhum sinal de erro.
    """
    alvo = _norm(valor)
    if not alvo:
        return None
    for tid, nome in TENANT_NAMES.items():
        if alvo == _norm(tid) or alvo == _norm(nome):
            return tid
    return None


# ------------------------------------------------------------------ #
#  Leitura da planilha                                                #
# ------------------------------------------------------------------ #

def _mapear_colunas(header: list[str]) -> dict:
    """Retorna {campo: indice} para as colunas reconhecidas."""
    norm = [_norm(h) for h in header]
    mapa = {}
    for campo, aliases in _COLS.items():
        idx = next((i for i, h in enumerate(norm) if h in aliases), None)
        if idx is not None:
            mapa[campo] = idx
    return mapa


def _mapear_colunas_por_empresa(header: list[str]) -> dict:
    """Retorna {tenant_id: indice} para colunas 'Resultado <Empresa>'.

    Layout de planilha unica com uma coluna de quantidade por empresa
    (ex.: 'Resultado Empresa A' e 'Resultado Empresa B'). Cada linha vira um
    item por empresa, sem precisar da coluna 'Empresa'.
    """
    out: dict[str, int] = {}
    for i, h in enumerate(header):
        n = _norm(h)
        for prefixo in ("resultado", "saldo final", "balanco"):
            if not n.startswith(prefixo):
                continue
            resto = n[len(prefixo):].strip(" -:")
            for tid, nome in TENANT_NAMES.items():
                if resto in (_norm(tid), _norm(nome)):
                    out[tid] = i
    return out


def ler_planilha(caminho: Path, aba: str | None = None) -> list[dict]:
    """Le a planilha e retorna uma lista de itens (um por SKU x empresa).

    Aceita dois layouts:
      * coluna unica 'Quant' + coluna 'Empresa';
      * uma coluna de quantidade por empresa ('Resultado <Empresa>', etc.),
        caso em que cada linha gera um item por empresa.
    """
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    if aba:
        if aba not in wb.sheetnames:
            raise ValueError(f"Aba {aba!r} nao existe. Abas: {wb.sheetnames}")
        ws = wb[aba]
    else:
        ws = wb.active
    print(f"  Aba      : {ws.title}")

    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]
    mapa = _mapear_colunas(header)
    cols_empresa = _mapear_colunas_por_empresa(header)

    if "sku" not in mapa:
        raise ValueError(f"Planilha precisa da coluna 'SKU'. Cabecalho lido: {header}")
    if not cols_empresa and "quant" not in mapa:
        raise ValueError(
            "Planilha precisa da coluna 'Quant' ou de colunas 'Resultado <Empresa>'. "
            f"Cabecalho lido: {header}"
        )
    if cols_empresa:
        rotulos = ", ".join(f"{t}='{header[i]}'" for t, i in sorted(cols_empresa.items()))
        print(f"  Colunas  : {rotulos}")
    for opcional in ("lote", "validade"):
        if opcional not in mapa:
            print(f"  [aviso] coluna '{opcional}' nao encontrada — ficara vazia na observacao.")

    def celula(row, campo):
        i = mapa.get(campo)
        return row[i] if (i is not None and i < len(row)) else None

    itens: list[dict] = []
    for n, row in enumerate(rows, start=2):
        if not row or all(c is None for c in row):
            continue

        sku = celula(row, "sku")
        if sku is None or not str(sku).strip():
            continue

        comum = {
            "linha": n,
            "sku": str(sku).strip(),
            "produto": str(celula(row, "produto") or "").strip(),
            "lote": str(celula(row, "lote") or "").strip(),
            "fabricacao": _parse_data(celula(row, "fabricacao")),
            "validade": _parse_data(celula(row, "validade")),
            "fabricante": str(celula(row, "fabricante") or "").strip(),
        }

        if cols_empresa:
            # uma coluna de quantidade por empresa
            for tid, i in cols_empresa.items():
                bruto = row[i] if i < len(row) else None
                # celula vazia != zero: vazio significa "esta empresa nao entra
                # nesta linha"; zerar estoque so acontece com um 0 explicito.
                if bruto is None or str(bruto).strip() == "":
                    continue
                quant = _parse_quant(bruto)
                if quant is None:
                    print(f"  [linha {n}/{tid}] ignorada (valor invalido: {bruto!r})")
                    continue
                if quant < 0:
                    print(f"  [linha {n}/{tid}] ignorada (valor negativo: {quant})")
                    continue
                itens.append({**comum, "quant": quant, "tenant": tid,
                              "empresa_planilha": header[i], "por_coluna": True})
        else:
            quant = _parse_quant(celula(row, "quant"))
            if quant is None:
                print(f"  [linha {n}] ignorada (Quant invalida: {celula(row, 'quant')!r})")
                continue
            if quant < 0:
                print(f"  [linha {n}] ignorada (Quant negativa: {quant})")
                continue
            itens.append({**comum, "quant": quant, "por_coluna": False,
                          "empresa_planilha": str(celula(row, "empresa") or "").strip(),
                          "tenant": _resolver_tenant(celula(row, "empresa"))})

    return itens


# ------------------------------------------------------------------ #
#  Agrupamento por SKU                                                #
# ------------------------------------------------------------------ #

def agrupar_por_sku(itens: list[dict]) -> list[dict]:
    """Soma as quantidades das linhas que repetem o mesmo (tenant, SKU).

    Necessario porque tipo 'B' REDEFINE o saldo: dois balancos seguidos para
    o mesmo produto fariam o segundo apagar o primeiro. Como um SKU pode
    aparecer varias vezes (um lote por linha), somamos tudo num unico
    lancamento e juntamos os lotes na observacao.
    """
    grupos: dict[tuple, dict] = {}
    for it in itens:
        chave = (it["tenant"], it["sku"])
        g = grupos.get(chave)
        if g is None:
            grupos[chave] = {
                "tenant": it["tenant"],
                "sku": it["sku"],
                "produto": it["produto"],
                "quant": it["quant"],
                "linhas": [it["linha"]],
                "lotes": [],
                "lotes_zerados": [],
            }
            g = grupos[chave]
        else:
            g["quant"] += it["quant"]
            g["linhas"].append(it["linha"])

        # So entra na observacao o lote que TEM saldo nesta empresa. Um SKU
        # pode ter dois lotes com o lote A zerado para uma empresa e cheio para
        # a outra; citar o lote zerado na observacao da
        # empresa que nao o tem seria registro errado — e ainda gastaria parte
        # dos 100 caracteres disponiveis.
        if (it["lote"] or it["validade"]) and it["quant"] > 0:
            g["lotes"].append({"lote": it["lote"], "validade": it["validade"]})
        elif it["lote"] and it["quant"] <= 0:
            g["lotes_zerados"].append(it["lote"])

    return list(grupos.values())


def montar_observacao(grupo: dict, saldo_antes: float | None = None,
                      nota: str = "") -> tuple[str, int]:
    """Monta o texto de 'observacoes'. Retorna (texto, lotes_omitidos).

    Formato:
      Balanço - Lote: 1234 | Valid: 27/03/2029 ; Lote: 5678 | Valid: 06/10/2029 | Saldo ant: 300

    Sem data no texto de proposito: a data da contagem vai no campo 'data' do
    proprio lancamento. Com varios lotes o separador entre eles e ' ; ' — o
    ' | ' ja separa os campos dentro de cada entrada.

    'Saldo ant' e o saldo lido no Tiny ANTES do lancamento; e o unico registro
    que sobra dele, ja que o balanco sobrescreve o saldo. Fica de fora quando a
    leitura do estoque falhou (saldo_antes=None).

    LIMITE DE 100 CARACTERES: o Tiny aceita a requisicao com texto maior (nao
    da erro) mas GRAVA so os 100 primeiros, cortando silenciosamente o final —
    justamente onde fica o saldo anterior. Por isso o texto e montado de tras
    para frente em prioridade: prefixo, 'Saldo ant' e a nota sempre cabem, e os
    lotes entram enquanto houver espaco. Os que nao couberem sao contados no
    retorno para virar aviso no preview e no log, nunca perda silenciosa.

    nota: texto livre anexado ao fim (ex.: 'Divergencia -95%', de --divergencia).
    """
    prefixo = "Balanço"
    partes = [
        f"Lote: {lt['lote'] or 'N/D'} | Valid: {lt['validade'] or 'N/D'}"
        for lt in grupo["lotes"]
    ]
    sufixo = " | ".join(s for s in (
        f"Saldo ant: {saldo_antes:g}" if saldo_antes is not None else "",
        nota,
    ) if s)

    def montar(n: int) -> str:
        corpo = " ; ".join(partes[:n])
        if sufixo:
            corpo = f"{corpo} | {sufixo}" if corpo else sufixo
        return f"{prefixo} - {corpo}" if corpo else prefixo

    for n in range(len(partes), -1, -1):
        texto = montar(n)
        if len(texto) <= MAX_OBSERVACOES:
            return texto, len(partes) - n

    # nem so o saldo coube (nao deve acontecer): corta como ultimo recurso
    return montar(0)[:MAX_OBSERVACOES], len(partes)


# ------------------------------------------------------------------ #
#  Saldo anterior historico                                           #
# ------------------------------------------------------------------ #

def carregar_skus_com_lancamento(caminho: Path, tenant: str) -> set:
    """SKUs que tiveram lancamento aplicado num balanco anterior.

    Le a coluna 'id_lancamento' do consolidado/log: linha com id preenchido
    significa que aquele SKU realmente gerou movimentacao.
    """
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        reg = list(csv.DictReader(f))
    cols = reg[0].keys() if reg else []
    if "id_lancamento" not in cols:
        raise ValueError(f"{caminho.name} nao tem a coluna 'id_lancamento'.")
    skus = set()
    for r in reg:
        if "empresa" in cols and _norm(r.get("empresa")) != _norm(tenant):
            continue
        if str(r.get("id_lancamento") or "").strip():
            skus.add(str(r.get("sku") or "").strip())
    return skus


def carregar_saldo_historico(caminho: Path, tenant: str) -> dict:
    """Le {sku: saldo_anterior} de um CSV de balanco anterior.

    Aceita o consolidado (colunas 'empresa','sku','saldo_anterior') e o log
    por tenant (colunas 'sku','saldo_antes'). Serve para regravar observacoes
    mantendo o saldo anterior REAL — depois de um balanco aplicado, o saldo
    vivo ja e o novo e nao serve mais como registro do que havia antes.
    """
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        reg = list(csv.DictReader(f))
    if not reg:
        raise ValueError(f"{caminho.name} esta vazio.")

    cols = reg[0].keys()
    col_saldo = next((c for c in ("saldo_anterior", "saldo_antes") if c in cols), None)
    if col_saldo is None or "sku" not in cols:
        raise ValueError(
            f"{caminho.name} precisa das colunas 'sku' e 'saldo_anterior' "
            f"(ou 'saldo_antes'). Colunas lidas: {list(cols)}"
        )

    hist: dict[str, float] = {}
    for r in reg:
        if "empresa" in cols and _norm(r.get("empresa")) != _norm(tenant):
            continue
        sku = str(r.get("sku") or "").strip()
        bruto = r.get(col_saldo)
        if not sku or bruto is None or str(bruto).strip() == "":
            continue
        valor = _parse_quant(bruto)
        if valor is not None:
            hist[sku] = valor
    return hist


# ------------------------------------------------------------------ #
#  Montagem do plano                                                  #
# ------------------------------------------------------------------ #

def _saldo_total(estoque: dict) -> float | None:
    try:
        return float(estoque.get("saldo"))
    except (TypeError, ValueError):
        return None


def _preco_custo(produto_full: dict) -> float:
    precos = produto_full.get("precos") or {}
    for campo in ("precoCusto", "precoCustoMedio"):
        try:
            v = float(precos.get(campo))
            if v > 0:
                return v
        except (TypeError, ValueError):
            continue
    return 0.0


def _nota_divergencia(saldo_antes: float | None, alvo: float,
                      limite_pct: float | None) -> str:
    """'Divergencia -95%' quando o alvo difere do saldo atual em >= limite_pct.

    Base e o saldo lido agora, empresa a empresa: num aviso (alvo = saldo) a
    variacao e zero e nada e anotado. Saldo atual <= 0 nao tem percentual.
    """
    if limite_pct is None or saldo_antes is None or saldo_antes <= 0:
        return ""
    pct = (alvo - saldo_antes) / saldo_antes * 100
    return f"Divergencia {pct:+.0f}%" if abs(pct) >= limite_pct else ""


def montar_preview(client: TinyClient, grupos: list[dict], data_api: str,
                   preco_zero: bool, forcar: bool = False,
                   saldo_hist: dict | None = None,
                   divergencia: float | None = None) -> list[dict]:
    """Localiza cada SKU, le o saldo atual e prepara o lancamento.

    forcar: lanca mesmo quando o saldo ja esta no alvo (usado para regravar a
    observacao de um balanco anterior).
    divergencia: limite em % para anotar 'Divergencia <pct>' na observacao.
    saldo_hist: {sku: saldo} com o saldo ANTERIOR historico. Quando informado,
    e ele que vai para 'Saldo ant' na observacao, e nao o saldo lido agora —
    numa regravacao o saldo vivo ja e o pos-balanco e nao serve de registro.
    """
    plano: list[dict] = []
    total = len(grupos)
    print(f"\n  Consultando {total} SKU(s) no Tiny...")

    for idx, g in enumerate(grupos, start=1):
        sku = g["sku"]
        print(f"  [{idx}/{total}] SKU={sku} ... ", end="", flush=True)

        obs_ini, _ = montar_observacao(g)
        base = {
            **g,
            "id": None,
            "saldo_antes": None,
            "saldo_depois": None,
            "preco_unitario": None,
            "observacao": obs_ini,
            "data": data_api,
            "id_lancamento": None,
            "aviso": "",
        }

        try:
            resumo = client.buscar_produto_por_sku(sku)
        except Exception as e:
            print(f"ERRO busca ({e})")
            plano.append({**base, "status": f"erro_busca: {e}"})
            continue

        if not resumo:
            print("NAO ENCONTRADO")
            plano.append({**base, "status": "nao_encontrado"})
            continue

        pid = resumo.get("id")
        base["id"] = pid

        try:
            produto_full = _desempacotar(client.obter_produto(pid))
        except Exception as e:
            print(f"ERRO GET produto ({e})")
            plano.append({**base, "status": f"erro_get_produto: {e}"})
            continue

        # nome real do Tiny (a coluna 'Produto' da planilha e so referencia)
        base["produto_tiny"] = produto_full.get("descricao", "")

        try:
            estoque = _desempacotar(client.obter_estoque(pid))
            base["saldo_antes"] = _saldo_total(estoque)
            base["estoque_antes"] = estoque
        except Exception as e:
            print(f"ERRO GET estoque ({e})")
            plano.append({**base, "status": f"erro_get_estoque: {e}"})
            continue

        # 'Saldo ant' na observacao: o historico manda quando informado, senao
        # o saldo lido agora. Numa regravacao o saldo vivo ja e o pos-balanco.
        saldo_obs = base["saldo_antes"]
        usou_hist = False
        if saldo_hist is not None:
            hist = saldo_hist.get(sku)
            if hist is not None:
                saldo_obs = hist
                usou_hist = True
        base["saldo_obs"] = saldo_obs
        nota = _nota_divergencia(base["saldo_antes"], g["quant"], divergencia)
        base["observacao"], omitidos = montar_observacao(g, saldo_obs, nota)

        preco = 0.0 if preco_zero else _preco_custo(produto_full)
        base["preco_unitario"] = preco

        avisos = []
        if not preco_zero and preco == 0.0:
            avisos.append("produto sem precoCusto — enviando precoUnitario=0")
        if len(g["linhas"]) > 1:
            avisos.append(f"{len(g['linhas'])} linhas somadas (linhas {g['linhas']})")
        if not g["lotes"]:
            avisos.append("sem lote/validade na planilha")
        if g.get("lotes_zerados"):
            avisos.append(f"lote(s) zerado(s) nesta empresa fora da obs: "
                          f"{', '.join(g['lotes_zerados'])}")
        if omitidos:
            avisos.append(f"{omitidos} lote(s) omitido(s) por causa do limite de "
                          f"{MAX_OBSERVACOES} chars")
        if saldo_hist is not None and not usou_hist:
            avisos.append("sem saldo historico para este SKU — usando o saldo atual")
        base["aviso"] = "; ".join(avisos)

        if (not forcar and base["saldo_antes"] is not None
                and abs(base["saldo_antes"] - g["quant"]) < 1e-9):
            print(f"OK (saldo ja e {g['quant']:g})")
            plano.append({**base, "status": "sem_mudanca"})
            continue

        print(f"OK ({base['saldo_antes']:g} -> {g['quant']:g})"
              if base["saldo_antes"] is not None else f"OK (-> {g['quant']:g})")
        plano.append({**base, "status": "pendente"})

    return plano


# ------------------------------------------------------------------ #
#  Preview / aplicacao / log                                          #
# ------------------------------------------------------------------ #

def imprimir_plano(plano: list[dict], tenant: str, deposito_id, preco_zero: bool) -> int:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    iguais = [p for p in plano if p["status"] == "sem_mudanca"]
    erros = [p for p in plano if p["status"] not in ("pendente", "sem_mudanca")]

    print("\n" + "=" * 84)
    print(f"  PLANO DE BALANCO DE ESTOQUE (tipo 'B') — {tenant.upper()}")
    print("=" * 84)
    print(f"  Deposito : {deposito_id if deposito_id else 'padrao do Tiny'}")
    print(f"  Preco    : {'ZERO (--preco-zero)' if preco_zero else 'precoCusto atual do produto'}")

    if pendentes:
        print(f"\n  {len(pendentes)} lancamento(s) sera(o) enviado(s):\n")
        for p in pendentes:
            antes = f"{p['saldo_antes']:g}" if p["saldo_antes"] is not None else "?"
            print(f"   SKU {p['sku']:>15}  (id={p['id']})  {p.get('produto_tiny', '')[:45]}")
            print(f"     saldo      : {antes}  ->  {p['quant']:g}")
            print(f"     preco unit.: {p['preco_unitario']:.2f}")
            print(f"     observacoes: {p['observacao']}  ({len(p['observacao'])}/{MAX_OBSERVACOES})")
            if p["aviso"]:
                print(f"     [aviso] {p['aviso']}")
            print()

    if iguais:
        print(f"  {len(iguais)} SKU(s) ja estao com o saldo alvo (serao ignorados)")

    if erros:
        print(f"\n  {len(erros)} com problemas:")
        for p in erros:
            print(f"   SKU {p['sku']}: {p['status']}")

    print("=" * 84)
    if pendentes:
        print("  ATENCAO: o lancamento e IRREVERSIVEL (a API nao tem DELETE de")
        print("  movimentacao) e tipo 'B' REDEFINE o saldo. Nao ha garantia de")
        print("  que lotes/validades ja existentes no ERP sejam preservados.")
        print("=" * 84)
    return len(pendentes)


def _regravar_via_bump(client: TinyClient, p: dict, deposito_id) -> dict:
    """Forca a gravacao da observacao quando o saldo ja e o alvo.

    O Tiny so cria movimentacao (e so grava a observacao) quando o saldo muda.
    Para regravar sem alterar o resultado final, lanca-se o alvo+1 e em seguida
    o alvo de volta — o SEGUNDO lancamento carrega a observacao correta e o
    saldo termina identico ao que estava.

    Efeito colateral: 2 movimentacoes no historico e o saldo fica 1 unidade
    acima por alguns instantes. Retorna a resposta do lancamento final.
    """
    comum = dict(tipo="B", preco_unitario=p["preco_unitario"], data=p["data"],
                 id_deposito=deposito_id)
    client.atualizar_estoque(p["id"], quantidade=p["quant"] + 1,
                             observacoes="Ajuste tecnico - regravacao de observacao",
                             **comum)
    p["bump_aplicado"] = True
    return client.atualizar_estoque(p["id"], quantidade=p["quant"],
                                    observacoes=p["observacao"], **comum)


def aplicar(client: TinyClient, plano: list[dict], deposito_id,
            regravar: bool = False) -> None:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    print(f"\n  Enviando {len(pendentes)} lancamento(s)"
          f"{' [modo regravacao]' if regravar else ''}...")

    for i, p in enumerate(pendentes, start=1):
        print(f"  [{i}/{len(pendentes)}] SKU={p['sku']} ... ", end="", flush=True)
        try:
            resp = client.atualizar_estoque(
                p["id"],
                tipo="B",
                quantidade=p["quant"],
                preco_unitario=p["preco_unitario"],
                data=p["data"],
                observacoes=p["observacao"],
                id_deposito=deposito_id,
            )
            # 204 = saldo ja era o alvo, nada gravado. Com --regravar, forca.
            if regravar and not (isinstance(resp, dict) and resp.get("idLancamento")):
                print("204 -> bump ... ", end="", flush=True)
                resp = _regravar_via_bump(client, p, deposito_id)
            resp = _desempacotar(resp) if isinstance(resp, dict) else {}
            p["id_lancamento"] = resp.get("idLancamento")
            if not p["id_lancamento"]:
                # HTTP 204 sem corpo: quando a quantidade pedida e igual ao
                # saldo atual o Tiny NAO cria movimentacao — ignora a
                # requisicao inteira, observacao inclusive. Nao e sucesso.
                p["status"] = "sem_efeito"
                p["aviso"] = ((p["aviso"] + "; ") if p["aviso"] else "") + \
                    "HTTP 204: saldo ja era o alvo, nenhum lancamento criado " \
                    "(observacao NAO gravada)"
                print("SEM EFEITO (204 — saldo ja era o alvo, obs nao gravada)")
                continue
            p["status"] = "aplicado"
            print(f"OK (idLancamento={p['id_lancamento']})")
        except Exception as e:
            p["status"] = f"erro_post: {e}"
            print(f"ERRO ({e})")
            continue

        # releitura: evidencia do antes/depois (essencial para conferir lotes)
        try:
            depois = _desempacotar(client.obter_estoque(p["id"]))
            p["saldo_depois"] = _saldo_total(depois)
            p["estoque_depois"] = depois
            print(f"      saldo confirmado: {p['saldo_depois']:g}")
        except Exception as e:
            p["aviso"] = (p["aviso"] + "; " if p["aviso"] else "") + f"releitura falhou: {e}"


def salvar_relatorios(plano: list[dict], tenant: str) -> tuple[Path, Path]:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = REPORTS_DIR / f"balanco_estoque_{tenant}_{stamp}.csv"
    json_path = REPORTS_DIR / f"balanco_estoque_{tenant}_{stamp}.json"

    campos = ["linhas", "sku", "id", "produto_tiny", "quant", "saldo_antes",
              "saldo_obs", "saldo_depois", "preco_unitario", "data",
              "observacao", "obs_len", "id_lancamento", "aviso", "status"]
    for p in plano:
        p["obs_len"] = len(p.get("observacao") or "")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        for p in plano:
            w.writerow({c: ("" if p.get(c) is None else
                            (",".join(map(str, p[c])) if c == "linhas" else p.get(c, "")))
                        for c in campos})

    # snapshot completo (inclui depositos antes/depois) para conferir lotes
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(plano, f, ensure_ascii=False, indent=2, default=str)

    return csv_path, json_path


# ------------------------------------------------------------------ #
#  Orquestracao                                                       #
# ------------------------------------------------------------------ #

def processar_tenant(tenant: str, itens: list[dict], data_api: str,
                     deposito_id, preco_zero: bool, dry_run: bool,
                     confirmado: bool = False, forcar: bool = False,
                     arq_hist: Path | None = None, regravar: bool = False,
                     somente_lancados: bool = False,
                     divergencia: float | None = None) -> None:
    print("\n" + "#" * 84)
    print(f"#  EMPRESA: {tenant.upper()}  ({len(itens)} linha(s) da planilha)")
    print("#" * 84)

    grupos = agrupar_por_sku(itens)
    if len(grupos) < len(itens):
        print(f"  {len(itens)} linhas -> {len(grupos)} SKU(s) unicos "
              f"(quantidades somadas por SKU)")

    saldo_hist = None
    if arq_hist:
        saldo_hist = carregar_saldo_historico(arq_hist, tenant)
        print(f"  Saldo ant: histórico de {arq_hist.name} "
              f"({len(saldo_hist)} SKU(s) para {tenant})")
        if somente_lancados:
            alvo = carregar_skus_com_lancamento(arq_hist, tenant)
            antes = len(grupos)
            grupos = [g for g in grupos if g["sku"] in alvo]
            print(f"  Filtro   : só SKUs com lancamento anterior "
                  f"({len(grupos)} de {antes})")
    if forcar:
        print("  [--forcar] lancando mesmo onde o saldo ja esta no alvo")
    if regravar:
        print("  [--regravar] onde o saldo ja e o alvo, lanca alvo+1 e volta "
              "(2 movimentacoes) para gravar a observacao")

    client = TinyClient(tenant)
    plano = montar_preview(client, grupos, data_api, preco_zero, forcar, saldo_hist,
                           divergencia)
    pendentes = imprimir_plano(plano, tenant, deposito_id, preco_zero)

    if dry_run:
        print("\n  [DRY-RUN] Nenhum lancamento enviado.")
    elif pendentes == 0:
        print(f"\n  [{tenant}] Nada a lancar.")
    else:
        if confirmado:
            print(f"\n  [{tenant}] --confirmar: enviando sem prompt.")
            resp = "CONFIRMO"
        else:
            resp = input(f"\n  [{tenant}] Confirma enviar {pendentes} lancamento(s)? "
                         f"Digite 'CONFIRMO' para prosseguir: ").strip()
        if resp != "CONFIRMO":
            print(f"  [{tenant}] Cancelado. Nenhum lancamento enviado.")
        else:
            aplicar(client, plano, deposito_id, regravar)
            ok = sum(1 for p in plano if p["status"] == "aplicado")
            print(f"\n  [{tenant}] Concluido: {ok}/{pendentes} lancamento(s) aplicado(s).")

    csv_path, json_path = salvar_relatorios(plano, tenant)
    print(f"  Log CSV      : {csv_path}")
    print(f"  Snapshot JSON: {json_path}")


def main():
    args = sys.argv[1:]
    flags = [a for a in args if a.startswith("-")]
    pos = [a for a in args if not a.startswith("-")]

    dry_run = "--dry-run" in flags
    preco_zero = "--preco-zero" in flags
    confirmado = "--confirmar" in flags
    forcar = "--forcar" in flags
    regravar = "--regravar" in flags
    somente_lancados = "--somente-lancados" in flags

    def valor_flag(nome, conv, default=None):
        for f in flags:
            if f.startswith(nome + "="):
                try:
                    return conv(f.split("=", 1)[1])
                except ValueError:
                    print(f"Valor invalido para {nome}: {f}")
                    sys.exit(1)
        if nome in flags:
            i = args.index(nome)
            if i + 1 < len(args):
                try:
                    return conv(args[i + 1])
                except ValueError:
                    print(f"Valor invalido para {nome}: {args[i + 1]}")
                    sys.exit(1)
        return default

    limite = valor_flag("--limite", int)
    deposito_id = valor_flag("--deposito", int)
    data_arg = valor_flag("--data", str)
    aba = valor_flag("--aba", str)
    divergencia = valor_flag("--divergencia", float)
    arq_hist_arg = valor_flag("--saldo-anterior", str)
    arq_hist = Path(arq_hist_arg) if arq_hist_arg else None
    if arq_hist and not arq_hist.exists():
        print(f"Arquivo de saldo anterior nao encontrado: {arq_hist}")
        sys.exit(1)

    if not pos:
        print(__doc__)
        sys.exit(1)

    caminho = Path(pos[0])
    if not caminho.exists():
        print(f"Arquivo nao encontrado: {caminho}")
        sys.exit(1)

    arg_tenant = (pos[1] if len(pos) > 1 else "auto").strip().lower()
    if arg_tenant != "auto" and arg_tenant not in TENANT_NAMES:
        print(f"Tenant desconhecido: {arg_tenant!r}. Use: {', '.join(TENANT_NAMES)} ou 'auto'.")
        sys.exit(1)

    agora = datetime.now()
    if data_arg:
        data_br = _parse_data(data_arg)
        try:
            dia = datetime.strptime(data_br, "%d/%m/%Y")
        except ValueError:
            print(f"Data invalida: {data_arg!r} (use dd/mm/aaaa)")
            sys.exit(1)
    else:
        data_br = agora.strftime("%d/%m/%Y")
        dia = agora

    # A API exige 'data e hora' e SO aceita 'YYYY-MM-DD HH:MM:SS' (separador
    # espaco, sem timezone). Confirmado por teste: ISO com 'T', com 'Z' ou com
    # offset, e a data pura 'YYYY-MM-DD', todos retornam 400 no campo 'data'.
    data_api = dia.replace(hour=agora.hour, minute=agora.minute,
                           second=agora.second, microsecond=0).strftime("%Y-%m-%d %H:%M:%S")

    print(f"\n  Planilha : {caminho.name}")
    print(f"  Data     : {data_br} (enviada como {data_api})")
    print(f"  Modo     : {'DRY-RUN (nao escreve)' if dry_run else 'APLICAR (pede confirmacao)'}")

    itens = ler_planilha(caminho, aba)
    print(f"  {len(itens)} item(ns) validos (SKU x empresa).")
    if not itens:
        print("Nada a fazer.")
        return

    por_coluna = bool(itens[0].get("por_coluna"))

    # Empresa: com colunas por empresa, o argumento da linha de comando FILTRA.
    # Com coluna 'Quant' unica, ele SOBRESCREVE. Sem argumento ('auto'), cada
    # linha precisa trazer a empresa — nunca adivinhamos, porque o mesmo SKU
    # pode existir em mais de uma empresa com ids diferentes.
    if arg_tenant != "auto":
        if por_coluna:
            antes = len(itens)
            itens = [it for it in itens if it["tenant"] == arg_tenant]
            print(f"  Filtro   : apenas {arg_tenant} ({len(itens)} de {antes} itens)")
            if not itens:
                print("Nada a fazer.")
                return
        else:
            for it in itens:
                it["tenant"] = arg_tenant
    elif not por_coluna:
        sem_empresa = [it for it in itens if not it["tenant"]]
        if sem_empresa:
            linhas = ", ".join(str(it["linha"]) for it in sem_empresa[:15])
            print(f"\n  ERRO: {len(sem_empresa)} linha(s) sem empresa reconhecida (linhas {linhas}).")
            print(f"  Valores aceitos na coluna 'Empresa': "
                  f"{', '.join(sorted(set(TENANT_NAMES) | set(TENANT_NAMES.values())))}")
            print("  Ou informe a empresa no comando: "
                  f"python {Path(__file__).name} <planilha> <tenant>")
            sys.exit(1)

    if limite:
        itens = itens[:limite]
        print(f"  [--limite {limite}] processando apenas os {len(itens)} primeiros itens.")

    por_tenant: dict[str, list[dict]] = {}
    for it in itens:
        por_tenant.setdefault(it["tenant"], []).append(it)

    tenants = sorted(por_tenant)
    print(f"  Empresas : {', '.join(tenants)}")

    if not validate_tenants(tenants):
        print("Autenticacao nao concluida. Encerrando.")
        sys.exit(1)

    for tenant in tenants:
        processar_tenant(tenant, por_tenant[tenant], data_api,
                         deposito_id, preco_zero, dry_run, confirmado,
                         forcar, arq_hist, regravar, somente_lancados,
                         divergencia)

    print("\n  Tudo concluido.")


if __name__ == "__main__":
    main()
