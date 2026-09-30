"""
Grava no Tiny uma lista de precos lida de planilha: promocional ou de venda.

Le uma planilha com colunas de SKU e de preco e escreve APENAS um campo de preco
por execucao -- por padrao `precos.precoPromocional`. O preco de venda
(`precos.preco`) nao e tocado: encerrada a acao promocional, basta zerar o
promocional (`--zerar`) para o produto voltar ao preco normal.

Serve para qualquer lista de preco (acao promocional, reajuste, reversao de
precos vinda do export da loja virtual): as colunas de SKU e de preco sao
configuraveis (--col-sku / --col-preco / --aba).

Uso:
    python atualizar_precos.py <planilha.xlsx> [tenant] [--commit] [opcoes]

    tenant   : tenant_id de data/tenants.json | todas   (default: todas)
    --commit : aplica de verdade. SEM ele nada e escrito (dry-run e o padrao).

Opcoes:
    --sku A,B,C      processa so esses SKUs (use para o teste inicial em 1-2 produtos)
    --limite N       processa so os N primeiros SKUs com preco (depois da deduplicacao)
    --zerar          grava 0 no promocional em vez do preco da planilha (encerra a acao)
    --campo X        'promocional' (default, grava precos.precoPromocional) ou
                     'venda' (grava precos.preco). Um de cada vez, nunca os dois.
    --sem-verificar  nao rele o produto depois do PUT (mais rapido, menos seguro)
    --col-sku NOME   coluna de SKU    (default: 'Codigo (SKU)' com acento, do export do Tiny)
    --col-preco NOME coluna de preco  (default: 'preco')
    --aba NOME       aba da planilha  (default: 'Produtos' se existir, senao a primeira)
    --log-prefixo X  prefixo do log CSV em reports/ (default: 'precos')

Campo gravado:
    Por padrao so `precos.precoPromocional` e tocado. `--campo venda` grava
    `precos.preco`; serve, por exemplo, para alinhar o preco cheio do ERP ao da
    tabela quando o promocional acabou ficando ACIMA do preco de venda (o
    snapshot de `relatorios/precos_promocionais.py` aponta esses casos).
    `--zerar` e recusado com `--campo venda`: zerar preco de venda nao encerra
    acao nenhuma, deixa o produto a R$ 0,00.

    O log CSV traz a coluna `campo` e os valores em `valor_atual` /
    `valor_novo` / `valor_relido`.

SKU repetido na planilha:
    Export do Shopify traz o mesmo SKU em mais de uma linha (mesmo produto ACTIVE e uma
    copia ARCHIVED/DRAFT), e os precos das duas linhas podem DIVERGIR -- a copia de
    rascunho as vezes traz preco de gaveta, abaixo do vigente. A linha ACTIVE vence. Se
    nao houver coluna de status, ou sobrar mais de uma ACTIVE com precos diferentes, o
    SKU e DESCARTADO e listado: gravar preco ambiguo no ERP e pior que deixar o produto
    de fora.

Fluxo (por empresa, seguindo o padrao de escrita do repo):
    1. Le a planilha; linhas sem preco sao IGNORADAS (nao zeram nada).
    2. Para cada SKU: localiza o produto e faz GET completo.
    3. Altera so o campo alvo; nenhum outro campo do produto e mexido.
    4. Mostra preview (atual -> novo) e pede confirmacao.
    5. PUT e, por padrao, releitura para conferir se o valor persistiu.
    6. Salva log CSV em reports/ -- que serve de BACKUP para reverter a acao.

Variacoes:
    Variacao NAO aceita PUT /produtos/{id} (da 404); o unico caminho e
    PUT /produtos/{pai}/variacoes/{id}, e o id do pai vem de produtoPai.id no GET da
    filha -- o SKU do pai nem sempre esta no export. A deteccao aqui e a mesma de
    gerenciar_produtos/atualizar_custo.py. Confirmado empiricamente que
    precos.precoCusto e precos.precoPromocional persistem nesse endpoint; mesmo
    assim a releitura vem ligada por padrao: se o Tiny aceitar o PUT e ignorar o
    campo, o log mostra 'nao_persistiu' em vez de dar sucesso falso. Rode primeiro
    em um SKU de cada tipo com --sku.
"""

import csv
import sys
from datetime import datetime
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import pandas as pd

from core.auth_tiny import validate_tenants
from core.tiny_client import TinyClient
from core import desempacotar_data as _desempacotar
from core.produtos_utils import preparar_payload_put

REPORTS_DIR = _SRC.parent / "reports"
from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS_PADRAO = get_all_tenant_ids()  # chaves de data/tenants.json

COL_SKU = "Código (SKU)"
COL_PRECO = "preco"
#: coluna de status do export do Shopify, usada so para desempatar SKU repetido
COL_STATUS = "product_status"

#: diferenca abaixo da qual o campo ja esta no valor alvo (ruido de float)
_TOLERANCIA = 0.005

#: apelido de CLI -> chave dentro de `precos` no JSON do produto
CAMPOS = {"promocional": "precoPromocional", "venda": "preco"}


# --------------------------------------------------------------------------- #
#  Leitura da planilha                                                         #
# --------------------------------------------------------------------------- #

def _parse_preco(valor) -> float | None:
    """Aceita 99.99, 'R$ 99,99', '99,99' e '1.234,56'. Vazio/invalido -> None."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).replace("R$", "").replace("\xa0", "").replace(" ", "").strip()
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") \
            else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


#: colunas de descricao aceitas, em ordem de preferencia (so para o preview/log)
_COLS_DESCRICAO = ("Descrição", "descricao", "Produto", "produto", "Nome", "nome")


def _descricao_da_linha(r) -> str:
    for c in _COLS_DESCRICAO:
        if c in r.index:
            v = str(r[c]).strip()
            if v and v.lower() != "nan":
                return v
    return ""


def _resolver_duplicados(linhas: list[dict]) -> tuple[list[dict], list[dict]]:
    """SKU repetido na planilha: vence a linha ACTIVE; preco ambiguo e descartado.

    Retorna (itens, ambiguos). Ver a secao 'SKU repetido' na docstring de topo.
    """
    por_sku: dict[str, list[dict]] = {}
    for it in linhas:
        por_sku.setdefault(it["sku"].upper(), []).append(it)

    itens: list[dict] = []
    ambiguos: list[dict] = []
    for _, grupo in por_sku.items():
        if len(grupo) == 1:
            itens.append(grupo[0])
            continue
        ativos = [g for g in grupo if g.get("status_planilha", "").upper() == "ACTIVE"]
        candidatos = ativos or grupo
        precos = {g["preco_novo"] for g in candidatos}
        if len(candidatos) == 1 or len(precos) == 1:
            itens.append(candidatos[0])
        else:
            ambiguos.append({**candidatos[0],
                             "precos": sorted(precos),
                             "linhas": [g["linha"] for g in candidatos]})

    itens.sort(key=lambda it: it["linha"])
    ambiguos.sort(key=lambda it: it["linha"])
    return itens, ambiguos


def ler_planilha(caminho: Path, skus_alvo: set[str] | None, limite: int | None,
                 col_sku: str = COL_SKU, col_preco: str = COL_PRECO,
                 aba: str | None = None) -> list[dict]:
    """Linhas com SKU e preco validos. Linha sem preco e ignorada de proposito."""
    df = pd.read_excel(caminho, sheet_name=aba if aba else _aba_padrao(caminho))
    for coluna in (col_sku, col_preco):
        if coluna not in df.columns:
            raise ValueError(
                f"A planilha precisa da coluna '{coluna}'. Colunas encontradas: "
                f"{list(df.columns)}"
            )

    linhas: list[dict] = []
    for i, r in df.iterrows():
        sku = str(r[col_sku]).strip()
        if not sku or sku.lower() == "nan":
            continue
        if skus_alvo is not None and sku.upper() not in skus_alvo:
            continue
        preco = _parse_preco(r[col_preco])
        if preco is None:
            continue
        if preco < 0:
            print(f"  [linha {i + 2}] ignorada (preco negativo: {preco})")
            continue
        status = str(r[COL_STATUS]).strip() if COL_STATUS in df.columns else ""
        linhas.append({"linha": i + 2, "sku": sku, "preco_novo": round(preco, 2),
                       "descricao": _descricao_da_linha(r),
                       "status_planilha": "" if status.lower() == "nan" else status})

    itens, ambiguos = _resolver_duplicados(linhas)
    if len(itens) != len(linhas):
        print(f"  {len(linhas) - len(itens) - len(ambiguos)} linha(s) duplicada(s) "
              "descartada(s) pelo desempate por status ACTIVE.")
    if ambiguos:
        print(f"  ATENCAO: {len(ambiguos)} SKU(s) repetidos com precos DIFERENTES e sem "
              "linha ACTIVE unica — DESCARTADOS (nenhum preco sera gravado neles):")
        for a in ambiguos:
            print(f"     {a['sku']}: precos {a['precos']} nas linhas {a['linhas']}")

    return itens[:limite] if limite else itens


def _aba_padrao(caminho: Path):
    """'Produtos' quando a planilha tem essa aba; senao a primeira."""
    try:
        return "Produtos" if "Produtos" in pd.ExcelFile(caminho).sheet_names else 0
    except Exception:
        return 0


# --------------------------------------------------------------------------- #
#  Preview                                                                     #
# --------------------------------------------------------------------------- #

def _valor(precos, chave: str) -> float | None:
    if not isinstance(precos, dict):
        return None
    try:
        v = precos.get(chave)
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def montar_plano(client: TinyClient, itens: list[dict], zerar: bool,
                 chave: str = "precoPromocional") -> list[dict]:
    """Para cada SKU: localiza, faz GET completo e prepara o payload do PUT."""
    plano: list[dict] = []
    print(f"\n  Buscando {len(itens)} produto(s) no Tiny...")

    for idx, item in enumerate(itens, start=1):
        sku = item["sku"]
        alvo = 0.0 if zerar else item["preco_novo"]
        print(f"  [{idx}/{len(itens)}] SKU={sku} ... ", end="", flush=True)

        try:
            resumo = client.buscar_produto_por_sku(sku)
        except Exception as e:
            print(f"ERRO busca ({e})")
            plano.append({**item, "status": f"erro_busca: {e}"})
            continue
        if not resumo:
            print("NAO ENCONTRADO nesta empresa")
            plano.append({**item, "status": "nao_encontrado"})
            continue

        produto_id = resumo.get("id")
        try:
            full = _desempacotar(client.obter_produto(produto_id))
        except Exception as e:
            print(f"ERRO GET ({e})")
            plano.append({**item, "id": produto_id, "status": f"erro_get: {e}"})
            continue

        atual = _valor(full.get("precos"), chave)
        tipo_var = str(full.get("tipoVariacao") or "").upper()
        registro = {**item, "id": produto_id, "campo": chave, "valor_atual": atual,
                    "valor_novo": alvo, "tipo": "normal", "alvo_id": produto_id}

        if tipo_var == "V":
            pai_id = (full.get("produtoPai") or {}).get("id")
            if not pai_id:
                print("ERRO (variacao sem produtoPai)")
                plano.append({**registro, "tipo": "variacao",
                              "status": "erro: variacao sem produtoPai"})
                continue
            registro.update(tipo="variacao", alvo_id=pai_id)

        if atual is not None and abs(atual - alvo) < _TOLERANCIA:
            print("OK (ja esta correto)")
            plano.append({**registro, "status": "sem_mudanca"})
            continue

        payload = preparar_payload_put(full)
        payload.setdefault("precos", {})[chave] = alvo
        extra = f" [variacao, pai={registro['alvo_id']}]" if registro["tipo"] == "variacao" else ""
        print(f"OK ({_fmt(atual)} -> {_fmt(alvo)}){extra}")
        plano.append({**registro, "payload": payload, "status": "pendente"})

    return plano


def _fmt(v) -> str:
    return f"R$ {v:.2f}" if isinstance(v, (int, float)) else "(vazio)"


def imprimir_plano(plano: list[dict], tenant: str, rotulo: str = "preco promocional") -> int:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    sem_mudanca = [p for p in plano if p["status"] == "sem_mudanca"]
    nao_achados = [p for p in plano if p["status"] == "nao_encontrado"]
    erros = [p for p in plano if p["status"] not in
             ("pendente", "sem_mudanca", "nao_encontrado")]

    print("\n" + "=" * 78)
    print(f"  PLANO — {rotulo.upper()} — {tenant.upper()}")
    print("=" * 78)

    if pendentes:
        variacoes = sum(1 for p in pendentes if p["tipo"] == "variacao")
        print(f"\n  {len(pendentes)} produto(s) a atualizar "
              f"({len(pendentes) - variacoes} normais + {variacoes} variacoes):\n")
        for p in pendentes[:25]:
            marca = " [variacao]" if p["tipo"] == "variacao" else ""
            print(f"   {p['sku']:>16}  {p['descricao'][:42]:<42} "
                  f"{_fmt(p['valor_atual']):>10} -> {_fmt(p['valor_novo']):>10}{marca}")
        if len(pendentes) > 25:
            print(f"   ... e mais {len(pendentes) - 25} (detalhe completo no log CSV)")

    if sem_mudanca:
        print(f"\n  {len(sem_mudanca)} ja estao com o {rotulo} correto (ignorados)")
    if nao_achados:
        print(f"  {len(nao_achados)} SKU(s) nao existem nesta empresa (ignorados)")
    if erros:
        print(f"\n  {len(erros)} com problema:")
        for p in erros[:15]:
            print(f"   {p['sku']}: {p['status']}")

    print("=" * 78)
    return len(pendentes)


# --------------------------------------------------------------------------- #
#  Aplicacao                                                                   #
# --------------------------------------------------------------------------- #

def aplicar(client: TinyClient, plano: list[dict], verificar: bool,
            chave: str = "precoPromocional") -> None:
    pendentes = [p for p in plano if p["status"] == "pendente"]
    print(f"\n  Aplicando {len(pendentes)} PUT(s)"
          f"{' com releitura de conferencia' if verificar else ''}...")

    for i, p in enumerate(pendentes, 1):
        print(f"  [{i}/{len(pendentes)}] SKU={p['sku']} ... ", end="", flush=True)
        try:
            if p["tipo"] == "variacao":
                client.atualizar_variacao(p["alvo_id"], p["id"], p["payload"])
            else:
                client.atualizar_produto(p["id"], p["payload"])
        except Exception as e:
            p["status"] = f"erro_put: {e}"
            print(f"ERRO ({e})")
            continue

        if not verificar:
            p["status"] = "atualizado"
            print("OK")
            continue

        try:
            relido = _desempacotar(client.obter_produto(p["id"]))
            gravado = _valor(relido.get("precos"), chave)
        except Exception as e:
            p["status"] = "atualizado_sem_conferir"
            print(f"OK (releitura falhou: {e})")
            continue

        p["valor_relido"] = gravado
        if gravado is not None and abs(gravado - p["valor_novo"]) < _TOLERANCIA:
            p["status"] = "atualizado"
            print("OK (conferido)")
        else:
            # o Tiny aceitou o PUT mas ignorou o campo — nao contar como sucesso
            p["status"] = "nao_persistiu"
            print(f"FALHOU: o Tiny devolveu {_fmt(gravado)} apos o PUT")


def salvar_log(plano: list[dict], tenant: str, commit: bool,
               prefixo: str = "precos") -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    modo = "commit" if commit else "dryrun"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"{prefixo}_{modo}_{tenant}_{stamp}.csv"

    campos = ["linha", "sku", "descricao", "id", "tipo", "alvo_id",
              "campo", "valor_atual", "valor_novo", "valor_relido", "status"]
    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        for p in plano:
            w.writerow({c: p.get(c, "") for c in campos})
    return arquivo


def processar_tenant(tenant: str, itens: list[dict], commit: bool,
                     zerar: bool, verificar: bool,
                     prefixo_log: str = "precos",
                     campo: str = "promocional") -> None:
    print("\n" + "#" * 78)
    print(f"#  EMPRESA: {tenant.upper()}")
    print("#" * 78)

    chave = CAMPOS[campo]
    rotulo = "preco promocional" if campo == "promocional" else "preco de venda"

    client = TinyClient(tenant)
    plano = montar_plano(client, itens, zerar, chave)
    pendentes = imprimir_plano(plano, tenant, rotulo)

    if pendentes == 0:
        print(f"\n  [{tenant}] Nada a aplicar.")
    elif not commit:
        print(f"\n  [{tenant}] DRY-RUN: nenhuma alteracao foi enviada ao Tiny.")
        print("  Para aplicar de verdade, repita o comando com --commit.")
    else:
        resposta = input(f"\n  [{tenant}] Confirma gravar o {rotulo} em {pendentes} "
                         "produto(s)? (s/N): ").strip().lower()
        if resposta != "s":
            print(f"  [{tenant}] Cancelado. Nada foi alterado.")
        else:
            aplicar(client, plano, verificar, chave)

    log = salvar_log(plano, tenant, commit and pendentes > 0, prefixo_log)
    ok = sum(1 for p in plano if p["status"] == "atualizado")
    falhou = [p for p in plano if p["status"] == "nao_persistiu"]
    if commit:
        print(f"\n  [{tenant}] {ok} atualizado(s) e conferido(s).")
        if falhou:
            print(f"  [{tenant}] ATENCAO: {len(falhou)} PUT(s) aceitos mas o valor NAO "
                  "persistiu no Tiny:")
            for p in falhou[:10]:
                print(f"     {p['sku']} ({p['tipo']})")
            print("  Nao trate esses como aplicados.")
    print(f"  Log salvo em: {log}")


def main() -> None:
    argumentos = sys.argv[1:]
    flags = [a for a in argumentos if a.startswith("--")]
    posicionais = [a for a in argumentos if not a.startswith("--")]

    commit = "--commit" in flags
    zerar = "--zerar" in flags
    verificar = "--sem-verificar" not in flags

    def valor_de(nome: str) -> str | None:
        for f in flags:
            if f.startswith(f"{nome}="):
                return f.split("=", 1)[1]
        if nome in flags:
            i = argumentos.index(nome)
            if i + 1 < len(argumentos) and not argumentos[i + 1].startswith("--"):
                return argumentos[i + 1]
        return None

    skus_txt = valor_de("--sku")
    skus_alvo = ({s.strip().upper() for s in skus_txt.split(",") if s.strip()}
                 if skus_txt else None)
    limite_txt = valor_de("--limite")
    limite = int(limite_txt) if limite_txt and limite_txt.isdigit() else None
    col_sku = valor_de("--col-sku") or COL_SKU
    col_preco = valor_de("--col-preco") or COL_PRECO
    aba = valor_de("--aba")
    prefixo_log = valor_de("--log-prefixo") or "precos"
    campo = (valor_de("--campo") or "promocional").strip().lower()
    if campo not in CAMPOS:
        print(f"  --campo aceita {' ou '.join(CAMPOS)}; recebido: {campo!r}")
        sys.exit(1)
    if zerar and campo != "promocional":
        print("  --zerar so vale para o promocional: zerar o preco de venda deixaria o "
              "produto a R$ 0,00, o que nao e encerramento de acao nenhuma.")
        sys.exit(1)

    # valor de flag nunca e posicional (o parser aceita "--flag valor" alem de "--flag=valor")
    consumidos = {skus_txt, limite_txt, aba, valor_de("--col-sku"),
                  valor_de("--col-preco"), valor_de("--log-prefixo"),
                  valor_de("--campo")} - {None}
    posicionais = [p for p in posicionais if p not in consumidos]

    if not posicionais:
        print(__doc__)
        sys.exit(1)

    caminho = Path(posicionais[0])
    if not caminho.exists():
        print(f"Planilha nao encontrada: {caminho}")
        sys.exit(1)

    arg_tenant = (posicionais[1] if len(posicionais) > 1 else "todas").strip().lower()
    tenants = (list(TENANTS_PADRAO) if arg_tenant in ("todas", "ambas", "all", "")
               else [t.strip() for t in arg_tenant.split(",") if t.strip()])

    print(f"\n  Planilha: {caminho.name}")
    print(f"  Colunas : SKU='{col_sku}'  preco='{col_preco}'"
          + (f"  aba='{aba}'" if aba else ""))
    print(f"  Campo   : precos.{CAMPOS[campo]}"
          + ("   <-- PRECO DE VENDA, nao o promocional" if campo == "venda" else ""))
    print(f"  Modo    : {'COMMIT (grava no Tiny)' if commit else 'DRY-RUN (nada e gravado)'}")
    if zerar:
        print("  Acao    : ZERAR o preco promocional (encerramento da acao)")
    if skus_alvo:
        print(f"  SKUs    : {', '.join(sorted(skus_alvo))}")
    if limite:
        print(f"  Limite  : {limite} linha(s)")

    itens = ler_planilha(caminho, skus_alvo, limite, col_sku, col_preco, aba)
    print(f"  {len(itens)} SKU(s) com preco valido.")
    if not itens:
        print("  Nada a fazer.")
        return

    if not validate_tenants(tenants, interactive=False):
        print("  Autenticacao nao concluida. Rode src/apps/auth/_reauth.py <tenant>.")
        sys.exit(1)

    for tenant in tenants:
        processar_tenant(tenant, itens, commit, zerar, verificar, prefixo_log, campo)

    print("\n  Concluido.")


if __name__ == "__main__":
    main()
