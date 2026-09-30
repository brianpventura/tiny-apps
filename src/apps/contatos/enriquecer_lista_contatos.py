"""
Anexa Email e Telefone a qualquer planilha que tenha uma coluna de CPF/CNPJ,
casando contra o cadastro de contatos das empresas de data/tenants.json.

Por que existe (e por que nao usa enriquecer_telefones.py): aquele script busca
contato a contato com GET /contatos?cpfCnpj=..., 1 request por cliente por
empresa. Para uma lista de 4.2 mil clientes isso custa ~8.500 requests (~70 min).
Aqui a coleta e um dump unico da base de contatos — ~1.072 requests (~9 min para
todas as empresas, ~6 min rodando uma por processo) — e o cruzamento e offline,
entao refazer o join com outra regra custa zero request.

DESCOBERTA QUE BARATEIA TUDO (set/2026): a LISTAGEM de /contatos ja devolve
'email', alem de 'telefone', 'celular', 'endereco', 'tipoPessoa' e 'situacao'.
O cache antigo data/contatos_raw_<tenant>.json nao tinha email porque
_scan_telefones.py recortava as chaves ao salvar, nao porque a API nao mandasse.
Este script grava o dump completo no mesmo arquivo (superset do esquema antigo,
os leitores existentes continuam funcionando) e preserva o dump anterior com
sufixo de data.

SEGUNDA DESCOBERTA, e por isso existe a fase 3: na LISTAGEM o campo 'celular'
vem SEMPRE null — mas no DETALHE (GET /contatos/{id}) ele costuma estar
preenchido. Ou seja, "celular vazio em 100% da base" era artefato da listagem,
nao a verdade do cadastro. Em amostra de 10 clientes sem telefone na listagem, 3
tinham celular no detalhe. A fase 3 busca o detalhe SO de quem ficou sem email
ou sem telefone na fase 2 — poucos requests, contra 107 mil se fosse a base
toda. Cuidado medido: quando 'telefone' existe, o 'celular' do detalhe as vezes
e uma copia reformatada e ate corrompida dele (ex.: telefone '11999990000'
virou celular '(11) 9999-0000', sem o 9). Por isso a ordem
telefone -> celular -> telefoneAdicional nunca muda: o celular so e usado quando
nao ha telefone.

Regras do cruzamento (todas offline, fase 2):
  - chave = CPF/CNPJ so digitos;
  - varios cadastros para o mesmo CPF na mesma empresa (medido: quase mil numa empresa):
    vence o de situacao != 'E', depois o mais completo (email+telefone), depois
    o de dataAtualizacao mais recente;
  - entre as empresas, a ordem de preferencia e a coluna 'Empresa' da planilha
    (a empresa de onde o cliente veio) e depois a outra — email e telefone podem
    vir de empresas diferentes, e a coluna 'Fonte contato' registra de onde;
  - cadastro com situacao 'E' (excluido logicamente, e como o Tiny marca quem
    pediu exclusao LGPD) so e usado se nao houver outro, e sai marcado.

Uso:
    # fase 1 — dump da base de contatos (online, ~1 req por 100 contatos por empresa)
    python enriquecer_lista_contatos.py <planilha.xlsx> --fases 1
    python enriquecer_lista_contatos.py <planilha.xlsx> --fases 1 <tenant>   # 1 empresa

    # fase 2 — cruzamento e gravacao (offline, nao valida token)
    python enriquecer_lista_contatos.py <planilha.xlsx> --fases 2 --dry-run
    python enriquecer_lista_contatos.py <planilha.xlsx> --fases 2

    # fase 3 — detalhe so dos vazios (online) e regravacao
    python enriquecer_lista_contatos.py <planilha.xlsx> --fases 3

    # tudo de uma vez
    python enriquecer_lista_contatos.py <planilha.xlsx> --fases 1,2,3

Flags: --col-doc (default 'CPF/CNPJ'), --col-empresa (default 'Empresa'),
--aba, --dry-run, --fases, --sem-backup.
"""

import argparse
import io
import json
import re
import shutil
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)

_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

from core.auth_tiny import validate_tenants  # noqa: E402
from core.produtos_utils import desempacotar_data  # noqa: E402
from core.tiny_client import TinyClient  # noqa: E402

from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS = get_all_tenant_ids()  # chaves de data/tenants.json
DATA = _SRC.parent / "data"

# campos guardados do item de listagem (superset do cache antigo)
CAMPOS = ("id", "nome", "cpfCnpj", "telefone", "celular", "email",
          "situacao", "tipoPessoa", "dataAtualizacao")


def _digitos(v) -> str:
    return re.sub(r"\D", "", str(v or ""))


def _txt(v) -> str:
    return str(v).strip() if v is not None else ""


def melhor_telefone(c: dict) -> str:
    """Primeiro nao-vazio entre telefone -> celular -> telefoneAdicional -> fone."""
    for campo in ("telefone", "celular", "telefoneAdicional", "fone"):
        v = _txt(c.get(campo))
        if v:
            return v
    return ""


# ------------------------------------------------------------------ #
#  Fase 1 — dump da base de contatos (online)                         #
# ------------------------------------------------------------------ #

def fase1(tenants: list[str], sem_backup: bool = False) -> None:
    for tenant in tenants:
        destino = DATA / f"contatos_raw_{tenant}.json"
        if destino.exists() and not sem_backup:
            try:
                marca = datetime.fromtimestamp(destino.stat().st_mtime).strftime("%Y%m%d")
                bkp = DATA / f"contatos_raw_{tenant}_{marca}.json"
                if not bkp.exists():
                    shutil.copy2(destino, bkp)
                    print(f"  [{tenant}] dump anterior preservado em {bkp.name}")
            except OSError as e:
                print(f"  [{tenant}] AVISO: nao foi possivel preservar o dump anterior ({e})")

        cli = TinyClient(tenant)
        total = {"n": 0}
        registros = []
        print(f"  [{tenant}] varrendo /contatos...", flush=True)
        for it in cli.contatos_todos(on_total=lambda t: total.__setitem__("n", t)):
            registros.append({k: it.get(k) for k in CAMPOS})
            if len(registros) % 5000 == 0:
                print(f"  [{tenant}] {len(registros)}/{total['n']}...", flush=True)
        destino.write_text(json.dumps(registros, ensure_ascii=False), encoding="utf-8")

        com_tel = sum(1 for c in registros if melhor_telefone(c))
        com_mail = sum(1 for c in registros if _txt(c.get("email")))
        n = len(registros) or 1
        print(f"  [{tenant}] {len(registros)} contatos (API informou {total['n']}) -> {destino.name}")
        print(f"  [{tenant}] com telefone: {com_tel} ({com_tel*100//n}%) | "
              f"com email: {com_mail} ({com_mail*100//n}%)")


# ------------------------------------------------------------------ #
#  Fase 2 — cruzamento offline e gravacao                             #
# ------------------------------------------------------------------ #

def _peso(c: dict) -> tuple:
    """Ordem de preferencia entre cadastros do mesmo CPF na mesma empresa."""
    vivo = 0 if _txt(c.get("situacao")).upper() == "E" else 1
    completude = bool(_txt(c.get("email"))) + bool(melhor_telefone(c))
    return (vivo, completude, _txt(c.get("dataAtualizacao")))


def _carregar(tenant: str) -> dict[str, list[dict]]:
    """CPF/CNPJ so digitos -> cadastros daquela empresa, do melhor para o pior.

    Guarda TODOS os cadastros (CPF com mais de um cadastro e comum): o segundo
    cadastro as vezes e o unico com telefone.
    """
    f = DATA / f"contatos_raw_{tenant}.json"
    if not f.exists():
        print(f"  AVISO: {f.name} nao existe — rode a fase 1.")
        return {}
    dados = json.loads(f.read_text(encoding="utf-8"))
    if dados and "email" not in dados[0]:
        print(f"  AVISO: {f.name} e do esquema antigo (sem email) — rode a fase 1.")

    mapa: dict[str, list[dict]] = {}
    for c in dados:
        doc = _digitos(c.get("cpfCnpj"))
        if doc:
            mapa.setdefault(doc, []).append(c)
    for lista in mapa.values():
        lista.sort(key=_peso, reverse=True)
    return mapa


def _carregar_detalhes() -> dict[str, dict]:
    """Cache da fase 3: 'tenant:id' -> detalhe completo do contato."""
    f = DATA / "contatos_detalhe.json"
    if not f.exists():
        return {}
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _resolver(doc: str, ordem: list[str], base: dict, det: dict) -> tuple:
    """Retorna (email, telefone, fontes, so_excluido, candidatos).

    Varre os cadastros na ordem de preferencia de empresa e, dentro de cada um,
    a listagem e depois o detalhe (fase 3). email cai para emailNfe; telefone
    segue telefone -> celular -> telefoneAdicional, nessa ordem e nunca outra.
    """
    email = tel = ""
    fonte_email = fonte_tel = ""
    candidatos, vivos = [], []
    for t in ordem:
        for c in base.get(t, {}).get(doc, []):
            candidatos.append((t, c))
            vivos.append(_txt(c.get("situacao")).upper() != "E")
    for t, c in candidatos:
        for src in (c, det.get(f"{t}:{c.get('id')}")):
            if not src:
                continue
            if not email:
                e = _txt(src.get("email")) or _txt(src.get("emailNfe"))
                if e:
                    email, fonte_email = e, t
            if not tel:
                p = melhor_telefone(src)
                if p:
                    tel, fonte_tel = p, t
    fontes = sorted({f for f in (fonte_email, fonte_tel) if f})
    so_excluido = bool(candidatos) and not any(vivos)
    return email, tel, fontes, so_excluido, candidatos


def fase2(caminho: Path, col_doc: str, col_empresa: str, aba: str | None,
          dry_run: bool, sem_backup: bool) -> list[tuple]:
    """Cruza e grava. Devolve os candidatos das linhas que ficaram incompletas
    (entrada da fase 3)."""
    base = {t: _carregar(t) for t in TENANTS}
    det = _carregar_detalhes()
    for t in TENANTS:
        print(f"  cadastro {t}: {len(base[t])} documentos distintos")
    if det:
        print(f"  detalhes em cache: {len(det)}")

    wb = openpyxl.load_workbook(caminho)
    ws = wb[aba] if aba else wb.active
    cab = [_txt(c.value) for c in ws[1]]

    def indice(nome: str) -> int | None:
        alvo = nome.strip().lower()
        for i, h in enumerate(cab, start=1):
            if h.strip().lower() == alvo:
                return i
        return None

    i_doc = indice(col_doc)
    if i_doc is None:
        print(f"ERRO: coluna '{col_doc}' nao encontrada. Cabecalho: {cab}")
        sys.exit(1)
    i_emp = indice(col_empresa)

    # colunas de saida: reaproveita se ja existirem
    saidas = {}
    for nome in ("Email", "Telefone", "Fonte contato"):
        j = indice(nome)
        if j is None:
            j = len(cab) + 1
            cab.append(nome)
            ws.cell(row=1, column=j, value=nome)
        saidas[nome] = j

    est = {"linhas": 0, "sem_doc": 0, "achou": 0, "nao_achou": 0,
           "com_email": 0, "com_tel": 0, "so_excluido": 0, "cruzou_empresa": 0}
    faltantes = []
    pendentes: list[tuple] = []

    for lin in range(2, ws.max_row + 1):
        doc = _digitos(ws.cell(row=lin, column=i_doc).value)
        if not doc:
            est["sem_doc"] += 1
            continue
        est["linhas"] += 1

        emp_lin = _txt(ws.cell(row=lin, column=i_emp).value).lower() if i_emp else ""
        ordem = [t for t in TENANTS if t in emp_lin]
        ordem += [t for t in TENANTS if t not in ordem]

        email, tel, fontes, so_excluido, candidatos = _resolver(doc, ordem, base, det)

        if candidatos:
            est["achou"] += 1
        else:
            est["nao_achou"] += 1
            faltantes.append((doc, _txt(ws.cell(row=lin, column=i_doc + 1).value)))
        if not email or not tel:
            pendentes.extend(candidatos)

        fonte = "; ".join(fontes)
        if len(fontes) > 1:
            est["cruzou_empresa"] += 1
        if so_excluido:
            fonte = (fonte + " (cadastro excluido)").strip() if fonte else "cadastro excluido"
            est["so_excluido"] += 1
        if email:
            est["com_email"] += 1
        if tel:
            est["com_tel"] += 1

        ws.cell(row=lin, column=saidas["Email"], value=email or None)
        ws.cell(row=lin, column=saidas["Telefone"], value=tel or None)
        ws.cell(row=lin, column=saidas["Fonte contato"], value=fonte or None)

    n = est["linhas"] or 1
    print()
    print(f"  linhas com documento          : {est['linhas']}")
    if est["sem_doc"]:
        print(f"  linhas SEM documento (puladas): {est['sem_doc']}")
    print(f"  com cadastro no Tiny          : {est['achou']} ({est['achou']*100//n}%)")
    print(f"  sem cadastro em nenhuma empresa: {est['nao_achou']}")
    print(f"  EMAIL preenchido              : {est['com_email']} ({est['com_email']*100//n}%)")
    print(f"  TELEFONE preenchido           : {est['com_tel']} ({est['com_tel']*100//n}%)")
    print(f"  contato de mais de 1 empresa  : {est['cruzou_empresa']}")
    print(f"  so cadastro excluido (sit=E)  : {est['so_excluido']}")
    if faltantes:
        print(f"  exemplos sem cadastro: {faltantes[:5]}")

    if dry_run:
        print("\n  [DRY-RUN] nada gravado. Rode sem --dry-run para escrever.")
        return pendentes

    if not sem_backup:
        bkp = caminho.with_name(caminho.stem + "_ORIGINAL.xlsx")
        if not bkp.exists():
            shutil.copy2(caminho, bkp)
            print(f"\n  backup: {bkp.name}")

    fill = PatternFill("solid", fgColor="DDDDDD")
    for nome, j in saidas.items():
        c = ws.cell(row=1, column=j)
        c.font = Font(bold=True)
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
    for j, larg in ((saidas["Email"], 34), (saidas["Telefone"], 20), (saidas["Fonte contato"], 24)):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = larg
    if ws.auto_filter.ref:
        ws.auto_filter.ref = "A1:%s%d" % (ws.cell(row=1, column=len(cab)).column_letter, ws.max_row)

    try:
        wb.save(caminho)
    except PermissionError:
        alt = caminho.with_name(caminho.stem + f"_{datetime.now():%H%M%S}.xlsx")
        wb.save(alt)
        print(f"  ARQUIVO ABERTO NO EXCEL — gravado em {alt.name}")
        return pendentes
    print(f"  gravado: {caminho}")
    return pendentes


# ------------------------------------------------------------------ #
#  Fase 3 — detalhe do contato so para os vazios (online)             #
# ------------------------------------------------------------------ #

def fase3(pendentes: list[tuple]) -> int:
    """Busca GET /contatos/{id} dos cadastros das linhas incompletas.

    So existe por causa do 'celular': a listagem devolve null nesse campo e o
    detalhe nao. Cache por 'tenant:id', entao repetir custa zero request.
    """
    det = _carregar_detalhes()
    alvos = []
    vistos = set()
    for t, c in pendentes:
        cid = c.get("id")
        chave = f"{t}:{cid}"
        if not cid or chave in det or chave in vistos:
            continue
        vistos.add(chave)
        alvos.append((t, cid, chave))
    print(f"  cadastros a detalhar: {len(alvos)} (cache ja tem {len(det)})")
    if not alvos:
        return 0

    clientes = {t: TinyClient(t) for t in {a[0] for a in alvos}}
    feitos = erros = 0

    def buscar(alvo):
        t, cid, chave = alvo
        return chave, desempacotar_data(clientes[t].obter_contato(cid))

    with ThreadPoolExecutor(max_workers=4) as pool:
        futuros = {pool.submit(buscar, a): a for a in alvos}
        for fut in as_completed(futuros):
            try:
                chave, dados = fut.result()
                det[chave] = {k: dados.get(k) for k in
                              ("id", "nome", "cpfCnpj", "telefone", "celular",
                               "telefoneAdicional", "email", "emailNfe", "situacao")}
                feitos += 1
            except Exception as e:
                erros += 1
                if erros <= 3:
                    print(f"    erro em {futuros[fut][2]}: {e}")
            if (feitos + erros) % 200 == 0:
                print(f"  {feitos + erros}/{len(alvos)}...", flush=True)

    (DATA / "contatos_detalhe.json").write_text(
        json.dumps(det, ensure_ascii=False), encoding="utf-8")
    print(f"  detalhes coletados: {feitos} (erros: {erros}) -> data/contatos_detalhe.json")
    return feitos


def main():
    ap = argparse.ArgumentParser(description="Anexa Email e Telefone a partir do cadastro do Tiny.")
    ap.add_argument("planilha", type=Path)
    ap.add_argument("tenants", nargs="*", default=None,
                    help="empresas da fase 1 (default: todas)")
    ap.add_argument("--fases", default="1,2")
    ap.add_argument("--col-doc", default="CPF/CNPJ")
    ap.add_argument("--col-empresa", default="Empresa")
    ap.add_argument("--aba", default=None)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--sem-backup", action="store_true")
    args = ap.parse_args()

    fases = {int(x) for x in args.fases.replace(" ", "").split(",") if x}
    tenants = args.tenants or TENANTS
    for t in tenants:
        if t not in TENANTS:
            print(f"ERRO: tenant desconhecido '{t}'")
            sys.exit(1)
    if not args.planilha.exists():
        print(f"ERRO: planilha nao encontrada: {args.planilha}")
        sys.exit(1)

    print("=" * 62)
    print("  ENRIQUECER LISTA COM EMAIL E TELEFONE DO CADASTRO")
    print("=" * 62)

    # so valida token quando ha fase online (a 2 tem de rodar com token morto)
    if fases & {1, 3}:
        if not validate_tenants(tenants, interactive=False):
            print("ERRO: token invalido. Rode src/apps/auth/_reauth.py <tenant>.")
            sys.exit(1)

    if 1 in fases:
        print("\n[FASE 1] dump do cadastro de contatos")
        fase1(tenants, args.sem_backup)

    pendentes = []
    if fases & {2, 3}:
        print("\n[FASE 2] cruzamento offline")
        pendentes = fase2(args.planilha, args.col_doc, args.col_empresa, args.aba,
                          args.dry_run, args.sem_backup)

    if 3 in fases:
        print("\n[FASE 3] detalhe dos contatos sem email/telefone")
        if fase3(pendentes):
            print("\n[FASE 2] recruzamento com os detalhes")
            fase2(args.planilha, args.col_doc, args.col_empresa, args.aba,
                  args.dry_run, args.sem_backup)


if __name__ == "__main__":
    main()
