"""Exclui contatos com dados ocultos (LGPD) no Tiny.

Contatos importados de marketplaces com regras de LGPD tem os dados salvos de
forma mascarada — o nome vem com asteriscos (ex: "A*****a", "9*****2", "****")
e cpfCnpj/telefone vazios. Esses cadastros travam o sistema e devem ser
removidos.

Definicao de "oculto": o campo `nome` contem o caractere '*'.

Uso:
    # dry-run (padrao) — apenas lista e conta, NAO exclui:
    python excluir_ocultos.py                 # todos os tenants
    python excluir_ocultos.py <tenant>        # so um tenant
    python excluir_ocultos.py --situacao A    # so ativos (A) / bloqueados (B) / todos (default)

    # execucao real (IRREVERSIVEL):
    python excluir_ocultos.py --executar
    python excluir_ocultos.py <tenant> --executar

Comportamento:
    - Varre a base de contatos ao vivo e identifica os mascarados.
    - Dry-run: grava reports/ocultos_<tenant>_<ts>.csv com os candidatos.
    - --executar: exclui LOGICAMENTE via PUT /contatos/{id} com situacao='E'
      (a API v3 nao tem DELETE de cadastro de contato), com rate limiting.
      * Checkpoint resumivel em data/excluir_ocultos_ck_<tenant>.json — ao
        reexecutar, pula os ja processados.
      * Falhas individuais sao registradas e NAO interrompem a execucao.
    - --limite N: processa no maximo N contatos (para teste controlado).
    - Grava log CSV em reports/ ao final.
"""
import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

_SRC = next(p for p in Path(__file__).resolve().parents if (p / "core").is_dir())
sys.path.insert(0, str(_SRC))

from core.auth_tiny import validate_tenants
from core.tiny_client import TinyClient
from core.produtos_utils import desempacotar_data

_ROOT = _SRC.parent
REPORTS_DIR = _ROOT / "reports"
DATA_DIR = _ROOT / "data"
from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS = get_all_tenant_ids()  # chaves de data/tenants.json


def _oculto(contato: dict) -> bool:
    """True se o contato tem o nome mascarado por asteriscos."""
    return "*" in str(contato.get("nome") or "")


def _ck_file(tenant: str) -> Path:
    return DATA_DIR / f"excluir_ocultos_ck_{tenant}.json"


def _ler_ck(tenant: str) -> set:
    f = _ck_file(tenant)
    if f.exists():
        try:
            return set(json.loads(f.read_text(encoding="utf-8")))
        except Exception:
            pass
    return set()


def _salvar_ck(tenant: str, ids: set) -> None:
    _ck_file(tenant).write_text(json.dumps(sorted(ids)), encoding="utf-8")


def coletar(client: TinyClient, situacao_filtro: str | None) -> list[dict]:
    """Varre a base e retorna os contatos ocultos (opcionalmente filtrando situacao)."""
    ocultos = []
    vistos = 0
    total_ref = {"n": None}

    def _on_total(t):
        total_ref["n"] = t
        print(f"    base total: {t} contatos")

    for it in client.contatos_todos(on_total=_on_total):
        vistos += 1
        if _oculto(it):
            if situacao_filtro and str(it.get("situacao", "")).upper() != situacao_filtro:
                continue
            ocultos.append({
                "id": it.get("id"),
                "nome": it.get("nome"),
                "cpfCnpj": it.get("cpfCnpj"),
                "situacao": it.get("situacao"),
                "status": "pendente",
            })
        if vistos % 5000 == 0:
            print(f"    ... {vistos} varridos, {len(ocultos)} ocultos ate agora", flush=True)

    print(f"    varredura concluida: {vistos} contatos, {len(ocultos)} ocultos")
    return ocultos


def salvar_log(linhas: list[dict], tenant: str, sufixo: str) -> Path:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = REPORTS_DIR / f"ocultos_{sufixo}_{tenant}_{stamp}.csv"
    with open(arquivo, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["id", "nome", "cpfCnpj", "situacao", "status"])
        w.writeheader()
        for l in linhas:
            w.writerow({k: l.get(k, "") for k in w.fieldnames})
    return arquivo


def resumo_situacao(ocultos: list[dict]) -> str:
    from collections import Counter
    c = Counter(str(o.get("situacao", "")).upper() or "?" for o in ocultos)
    return ", ".join(f"{k}={v}" for k, v in sorted(c.items()))


def processar_tenant(tenant: str, executar: bool, situacao_filtro: str | None, limite: int | None) -> dict:
    print(f"\n{'='*70}\n  TENANT: {tenant}\n{'='*70}")
    client = TinyClient(tenant)

    print("  Varrendo base de contatos ao vivo...")
    ocultos = coletar(client, situacao_filtro)

    if not ocultos:
        print("  Nenhum contato oculto encontrado.")
        return {"tenant": tenant, "total": 0, "excluidos": 0, "erro": 0}

    print(f"\n  {len(ocultos)} contato(s) oculto(s) encontrado(s).  situacao: {resumo_situacao(ocultos)}")
    print("  Amostra:")
    for o in ocultos[:10]:
        print(f"    id={o['id']}  nome={o['nome']!r}  situacao={o['situacao']}")

    if not executar:
        log = salvar_log(ocultos, tenant, "dryrun")
        print(f"\n  [DRY-RUN] Nada foi excluido. Lista salva em: {log}")
        print(f"  Para excluir de verdade: python excluir_ocultos.py {tenant} --executar")
        return {"tenant": tenant, "total": len(ocultos), "excluidos": 0, "erro": 0}

    # ---- execucao real ----
    ja_feitos = _ler_ck(tenant)
    if ja_feitos:
        print(f"\n  Checkpoint: {len(ja_feitos)} ja processados em execucao anterior — serao pulados.")

    alvos = ocultos if limite is None else ocultos[:limite]
    total = len(alvos)
    if limite is not None:
        print(f"\n  [LIMITE] processando apenas {total} contato(s) para teste.")

    exc = err = 0
    for i, o in enumerate(alvos, 1):
        cid = o["id"]
        if cid in ja_feitos:
            o["status"] = "ja_processado"
            continue
        # tipos vem do detalhe (a listagem nao traz); evita tipo 'Outro' indevido
        tipos_ids = None
        try:
            det = desempacotar_data(client.obter_contato(cid))
            tipos_ids = [t["id"] for t in (det.get("tipos") or []) if isinstance(t, dict) and t.get("id")]
        except Exception:
            pass
        print(f"  [{i}/{total}] excluindo id={cid} {str(o['nome'])[:24]!r} ... ", end="", flush=True)
        try:
            client.excluir_contato(cid, nome=str(o["nome"]), tipos_ids=tipos_ids)
            # verifica: relê e confirma situacao='E'
            apos = desempacotar_data(client.obter_contato(cid)).get("situacao")
            if str(apos).upper() == "E":
                o["status"] = "excluido"
                exc += 1
                print("OK (situacao=E)")
            else:
                o["status"] = f"nao_confirmado: situacao apos={apos!r}"
                err += 1
                print(f"FALHA (situacao apos={apos!r})")
        except Exception as e:
            o["status"] = f"erro: {str(e)[:120]}"
            err += 1
            print(f"ERRO ({str(e)[:60]})")
        ja_feitos.add(cid)
        if i % 50 == 0:
            _salvar_ck(tenant, ja_feitos)

    _salvar_ck(tenant, ja_feitos)
    log = salvar_log(ocultos, tenant, "exclusao")
    print(f"\n  Concluido: {exc} excluido(s), {err} erro(s)/nao confirmado(s).")
    print(f"  Log salvo em: {log}")
    return {"tenant": tenant, "total": total, "excluidos": exc, "erro": err}


def main():
    ap = argparse.ArgumentParser(description="Exclui contatos ocultos (LGPD) no Tiny.")
    ap.add_argument("tenant", nargs="?", choices=TENANTS, help="tenant especifico (default: todos)")
    ap.add_argument("--executar", action="store_true", help="executa a exclusao (sem isto = dry-run)")
    ap.add_argument("--situacao", choices=["A", "B", "I"], help="filtra por situacao antes de excluir; default=todos")
    ap.add_argument("--limite", type=int, help="processa no maximo N contatos (teste controlado)")
    args = ap.parse_args()

    tenants = [args.tenant] if args.tenant else TENANTS

    if not validate_tenants(tenants, interactive=False):
        print("\n  Tokens invalidos. Reautentique antes:")
        for t in tenants:
            print(f"    python src/apps/auth/_reauth.py {t}")
        sys.exit(1)

    if args.executar:
        print("\n  *** MODO EXECUCAO — EXCLUSAO IRREVERSIVEL ***")
    else:
        print("\n  Modo DRY-RUN (nada sera excluido). Use --executar para aplicar.")

    resultados = [processar_tenant(t, args.executar, args.situacao, args.limite) for t in tenants]

    print(f"\n{'='*70}\n  RESUMO GERAL\n{'='*70}")
    for r in resultados:
        print(f"  {r['tenant']:10} ocultos={r['total']:6}  excluidos={r['excluidos']:6}  "
              f"erro={r['erro']:6}")


if __name__ == "__main__":
    main()
