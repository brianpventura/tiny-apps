"""
Corrige o campo Telefone dos contatos do Tiny (grava so digitos).

Alvo: contatos do bucket 'inserir_9' (movel sem o 9) e 'copiar_celular'.
Escreve o telefone proposto SO COM DIGITOS (ex: 11999990000).

Contrato do PUT /contatos/{id} (validado empiricamente):
  - 'nome' obrigatorio (nao-vazio)
  - 'tipos' deve ser lista de IDs inteiros; se omitido, o Tiny adiciona um
    tipo 'Outro' indevido -> reenviamos os tipos atuais.
  - demais campos preservados quando omitidos.

Uso:
    python corrigir_telefones.py            # dry-run: so gera a planilha
    python corrigir_telefones.py --commit   # aplica os PUTs

Fonte: data/contatos_raw_<tenant>.json (cache gerado pela varredura).
Antes de cada PUT o detalhe e' relido (idempotente: pula quem ja esta' conforme).
Loga o resultado em reports/corrigir_telefones_<ts>.csv
"""
import sys, io, json, csv, argparse
from pathlib import Path
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
for _p in Path(__file__).resolve().parents:
    if (_p / "core").is_dir():
        sys.path.insert(0, str(_p)); _SRC = _p; break
sys.path.insert(0, str(Path(__file__).resolve().parent))
from core.tiny_client import TinyClient
from core.produtos_utils import desempacotar_data
from analisar_telefones import analisar_contato, classificar_numero, so_digitos, _nacional

from core.auth_tiny import get_all_tenant_ids  # noqa: E402

TENANTS = get_all_tenant_ids()  # chaves de data/tenants.json
DATA = _SRC.parent / "data"
REPORTS = _SRC.parent / "reports"


def carregar_cache(tenant):
    f = DATA / f"contatos_raw_{tenant}.json"
    with open(f, encoding="utf-8") as fh:
        return json.load(fh)


def classificar_base():
    """Retorna (sujeitos, anomalos) como listas de dicts a partir do cache."""
    sujeitos, anomalos = [], []
    for tenant in TENANTS:
        for it in carregar_cache(tenant):
            tel, cel = it.get("telefone"), it.get("celular")
            res = analisar_contato(tel, cel)
            linha = {"tenant": tenant, "id": it.get("id"), "nome": it.get("nome"),
                     "cpfCnpj": it.get("cpfCnpj"), "telefone_atual": tel,
                     "celular_atual": cel, "telefone_proposto": res["telefone_proposto"],
                     "bucket": res["bucket"], "motivos": "; ".join(res["motivos"])}
            if res["sujeito"]:
                sujeitos.append(linha)
            elif res["bucket"] == "anomalo":
                anomalos.append(linha)
    return sujeitos, anomalos


def gerar_planilha(sujeitos, anomalos):
    import openpyxl
    wb = openpyxl.Workbook()
    cols = ["tenant","id","nome","cpfCnpj","telefone_atual","celular_atual",
            "telefone_proposto","bucket","motivos"]
    ws = wb.active
    ws.title = "sujeitos"
    ws.append(cols)
    for l in sujeitos:
        ws.append([l[c] for c in cols])
    wa = wb.create_sheet("anomalos_revisao_manual")
    wa.append(cols)
    for l in anomalos:
        wa.append([l[c] for c in cols])
    out = REPORTS / "analise_telefones_FULL.xlsx"
    wb.save(out)
    return out


def aplicar(sujeitos):
    clients = {t: TinyClient(t) for t in TENANTS}
    log_rows = []
    stats = Counter()
    total = len(sujeitos)
    for i, l in enumerate(sujeitos, 1):
        tenant, cid = l["tenant"], l["id"]
        c = clients[tenant]
        proposto = l["telefone_proposto"]
        status = ""
        try:
            det = desempacotar_data(c.obter_contato(cid))
            nome = det.get("nome") or ""
            tipos_ids = [t["id"] for t in (det.get("tipos") or []) if isinstance(t, dict) and t.get("id")]
            tel_live = det.get("telefone")
            cel_live = det.get("celular")
            # re-decide com base no estado atual (idempotente)
            res = analisar_contato(tel_live, cel_live)
            if not res["sujeito"]:
                status = "ja_conforme_pulado"
                stats[status] += 1
            else:
                alvo = res["telefone_proposto"]
                body = {"nome": nome, "telefone": alvo, "tipos": tipos_ids}
                c.atualizar_contato(cid, body)
                # verifica
                dep = desempacotar_data(c.obter_contato(cid))
                if so_digitos(dep.get("telefone")) == alvo:
                    status = "ok"
                    proposto = alvo
                    stats[status] += 1
                else:
                    status = f"falha_verificacao(gravou={dep.get('telefone')!r})"
                    stats["falha"] += 1
        except Exception as e:
            status = f"erro: {str(e)[:120]}"
            stats["erro"] += 1
        log_rows.append({**l, "telefone_proposto": proposto, "status": status})
        if i % 25 == 0 or i == total:
            print(f"  {i}/{total}  ok={stats['ok']} pulado={stats['ja_conforme_pulado']} "
                  f"falha={stats['falha']} erro={stats['erro']}", flush=True)
    # log csv
    out = REPORTS / "corrigir_telefones_resultado.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(log_rows[0].keys()))
        w.writeheader()
        w.writerows(log_rows)
    print(f"\n  Resumo: {dict(stats)}")
    print(f"  Log: {out}")
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true", help="aplica os PUTs (senao, so gera planilha)")
    args = ap.parse_args()

    sujeitos, anomalos = classificar_base()
    print(f"Sujeitos a mudanca: {len(sujeitos)} | Anomalos (revisao manual): {len(anomalos)}")
    by_tenant = Counter(l["tenant"] for l in sujeitos)
    print(f"  por tenant: {dict(by_tenant)}")

    out = gerar_planilha(sujeitos, anomalos)
    print(f"Planilha: {out}  (abas: sujeitos, anomalos_revisao_manual)")

    if not args.commit:
        print("\n[DRY-RUN] Nada gravado. Rode com --commit para aplicar.")
        return
    print(f"\n[COMMIT] Aplicando {len(sujeitos)} correcoes (GET+PUT+verify)...")
    aplicar(sujeitos)


if __name__ == "__main__":
    main()
