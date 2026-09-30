"""Teste offline do pipeline -- sem MySQL, sem API, sem token.

O repo nao tem suite de testes, e a verificacao de verdade e a fase 5 contra o
banco carregado. Mas tres classes de erro daqui nao precisam de banco nenhum
para aparecer, e sao justamente as que so apareceriam depois de 5h de coleta:

  1. lista de colunas do etl_load fora de sincronia com o DDL (typo numa
     coluna = a carga inteira falha no primeiro lote);
  2. escrita em coluna GENERATED (o MySQL recusa);
  3. transform quebrando em campo ausente, data vazia ou bloco nao preenchido.

Uso:
    .venv\\Scripts\\python.exe src\\apps\\pipeline_etl\\_teste_offline.py
    .venv\\Scripts\\python.exe src\\apps\\pipeline_etl\\_teste_offline.py --amostra
        (usa data/amostra_pedido_<tenant>.json, se a sonda de schema ja rodou)

Sai 0 se tudo passou, 1 caso contrario.
"""

import argparse
import io
import json
import re
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                              line_buffering=True)

import _bootstrap  # noqa: E402
import etl_load as L  # noqa: E402
import etl_transform as T  # noqa: E402

FALHAS = []


def ok(msg):
    print("  [ok]    {}".format(msg))


def falha(msg):
    FALHAS.append(msg)
    print("  [FALHA] {}".format(msg))


# ------------------------------------------------------------------ #
#  1. Colunas do load x DDL                                           #
# ------------------------------------------------------------------ #

def colunas_do_ddl() -> dict:
    ddl = (_bootstrap.SQL_DIR / "001_schema.sql").read_text(encoding="utf-8")
    tabelas = {}
    for m in re.finditer(r"CREATE TABLE IF NOT EXISTS (\w+) \((.*?)\n\) ENGINE",
                         ddl, re.S):
        nome, corpo = m.group(1), m.group(2)
        cols, geradas = [], []
        for linha in corpo.split("\n"):
            s = linha.strip()
            if not s or s.startswith(("--", "PRIMARY KEY", "KEY ", "UNIQUE",
                                      "CONSTRAINT", ")")):
                continue
            mm = re.match(r"^(\w+)\s+[A-Z]", s)
            if mm:
                cols.append(mm.group(1))
                if "GENERATED ALWAYS AS" in s:
                    geradas.append(mm.group(1))
        tabelas[nome] = {"colunas": cols, "geradas": geradas}
    return tabelas


def testar_colunas(tabelas: dict) -> None:
    print("\n1. Colunas do etl_load contra o DDL")
    alvos = [("pedido", L.COLS_PEDIDO, "COLS_PEDIDO"),
             ("pedido_item", L.COLS_ITEM, "COLS_ITEM"),
             ("pedido_endereco", L.COLS_ENDERECO, "COLS_ENDERECO"),
             ("pedido_parcela", L.COLS_PARCELA, "COLS_PARCELA"),
             ("pedido_pagamento_integrado", L.COLS_PAGAMENTO, "COLS_PAGAMENTO"),
             ("pedido_marcador", L.COLS_MARCADOR, "COLS_MARCADOR")]
    for chave, (tab, cols, chaves, _n) in L._DIMS.items():
        alvos.append((tab, cols, "dim:" + chave))
        faltando = [c for c in chaves if c not in tabelas.get(tab, {}).get("colunas", [])]
        if faltando:
            falha("chaves da dim {} ausentes no DDL: {}".format(tab, faltando))

    for tabela, cols, rotulo in alvos:
        existentes = tabelas.get(tabela, {}).get("colunas", [])
        if not existentes:
            falha("tabela {} nao existe no DDL".format(tabela))
            continue
        faltando = [c for c in cols if c not in existentes]
        if faltando:
            falha("{}: colunas ausentes em {} -> {}".format(rotulo, tabela, faltando))
            continue
        geradas = [c for c in cols if c in tabelas[tabela]["geradas"]]
        if geradas:
            falha("{}: escreve em coluna GENERATED de {} -> {}".format(
                rotulo, tabela, geradas))
            continue
        ok("{:<26} {} colunas conferem com {}".format(rotulo, len(cols), tabela))


def testar_placeholders() -> None:
    print("\n2. Placeholders de cada INSERT")
    for nome in sorted(n for n in dir(L) if n.startswith("SQL_")):
        sql = getattr(L, nome)
        if not isinstance(sql, str) or "VALUES" not in sql:
            continue
        vals = re.search(r"VALUES \(([^)]*)\)", sql)
        cols_m = re.search(r"INSERT (?:IGNORE )?INTO \w+ \(([^)]*)\)", sql)
        if not vals or not cols_m:
            continue
        # conta SLOTS separados por virgula: ha literais ocupando posicao.
        n_val = len(vals.group(1).split(","))
        n_col = len(cols_m.group(1).split(","))
        if n_val != n_col:
            falha("{}: {} valores para {} colunas".format(nome, n_val, n_col))
        else:
            ok("{:<24} {} valores = {} colunas".format(nome, n_val, n_col))


# ------------------------------------------------------------------ #
#  3. Transform                                                       #
# ------------------------------------------------------------------ #

PEDIDO_MINIMO = {"id": 1, "situacao": 6, "dataCriacao": "2026-09-03",
                 "itens": [{"produto": {"id": 9, "sku": "X1",
                                        "descricao": "P", "tipo": "P"},
                            "quantidade": 2, "valorUnitario": "10.50"}]}

PEDIDO_SUJO = {"id": 2, "situacao": None, "data": "", "dataEntrega": "",
               "dataEnvio": None, "valorTotalProdutos": None,
               "cliente": None, "ecommerce": [], "naturezaOperacao": "",
               "itens": [{"produto": {}, "quantidade": 0, "valorUnitario": None},
                         {"quantidade": "1,5", "valorUnitario": "0"}],
               "parcelas": None, "marcadores": ["urgente", {"descricao": "nf"}]}


def testar_transform() -> None:
    print("\n3. Transform (pedido minimo, pedido sujo, so listagem)")
    ctx = T.contexto("empresa_a")

    r = T.transformar(ctx, 1, {}, PEDIDO_MINIMO)
    p = r["pedido"]
    if p["qtd_itens"] != 1 or str(p["total_unidades"]) != "2":
        falha("pedido minimo: qtd_itens={} total_unidades={}".format(
            p["qtd_itens"], p["total_unidades"]))
    elif str(r["itens"][0]["valor_unitario"]) != "10.50":
        falha("pedido minimo: valor_unitario virou {!r}".format(
            r["itens"][0]["valor_unitario"]))
    elif p["tem_detalhe"] != 1 or p["eh_situacao_valida"] != 1:
        falha("pedido minimo: flags erradas")
    else:
        ok("pedido minimo: 1 item, 2 un, R$ 10,50 em Decimal")

    try:
        r2 = T.transformar(ctx, 2, {}, PEDIDO_SUJO)
    except Exception as e:
        falha("pedido sujo estourou: {!r}".format(e))
    else:
        p2 = r2["pedido"]
        problemas = []
        # '' em coluna DATE erra com sql_mode estrito -- tem de virar None.
        for campo in ("data_pedido", "data_entrega", "data_envio"):
            if p2[campo] is not None:
                problemas.append("{}={!r}".format(campo, p2[campo]))
        if p2["situacao"] is not None:
            problemas.append("situacao nao virou None")
        if len(r2["itens"]) != 2:
            problemas.append("perdeu item (qtd<=0 deve ser MANTIDO)")
        if str(r2["itens"][1]["quantidade"]) != "1.5":
            problemas.append("virgula decimal: {!r}".format(
                r2["itens"][1]["quantidade"]))
        if len(r2["marcadores"]) != 2:
            problemas.append("marcadores str+dict: {}".format(r2["marcadores"]))
        if problemas:
            falha("pedido sujo: " + "; ".join(problemas))
        else:
            ok("pedido sujo: datas vazias -> None, item qtd=0 mantido, "
               "virgula decimal, marcador str e dict")

    r3 = T.transformar(ctx, 3, PEDIDO_MINIMO, None)
    if r3["pedido"]["tem_detalhe"] != 0:
        falha("so listagem: tem_detalhe deveria ser 0")
    else:
        ok("so listagem: tem_detalhe=0 (nao entra em v_venda_item)")

    h1 = T.hash_payload(PEDIDO_MINIMO)
    h2 = T.hash_payload(json.loads(json.dumps(PEDIDO_MINIMO)))
    if h1 != h2 or len(h1) != 64:
        falha("hash nao e estavel/canonico")
    else:
        ok("hash canonico estavel (SHA-256, 64 chars)")


def testar_amostra() -> None:
    """Roda o transform sobre o JSON real gravado pela sonda de schema."""
    print("\n4. Transform sobre a amostra real de data/")
    achou = False
    for arq in sorted(_bootstrap.DATA_DIR.glob("amostra_pedido_*.json")):
        achou = True
        tenant = arq.stem.replace("amostra_pedido_", "")
        ctx = T.contexto(tenant)
        try:
            amostras = json.loads(arq.read_text(encoding="utf-8"))
        except Exception as e:
            falha("{}: ilegivel ({})".format(arq.name, e))
            continue
        for a in amostras:
            try:
                r = T.transformar(ctx, a["id"], a.get("listagem") or {},
                                  a.get("detalhe"))
            except Exception as e:
                falha("{} pedido {}: {!r}".format(arq.name, a.get("id"), e))
                continue
            p = r["pedido"]
            soma = sum(i["quantidade"] * i["valor_unitario"] for i in r["itens"])
            delta = abs(soma - p["valor_total_produtos"])
            marca = "ok" if delta <= 1 else "DELTA"
            print("     {} pedido {}: {} itens, R$ {} (soma {}, dif {}) [{}]"
                  .format(tenant, a["id"], p["qtd_itens"],
                          p["valor_total_produtos"], soma, delta, marca))
    if not achou:
        print("     (nenhuma amostra ainda -- rode --sondar-schema primeiro)")


def main() -> int:
    ap = argparse.ArgumentParser(description="Teste offline do pipeline ETL.")
    ap.add_argument("--amostra", action="store_true",
                    help="tambem roda sobre data/amostra_pedido_*.json")
    args = ap.parse_args()

    print("\n  Teste offline do pipeline_etl (sem MySQL, sem API, sem token)")
    tabelas = colunas_do_ddl()
    print("  {} tabelas lidas do DDL.".format(len(tabelas)))
    testar_colunas(tabelas)
    testar_placeholders()
    testar_transform()
    if args.amostra:
        testar_amostra()

    print("\n  RESULTADO: {}".format(
        "TUDO OK" if not FALHAS else "{} FALHA(S)".format(len(FALHAS))))
    return 1 if FALHAS else 0


if __name__ == "__main__":
    sys.exit(main())
