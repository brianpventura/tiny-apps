"""Fachada das regras de negocio para o pipeline -- IMPORTADAS, nunca recopiadas.

A politica (situacoes validas, naturezas de venda, canais) vem do perfil de
`core/regras_vendas.py` (config/regras.padrao.json + regras.json da empresa).
Este modulo acrescenta o que e proprio do pipeline: a VERSAO das regras gravada
em `pedido.regras_versao` e as janelas de paginacao.

VERSAO DAS REGRAS
  `cfg_regras` guarda, por tenant, o hash do perfil efetivo e um numero de
  versao. Na fase 4:
    * sem registro (banco novo): registra a versao 1 e segue;
    * hash igual: segue com a versao gravada;
    * hash diferente: a regra mudou. Sem `aplicar=True` (flag
      --aplicar-regras) o pipeline PARA e mostra o que mudou -- trocar de regra
      por acidente (ex.: TINY_APPS_CONFIG nao definida, caindo no padrao
      generico) reprocessaria o warehouse inteiro com a regra errada. Com a
      flag, a versao sobe e todo pedido com `regras_versao` menor e
      renormalizado offline, sem request.
  REGRAS_CODIGO entra no hash: subir essa constante forca o reprocessamento
  quando o CODIGO que deriva as colunas mudar (nao a configuracao).
"""

import json
from datetime import date, datetime

import _bootstrap  # noqa: F401  (efeito colateral: sys.path)

from core import regras_vendas as _rv  # noqa: E402
from core.auth_tiny import get_all_tenant_ids  # noqa: E402

REGRAS_CODIGO = 1

SITUACOES = _rv.SITUACOES
perfil = _rv.perfil
origem_perfil = _rv.origem

norm = _rv.norm
clientes_excluidos = _rv.clientes_excluidos
eh_excluido = _rv.eh_excluido
extrair_produto = _rv.extrair_produto
para_float = _rv.para_float
data_de = _rv.data_de
parse_data = _rv.parse_data

NOMES = _rv.NOMES
TENANTS_PADRAO = get_all_tenant_ids()


class RegrasMudaram(Exception):
    """O perfil efetivo mudou desde a ultima carga e nao houve --aplicar-regras."""


def _hash(tenant: str) -> str:
    return f"{REGRAS_CODIGO}:{perfil(tenant).hash}"


def _diff(antigo: dict, novo: dict, prefixo: str = "") -> list[str]:
    linhas = []
    for k in sorted(set(antigo) | set(novo)):
        a, b = antigo.get(k), novo.get(k)
        if isinstance(a, dict) and isinstance(b, dict):
            linhas += _diff(a, b, f"{prefixo}{k}.")
        elif a != b:
            linhas.append(f"{prefixo}{k}: {json.dumps(a, ensure_ascii=False)[:120]} -> "
                          f"{json.dumps(b, ensure_ascii=False)[:120]}")
    return linhas


def versao_regras(conn, tenant: str, aplicar: bool = False) -> int:
    """Versao das regras a gravar em pedido.regras_versao (ver docstring)."""
    import etl_db

    p = perfil(tenant)
    h = _hash(tenant)
    atual = {"vendas": p.dados.get("vendas"), "canais": p.dados.get("canais")}
    r = etl_db.consultar(conn, "SELECT versao, hash, perfil FROM cfg_regras WHERE tenant = %s",
                         (tenant,))
    if not r:
        etl_db.executar(conn, "INSERT INTO cfg_regras (tenant, versao, hash, perfil, atualizado_em) "
                              "VALUES (%s, 1, %s, %s, %s)",
                        (tenant, h, json.dumps(atual, ensure_ascii=False), datetime.now()))
        conn.commit()
        return 1
    versao, h_antigo = int(r[0]["versao"]), r[0]["hash"]
    if h_antigo == h:
        return versao
    anterior = json.loads(r[0]["perfil"]) if r[0]["perfil"] else {}
    mudancas = _diff(anterior, atual) or ["(codigo das regras: REGRAS_CODIGO mudou)"]
    if not aplicar:
        raise RegrasMudaram(
            f"as regras de '{tenant}' mudaram desde a ultima carga "
            f"(perfil: {origem_perfil()}):\n      "
            + "\n      ".join(mudancas[:15])
            + "\n    Se a mudanca e intencional, rode de novo com --aplicar-regras "
              "(reprocessa offline, sem request).\n    Se nao e, confira a variavel "
              "TINY_APPS_CONFIG e o regras.json da empresa.")
    nova = versao + 1
    etl_db.executar(conn, "UPDATE cfg_regras SET versao = %s, hash = %s, perfil = %s, "
                          "atualizado_em = %s WHERE tenant = %s",
                    (nova, h, json.dumps(atual, ensure_ascii=False), datetime.now(), tenant))
    conn.commit()
    print(f"  [{tenant}] regras mudaram -> versao {versao} -> {nova}:")
    for m in mudancas[:15]:
        print(f"      {m}")
    return nova


def janelas_mensais(inicio: date, fim: date) -> list[tuple[date, date]]:
    """Fatia [inicio, fim] em janelas de mes civil.

    Existe porque a API estoura em
    offset profundo (~57k devolve HTTP 400): paginar sempre dentro de janela
    mensal mantem o offset abaixo de ~6k.
    """
    janelas, ano, mes = [], inicio.year, inicio.month
    while (ano, mes) <= (fim.year, fim.month):
        ini = date(ano, mes, 1)
        prox = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
        ult = min(fim, date.fromordinal(prox.toordinal() - 1))
        janelas.append((max(ini, inicio), ult))
        mes += 1
        if mes > 12:
            mes, ano = 1, ano + 1
    return janelas


def metade(ini: date, fim: date) -> list[tuple[date, date]]:
    """Divide uma janela ao meio. Usada quando paginacao.total passa do teto."""
    if ini >= fim:
        return [(ini, fim)]
    meio = date.fromordinal((ini.toordinal() + fim.toordinal()) // 2)
    if meio >= fim:
        meio = date.fromordinal(fim.toordinal() - 1)
    return [(ini, meio), (date.fromordinal(meio.toordinal() + 1), fim)]
