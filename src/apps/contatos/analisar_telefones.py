"""
Analisa a base de contatos do Tiny e identifica quais estao sujeitos a
normalizacao do campo Telefone. NAO altera nada (somente leitura).

Regras (padrao brasileiro):
  - Numero movel  = DDD (2) + 9 digitos, sendo o 1o apos o DDD = 9  -> 11 digitos
  - Numero fixo   = DDD (2) + 8 digitos, sendo o 1o apos o DDD em 2-5 -> 10 digitos
  - Movel "sem o 9" = 10 digitos com 1o apos o DDD em {6,7,8,9} -> precisa inserir 9

Decisao por contato:
  1. Sem telefone E sem celular  -> IGNORA (bucket 'sem_contato')
  2. Telefone vazio e celular preenchido -> copia celular p/ telefone (bucket 'copiar_celular')
  3. Telefone e' movel sem o 9    -> inserir 9 (bucket 'inserir_9')
  4. Telefone ja' valido          -> IGNORA (bucket 'conforme')
  5. Formato inesperado           -> bucket 'anomalo' (revisao manual)

O numero considerado como fonte do Telefone final e':
  telefone, se preenchido; senao celular.
E' entao normalizado (insere 9 se for movel sem o 9).
'Sujeito a mudanca' = copiar_celular OR inserir_9 (ou ambos).
"""

import re

# DDDs validos no Brasil (todos os codigos de area realmente atribuidos)
DDDS_VALIDOS = {
    11,12,13,14,15,16,17,18,19, 21,22,24,27,28, 31,32,33,34,35,37,38,
    41,42,43,44,45,46,47,48,49, 51,53,54,55, 61,62,63,64,65,66,67,68,69,
    71,73,74,75,77,79, 81,82,83,84,85,86,87,88,89,
    91,92,93,94,95,96,97,98,99,
}


def so_digitos(valor) -> str:
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def _nacional(digitos: str) -> str:
    """Remove o codigo de pais 55, quando presente, devolvendo DDD+local."""
    d = digitos
    # 55 + 11 (movel) ou 55 + 10 (fixo)
    if len(d) in (12, 13) and d.startswith("55"):
        d = d[2:]
    return d


def classificar_numero(valor) -> dict:
    """Classifica um unico numero. Retorna dict com:
       tipo: 'movel' | 'movel_sem9' | 'fixo' | 'vazio' | 'anomalo'
       nacional: digitos DDD+local (normalizado de pais)
       normalizado: digitos ja com o 9 inserido quando aplicavel
       ddd, local
    """
    d = _nacional(so_digitos(valor))
    if not d:
        return {"tipo": "vazio", "nacional": "", "normalizado": "", "ddd": "", "local": ""}

    # Precisa de pelo menos DDD(2) + local. Consideramos DDD valido.
    if len(d) < 10:
        return {"tipo": "anomalo", "nacional": d, "normalizado": d, "ddd": "", "local": ""}

    ddd = d[:2]
    local = d[2:]

    if int(ddd) not in DDDS_VALIDOS:
        return {"tipo": "anomalo", "nacional": d, "normalizado": d, "ddd": ddd, "local": local}

    if len(local) == 9:
        if local[0] == "9":
            return {"tipo": "movel", "nacional": d, "normalizado": d, "ddd": ddd, "local": local}
        # 9 digitos mas nao comeca com 9 -> inesperado
        return {"tipo": "anomalo", "nacional": d, "normalizado": d, "ddd": ddd, "local": local}

    if len(local) == 8:
        p = local[0]
        if p in "2345":
            return {"tipo": "fixo", "nacional": d, "normalizado": d, "ddd": ddd, "local": local}
        if p in "6789":
            # movel sem o 9 -> insere 9
            norm = ddd + "9" + local
            return {"tipo": "movel_sem9", "nacional": d, "normalizado": norm, "ddd": ddd, "local": local}
        return {"tipo": "anomalo", "nacional": d, "normalizado": d, "ddd": ddd, "local": local}

    # local com 10+ digitos ou outros -> anomalo
    return {"tipo": "anomalo", "nacional": d, "normalizado": d, "ddd": ddd, "local": local}


def analisar_contato(telefone, celular) -> dict:
    """Decide o bucket e o telefone proposto para um contato.

    Retorna dict:
      bucket: 'sem_contato' | 'conforme' | 'copiar_celular' | 'inserir_9' | 'anomalo'
      sujeito: bool  (True se precisa de escrita)
      telefone_proposto: digitos propostos p/ o campo Telefone (ou "")
      motivos: lista de strings
    """
    ct = classificar_numero(telefone)
    cc = classificar_numero(celular)
    tel_vazio = ct["tipo"] == "vazio"
    cel_vazio = cc["tipo"] == "vazio"

    if tel_vazio and cel_vazio:
        return {"bucket": "sem_contato", "sujeito": False,
                "telefone_proposto": "", "motivos": ["sem telefone e sem celular"]}

    motivos = []
    copiou = False

    # Fonte do Telefone final
    if tel_vazio and not cel_vazio:
        fonte = cc
        copiou = True
        motivos.append("telefone vazio -> copiar do celular")
    else:
        fonte = ct

    if fonte["tipo"] == "anomalo":
        return {"bucket": "anomalo", "sujeito": False,
                "telefone_proposto": fonte["normalizado"],
                "motivos": motivos + [f"formato inesperado ({fonte['nacional'] or 'vazio'})"]}

    proposto = fonte["normalizado"]

    if fonte["tipo"] == "movel_sem9":
        motivos.append("movel sem o 9 -> inserir 9")

    # Determina se ha mudanca real no campo Telefone (comparando digitos)
    tel_atual_dig = _nacional(so_digitos(telefone))
    mudou = (proposto != tel_atual_dig)

    if not mudou:
        return {"bucket": "conforme", "sujeito": False,
                "telefone_proposto": proposto, "motivos": ["telefone ja conforme"]}

    if copiou and fonte["tipo"] == "movel_sem9":
        bucket = "copiar_celular"  # copia E insere 9; prioriza a copia como motivo primario
    elif copiou:
        bucket = "copiar_celular"
    else:
        bucket = "inserir_9"

    return {"bucket": bucket, "sujeito": True,
            "telefone_proposto": proposto, "motivos": motivos}
