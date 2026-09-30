"""Utilitarios compartilhados para manipular o payload de produtos do Tiny V3.

Centraliza quirks conhecidos da API para que todas as aplicacoes (renomear,
categorizar, inativar/excluir, ...) tratem o payload de forma consistente.
"""


def desempacotar_data(resp: dict) -> dict:
    """Tiny costuma envelopar payloads em {'data': {...}}. Extrai se houver."""
    if isinstance(resp, dict) and "data" in resp and isinstance(resp["data"], dict):
        return resp["data"]
    return resp


def sanitizar_payload_produto(produto: dict) -> dict:
    """Remove do JSON do GET campos placeholder que o Tiny rejeita no PUT.

    ATENCAO: NAO use para PUT de produto existente. O GET nao devolve
    `fornecedores[].padrao`, e o filtro abaixo exige esse campo -- entao
    fornecedores REAIS tambem sao descartados e o PUT apaga o vinculo do
    produto com o fornecedor. Para PUT use `preparar_payload_put`.

    Caso conhecido: 'fornecedores' vem com [{id:0, padrao:null}] quando nao ha
    fornecedor cadastrado — o PUT exige id>0 e padrao nao-nulo. Filtramos esses
    placeholders, e se a lista ficar vazia removemos o campo.

    Retorna uma copia rasa; nao muta o dict original.
    """
    payload = dict(produto)
    forn = payload.get("fornecedores")
    if isinstance(forn, list):
        validos = [f for f in forn if isinstance(f, dict)
                   and (f.get("id") or 0) > 0
                   and f.get("padrao") is not None]
        if validos:
            payload["fornecedores"] = validos
        else:
            payload.pop("fornecedores", None)
    return payload


def corrigir_fornecedores(alvo: dict) -> None:
    """Deixa o array 'fornecedores' de `alvo` aceitavel para o PUT, PRESERVANDO os reais.

    Diferenca para sanitizar_payload_produto: aquele exige 'padrao' nao-nulo e por isso
    descarta fornecedores reais (o GET nao devolve esse campo). Aqui o 'padrao' e
    reinjetado — exatamente UM fornecedor fica com padrao=True e os demais com False —
    de modo que o vinculo do produto sobrevive ao PUT. Use este quando o objetivo e
    alterar um campo sem mexer no fornecedor.

    Muta `alvo` no lugar. Placeholders ([{id:0, padrao:null}]) continuam sendo removidos.
    """
    forn = alvo.get("fornecedores")
    if not isinstance(forn, list):
        return
    validos = [f for f in forn if isinstance(f, dict) and (f.get("id") or 0) > 0]
    if not validos:
        alvo.pop("fornecedores", None)
        return
    idx_padrao = next((i for i, f in enumerate(validos) if f.get("padrao") is True), 0)
    for i, f in enumerate(validos):
        f["padrao"] = (i == idx_padrao)
    alvo["fornecedores"] = validos


def preparar_payload_put(produto_full: dict) -> dict:
    """Copia o JSON do GET e o torna PUT-safe SEM alterar dado de negocio.

    A unica transformacao e corrigir 'fornecedores' (no produto e em cada variacao).
    O chamador altera depois so o campo que quer gravar.
    """
    import copy

    payload = copy.deepcopy(produto_full)
    corrigir_fornecedores(payload)
    for var in payload.get("variacoes", []) or []:
        if isinstance(var, dict):
            corrigir_fornecedores(var)
    return payload
