"""Utilitarios de contato (telefone) compartilhados entre as aplicacoes.

O detalhe do pedido so traz 'fone'; quando vazio, o CONTATO no Tiny pode ter
'telefone' e/ou 'celular'. Estas funcoes resolvem o melhor telefone a partir do
cadastro de contatos (busca por CPF/CNPJ).
"""

from .produtos_utils import desempacotar_data


def melhor_telefone(contato: dict) -> str:
    """Primeiro nao-vazio entre telefone -> celular -> telefoneAdicional -> fone."""
    if not isinstance(contato, dict):
        return ""
    for campo in ("telefone", "celular", "telefoneAdicional", "fone"):
        v = contato.get(campo)
        if v and str(v).strip():
            return str(v).strip()
    return ""


def buscar_telefone_por_cpf(clients: dict, cpf: str) -> str:
    """Procura o contato por CPF/CNPJ em cada tenant e retorna o melhor telefone.

    clients: {tenant_id: TinyClient}. Retorna "" se nao achar em nenhum.
    """
    cpf = (cpf or "").strip()
    if not cpf:
        return ""
    for client in clients.values():
        try:
            resp = desempacotar_data(client.buscar_contatos(cpfCnpj=cpf, limit=10))
        except Exception:
            continue
        itens = resp.get("itens", []) if isinstance(resp, dict) else (resp or [])
        if not itens:
            continue
        tel = melhor_telefone(itens[0])
        if not tel and itens[0].get("id"):
            try:
                tel = melhor_telefone(desempacotar_data(client.obter_contato(itens[0]["id"])))
            except Exception:
                tel = ""
        if tel:
            return tel
    return ""
