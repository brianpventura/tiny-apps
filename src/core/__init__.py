"""Modulo nucleo: autenticacao OAuth + cliente da API V3 do Tiny.

Uso tipico em qualquer aplicacao:

    from core import TinyClient, validate_tenants

    if not validate_tenants(["empresa_a"]):
        sys.exit(1)
    client = TinyClient("empresa_a")
    produtos = client.listar_produtos()

Os tenant_ids validos sao as chaves de data/tenants.json.
"""

from .auth_tiny import (TinyTokenManager, TokenExpiredError, TENANT_NAMES,
                        validate_tenants, capture_authorization_code,
                        get_all_tenant_ids)
from .tiny_client import TinyClient
from .rate_limiter import get_limiter
from .produtos_utils import desempacotar_data, sanitizar_payload_produto, preparar_payload_put
from .categorias_utils import achatar_categorias, ResolvedorCategoria, norm_cat
from .contatos_utils import melhor_telefone, buscar_telefone_por_cpf

__all__ = [
    "TinyClient",
    "TinyTokenManager",
    "TokenExpiredError",
    "TENANT_NAMES",
    "get_all_tenant_ids",
    "validate_tenants",
    "capture_authorization_code",
    "get_limiter",
    "desempacotar_data",
    "sanitizar_payload_produto",
    "preparar_payload_put",
    "achatar_categorias",
    "ResolvedorCategoria",
    "norm_cat",
    "melhor_telefone",
    "buscar_telefone_por_cpf",
]
