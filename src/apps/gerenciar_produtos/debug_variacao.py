"""Diagnostico: imprime o JSON bruto retornado pela busca por SKU.

Uso:
    python debug_variacao.py <sku> [tenant_id]
"""

import json
import sys
from pathlib import Path

# adiciona a pasta src/ ao sys.path (sobe ate achar a pasta que contem 'core')
for _p in Path(__file__).resolve().parents:
    if (_p / 'core').is_dir():
        sys.path.insert(0, str(_p))
        break

from core.auth_tiny import resolver_tenant

from core.auth_tiny import validate_tenants
from core.tiny_client import TinyClient


def main():
    if len(sys.argv) < 2:
        print("Uso: python debug_variacao.py <sku> [tenant_id]")
        sys.exit(1)

    sku = sys.argv[1]
    tenant_id = resolver_tenant(sys.argv[2] if len(sys.argv) > 2 else None)

    if not validate_tenants([tenant_id]):
        sys.exit(1)

    client = TinyClient(tenant_id)

    print(f"\n=== GET /produtos?codigo={sku} ===")
    resp = client._get("produtos", params={"codigo": sku, "limit": 10})
    if "data" in resp:
        resp = resp["data"]
    itens = resp.get("itens", [])
    print(f"({len(itens)} item(ns))\n")
    for it in itens:
        print(json.dumps(it, indent=2, ensure_ascii=False))
        print()

    if not itens:
        return

    pid = itens[0].get("id")
    print(f"\n=== GET /produtos/{pid} ===")
    try:
        full = client.obter_produto(pid)
        print(json.dumps(full, indent=2, ensure_ascii=False))
    except Exception as e:
        print(f"FALHOU: {e}")


if __name__ == "__main__":
    main()
