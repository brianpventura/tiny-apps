"""Reautenticacao nao-interativa (sem input) de um tenant do Tiny V3.

Abre o navegador na URL de autorizacao e captura o `code` via callback local,
trocando por tokens. Uso:

    python _reauth.py <tenant_id>      # chave de data/tenants.json
"""
import sys
from pathlib import Path

for _p in Path(__file__).resolve().parents:
    if (_p / "core").is_dir():
        sys.path.insert(0, str(_p))
        break

from core.auth_tiny import (TinyTokenManager, capture_authorization_code,
                            TENANT_NAMES, resolver_tenant)


def main():
    tenant = resolver_tenant(sys.argv[1] if len(sys.argv) > 1 else None)
    nome = TENANT_NAMES.get(tenant, tenant)
    tm = TinyTokenManager(tenant)

    healthy, msg = tm.is_healthy()
    if healthy:
        print(f"  [OK] {nome} ja autenticado — {msg}")
        return

    print(f"  === RE-AUTENTICACAO: {nome} ({tenant}) ===")
    print(f"  Faca login no Tiny com a conta da {nome} na janela do navegador.")
    code = capture_authorization_code(tm.get_auth_url())
    tm.bootstrap(code)
    healthy, msg = tm.is_healthy()
    print(f"  [{'OK' if healthy else 'FALHOU'}] {nome}: {msg}")


if __name__ == "__main__":
    main()
