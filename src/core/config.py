"""Localizacao dos arquivos de configuracao de NEGOCIO.

Dois tipos de arquivo, em lugares diferentes de proposito:

  * segredos e estado (tenants.json, tokens_*.json, db.json, caches) ficam
    sempre em data/ -- nunca versionados;
  * configuracao de negocio (regras.json, linhas_produto.json,
    verificacao.json...) e o que muda de uma empresa para outra. Por padrao
    tambem fica em data/, mas a variavel de ambiente TINY_APPS_CONFIG pode
    apontar para outra pasta -- tipicamente a pasta config/ de um repositorio
    privado da empresa, onde esses arquivos sao versionados.

Os padroes genericos, que valem para qualquer empresa, ficam em config/ na raiz
deste repositorio (ex.: config/regras.padrao.json).
"""

import json
import os
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent.parent
DATA_DIR = RAIZ / "data"
PADROES_DIR = RAIZ / "config"


def config_dir() -> Path:
    """Pasta da configuracao de negocio da empresa (TINY_APPS_CONFIG ou data/)."""
    env = os.environ.get("TINY_APPS_CONFIG")
    return Path(env).expanduser().resolve() if env else DATA_DIR


def caminho(nome: str) -> Path:
    """Caminho de um arquivo de configuracao de negocio (pode nao existir)."""
    return config_dir() / nome


def ler_json(nome: str, padrao=None):
    """Le um JSON de configuracao de negocio; `padrao` se o arquivo nao existir."""
    p = caminho(nome)
    if not p.exists():
        return padrao
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SystemExit(f"  [ERRO] {p} nao e um JSON valido: {e}")
