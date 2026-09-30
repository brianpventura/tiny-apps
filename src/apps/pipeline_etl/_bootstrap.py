"""Bootstrap de sys.path e diretorios do pipeline_etl.

Todo modulo `etl_*` e o orquestrador comecam com `import _bootstrap`. E o mesmo
idioma dos outros apps do repo (sobe os `parents` ate achar a pasta que contem
`core/`), com um acrescimo: a propria pasta do pipeline entra no sys.path,
para os modulos irmaos se importarem sem virar pacote.

As regras de negocio vem de `core/regras_vendas.py` (via `etl_regras.py`).

ATENCAO: a pasta do pipeline fica na frente da stdlib no sys.path. Por isso todo
modulo daqui tem prefixo `etl_` -- um arquivo chamado `types.py` ou `json.py`
nesta pasta derrubaria o interpretador.
"""

import sys
from pathlib import Path

_AQUI = Path(__file__).resolve().parent
_SRC = next(p for p in _AQUI.parents if (p / "core").is_dir())
_RAIZ = _SRC.parent

for _p in (str(_AQUI), str(_SRC)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

AQUI = _AQUI
SRC = _SRC
RAIZ = _RAIZ
DATA_DIR = _RAIZ / "data"
REPORTS_DIR = _RAIZ / "reports"
SQL_DIR = _AQUI / "sql"
TENANTS_FILE = DATA_DIR / "tenants.json"
DB_FILE = DATA_DIR / "db.json"
