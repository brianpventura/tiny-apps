"""Portao contra vazamento: procura dado sensivel nos arquivos que iriam para o git.

Varre os arquivos rastreados e os novos nao ignorados (`git ls-files -co
--exclude-standard`) atras de:

  * padroes fixos: CPF, CNPJ, telefone, email, JWT, client_secret preenchido,
    caminho de usuario do Windows;
  * termos proibidos da empresa, um por linha, em `_privado/termos_proibidos.txt`
    (arquivo local, ignorado pelo git): nomes das empresas, razoes sociais,
    nomes de pessoas, nomes de integracao, SKUs que nao devem aparecer...
    Linhas vazias e comecadas por `#` sao ignoradas. A comparacao ignora
    caixa e acento.

Valores de exemplo (12.345.678/0001-00, (11) 99999-0000, @example.com etc.)
sao permitidos.

Uso (na raiz do repo):
    python scripts/verificar_sensiveis.py            # sai com codigo 1 se achar algo
    python scripts/verificar_sensiveis.py --resumo   # so a contagem por arquivo

Rodar antes de todo commit. Como hook: `.git/hooks/pre-commit` com
    #!/bin/sh
    exec .venv/Scripts/python.exe scripts/verificar_sensiveis.py
"""

import argparse
import re
import subprocess
import sys
import unicodedata
from collections import Counter
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
TERMOS = RAIZ / "_privado" / "termos_proibidos.txt"

PADROES = {
    "cpf": re.compile(r"(?<![\d.])\d{3}\.\d{3}\.\d{3}-\d{2}(?![\d])"),
    "cnpj": re.compile(r"(?<![\d.])\d{2}\.?\d{3}\.?\d{3}/\d{4}-?\d{2}(?!\d)"),
    "telefone": re.compile(r"\(\d{2}\)\s?9?\d{4}-\d{4}|(?<!\d)\d{2}9\d{8}(?!\d)"),
    "email": re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"),
    "jwt": re.compile(r"eyJ[\w-]{10,}\.[\w-]{10,}\.[\w-]{10,}"),
    "client_secret": re.compile(r"client_secret\"\s*:\s*\"(?!SEU_)[^\"]{8,}\""),
    "caminho_usuario": re.compile(r"[A-Za-z]:\\+Users\\+(?!<)[^\\\s\"']+", re.I),
    # Valores de faturamento (R$ 1.000,00 para cima) e ids de pedido/produto
    # (9+ digitos) em comentario ou docstring costumam ser numero real da base.
    "valor_rs": re.compile(r"R\$ ?\d{1,3}(?:\.\d{3})+(?:,\d{2})?"),
    "id_longo": re.compile(r"(?<![\w.])\d{9,}(?![\w.])"),
}

# Valores de exemplo que os docs usam de proposito.
PERMITIDOS = [
    re.compile(r"12\.?345\.?678/0001-?00"),
    re.compile(r"123\.456\.789-00"),
    re.compile(r"\(11\)\s?9?9999-0000|11999990000"),
    re.compile(r"@example\.(com|org)|noreply@anthropic\.com"),
    re.compile(r"(?<!\d)(1{9}|2{9})(?!\d)|R\$ 1\.000,00"),
]

EXTENSOES_TEXTO = {".py", ".sql", ".md", ".txt", ".json", ".toml", ".cfg",
                   ".ini", ".yml", ".yaml", ".csv", ".gitignore", ""}


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s.lower())
    return "".join(c for c in s if not unicodedata.combining(c))


def arquivos() -> list[Path]:
    saida = subprocess.run(
        ["git", "ls-files", "-co", "--exclude-standard"], cwd=RAIZ,
        capture_output=True, text=True, check=True).stdout
    return [RAIZ / p for p in saida.splitlines() if p]


def termos() -> list[str]:
    if not TERMOS.exists():
        return []
    linhas = TERMOS.read_text(encoding="utf-8").splitlines()
    return [_norm(t.strip()) for t in linhas if t.strip() and not t.startswith("#")]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--resumo", action="store_true",
                    help="so a contagem de ocorrencias por arquivo")
    args = ap.parse_args()

    lista_termos = termos()
    if not lista_termos:
        print(f"[AVISO] {TERMOS.relative_to(RAIZ)} ausente ou vazio: "
              "so os padroes fixos serao checados.")

    achados: list[tuple[str, int, str, str]] = []
    binarios: list[str] = []
    for arq in arquivos():
        rel = arq.relative_to(RAIZ).as_posix()
        if arq.suffix.lower() not in EXTENSOES_TEXTO:
            binarios.append(rel)
            continue
        try:
            texto = arq.read_text(encoding="utf-8")
        except (UnicodeDecodeError, FileNotFoundError):
            binarios.append(rel)
            continue
        for n, linha in enumerate(texto.splitlines(), 1):
            limpa = linha
            for p in PERMITIDOS:
                limpa = p.sub("", limpa)
            for nome, rx in PADROES.items():
                for m in rx.finditer(limpa):
                    achados.append((rel, n, nome, m.group(0)))
            ln = _norm(limpa)
            for t in lista_termos:
                if t in ln:
                    achados.append((rel, n, "termo", t))

    if binarios:
        print("Arquivos nao-texto que iriam para o git (confira um a um):")
        for b in binarios:
            print(f"  {b}")
        print()

    if args.resumo:
        for arq, qtd in Counter(a[0] for a in achados).most_common():
            print(f"{qtd:5d}  {arq}")
    else:
        for arq, n, tipo, valor in achados:
            print(f"{arq}:{n}: [{tipo}] {valor}")

    print(f"\n{len(achados)} ocorrencia(s) em "
          f"{len({a[0] for a in achados})} arquivo(s).")
    return 1 if achados else 0


if __name__ == "__main__":
    sys.exit(main())
