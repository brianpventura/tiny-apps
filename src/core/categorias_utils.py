"""Utilitarios para a arvore de categorias do Tiny V3.

A arvore vem de GET /categorias/todas: cada no tem {id, descricao, filhas:[...]}.
Aqui achatamos a arvore e resolvemos pares (categoria, subcategoria) -> id.
"""


def norm_cat(s: str) -> str:
    """Normaliza nome de categoria para comparacao (trim + casefold)."""
    return str(s or "").strip().casefold()


def achatar_categorias(arvore: list) -> list[dict]:
    """Percorre a arvore e retorna [{id, nome, pai_nome, caminho}] para cada no."""
    flat: list[dict] = []

    def visitar(no, pai_nome, prefixo):
        nome = no.get("descricao") or no.get("nome") or ""
        caminho = f"{prefixo} >> {nome}" if prefixo else nome
        flat.append({
            "id": no.get("id"),
            "nome": nome,
            "pai_nome": pai_nome,
            "caminho": caminho,
        })
        for filha in (no.get("filhas") or []):
            visitar(filha, nome, caminho)

    for raiz in (arvore or []):
        visitar(raiz, None, "")
    return flat


class ResolvedorCategoria:
    """Resolve (categoria, subcategoria) -> {id, caminho, erro}, com cache.

    Modelo de 2 niveis:
      - categoria sem subcategoria  -> categoria RAIZ (pai_nome=None) com aquele nome
      - categoria + subcategoria    -> filha 'subcategoria' cujo pai se chama 'categoria'
    Nomes repetidos em ramos diferentes sao desambiguados pelo nome do pai.
    """

    def __init__(self, flat: list[dict]):
        self.flat = flat
        self._cache: dict[tuple, dict] = {}

    def atualizar(self, flat: list[dict]) -> None:
        """Substitui a base achatada (apos criar novas categorias) e limpa cache."""
        self.flat = flat
        self._cache.clear()

    def resolver(self, categoria: str, subcategoria: str) -> dict:
        chave = (norm_cat(categoria), norm_cat(subcategoria))
        if chave not in self._cache:
            self._cache[chave] = self._resolver(categoria, subcategoria)
        return self._cache[chave]

    def _resolver(self, categoria: str, subcategoria: str) -> dict:
        cat, sub = norm_cat(categoria), norm_cat(subcategoria)

        if not cat:
            return {"id": None, "caminho": None, "erro": "sem_categoria_na_planilha"}

        if sub:
            matches = [n for n in self.flat
                       if norm_cat(n["nome"]) == sub and norm_cat(n["pai_nome"] or "") == cat]
            alvo = f"{categoria} >> {subcategoria}"
        else:
            matches = [n for n in self.flat
                       if norm_cat(n["nome"]) == cat and n["pai_nome"] is None]
            alvo = categoria

        if len(matches) == 1:
            return {"id": matches[0]["id"], "caminho": matches[0]["caminho"], "erro": None}
        if not matches:
            return {"id": None, "caminho": None, "erro": f"categoria_inexistente ({alvo})"}
        caminhos = " | ".join(m["caminho"] for m in matches)
        return {"id": None, "caminho": None,
                "erro": f"ambiguo: {len(matches)} categorias casam [{caminhos}]"}

    # ------------------------------------------------------------------ #
    #  Apoio a criacao de categorias faltantes                            #
    # ------------------------------------------------------------------ #

    def id_raiz(self, categoria: str):
        """Retorna o id da categoria raiz com esse nome, ou None."""
        cat = norm_cat(categoria)
        for n in self.flat:
            if norm_cat(n["nome"]) == cat and n["pai_nome"] is None:
                return n["id"]
        return None

    def existe_filha(self, categoria_pai: str, subcategoria: str) -> bool:
        """True se ja existe a filha 'subcategoria' sob o pai 'categoria_pai'."""
        cat, sub = norm_cat(categoria_pai), norm_cat(subcategoria)
        return any(norm_cat(n["nome"]) == sub and norm_cat(n["pai_nome"] or "") == cat
                   for n in self.flat)
