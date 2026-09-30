import time
import requests
from .auth_tiny import TinyTokenManager
from .rate_limiter import get_limiter

BASE_URL = "https://api.tiny.com.br/public-api/v3"
_LIMIT               = 100   # itens por página (máximo da API)
_MAX_RETRIES         = 6     # tentativas antes de desistir

# Códigos HTTP que justificam nova tentativa (erros transitórios do servidor)
_STATUS_RETRY = {429, 500, 502, 503, 504}


class TinyClient:
    def __init__(self, tenant_id: str):
        self.tenant_id = tenant_id
        self._token_manager = TinyTokenManager(tenant_id)
        self._session = requests.Session()

    def _get(self, endpoint: str, params: dict = None) -> dict:
        url = f"{BASE_URL}/{endpoint.lstrip('/')}"
        limiter = get_limiter()
        for tentativa in range(_MAX_RETRIES):
            try:
                limiter.acquire()  # aguarda slot disponível no rate limiter
                headers = self._token_manager.auth_header()
                resp = self._session.get(url, headers=headers, params=params, timeout=60)
                if not resp.ok:
                    print(f"\n  [HTTP {resp.status_code}] {endpoint} -> {resp.text[:300]}")
                resp.raise_for_status()
                return resp.json()
            except requests.exceptions.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else 0
                if status == 429:
                    # Extrai Retry-After se disponível
                    retry_after = None
                    if exc.response is not None:
                        ra = exc.response.headers.get("Retry-After")
                        if ra:
                            try:
                                retry_after = float(ra)
                            except ValueError:
                                pass
                    espera = retry_after or 5.0
                    limiter.cooldown(espera)
                    if tentativa < _MAX_RETRIES - 1:
                        print(f"\n  [RATE LIMIT] HTTP 429 — aguardando {espera:.0f}s...")
                        time.sleep(espera)
                    else:
                        raise
                elif status in _STATUS_RETRY and tentativa < _MAX_RETRIES - 1:
                    espera = 5 * (tentativa + 1)  # 5s, 10s, 15s, 20s, 25s, 30s
                    print(f"\n  [RETRY {tentativa+1}/{_MAX_RETRIES-1}] HTTP {status} — aguardando {espera}s...")
                    time.sleep(espera)
                else:
                    raise
            except (requests.exceptions.ConnectionError,
                    requests.exceptions.ChunkedEncodingError,
                    requests.exceptions.Timeout):
                self._session = requests.Session()
                espera = 5 * (tentativa + 1)
                if tentativa < _MAX_RETRIES - 1:
                    print(f"\n  [RETRY {tentativa+1}/{_MAX_RETRIES-1}] Conexão perdida — aguardando {espera}s...")
                    time.sleep(espera)
                else:
                    raise

    def _post(self, endpoint: str, body: dict) -> dict:
        url = f"{BASE_URL}/{endpoint.lstrip('/')}"
        limiter = get_limiter()
        for tentativa in range(_MAX_RETRIES):
            limiter.acquire()
            headers = {
                **self._token_manager.auth_header(),
                "Content-Type": "application/json",
            }
            try:
                resp = self._session.post(url, headers=headers, json=body, timeout=60)
                if not resp.ok:
                    print(f"\n  [HTTP {resp.status_code}] POST {endpoint} -> {resp.text[:300]}")
                resp.raise_for_status()
                return resp.json() if resp.text else {}
            except requests.exceptions.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else 0
                if status == 429:
                    ra = exc.response.headers.get("Retry-After") if exc.response is not None else None
                    espera = float(ra) if ra else 5.0
                    limiter.cooldown(espera)
                    if tentativa < _MAX_RETRIES - 1:
                        print(f"\n  [RATE LIMIT] HTTP 429 — aguardando {espera:.0f}s...")
                        time.sleep(espera)
                    else:
                        raise
                elif status in _STATUS_RETRY and tentativa < _MAX_RETRIES - 1:
                    espera = 5 * (tentativa + 1)
                    print(f"\n  [RETRY {tentativa+1}/{_MAX_RETRIES-1}] HTTP {status} — aguardando {espera}s...")
                    time.sleep(espera)
                else:
                    raise

    def _delete(self, endpoint: str) -> dict:
        url = f"{BASE_URL}/{endpoint.lstrip('/')}"
        limiter = get_limiter()
        for tentativa in range(_MAX_RETRIES):
            limiter.acquire()
            headers = self._token_manager.auth_header()
            try:
                resp = self._session.delete(url, headers=headers, timeout=60)
                if not resp.ok:
                    print(f"\n  [HTTP {resp.status_code}] DELETE {endpoint} -> {resp.text[:300]}")
                resp.raise_for_status()
                return resp.json() if resp.text else {}
            except requests.exceptions.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else 0
                if status == 429:
                    ra = exc.response.headers.get("Retry-After") if exc.response is not None else None
                    espera = float(ra) if ra else 5.0
                    limiter.cooldown(espera)
                    if tentativa < _MAX_RETRIES - 1:
                        print(f"\n  [RATE LIMIT] HTTP 429 — aguardando {espera:.0f}s...")
                        time.sleep(espera)
                    else:
                        raise
                elif status in _STATUS_RETRY and tentativa < _MAX_RETRIES - 1:
                    espera = 5 * (tentativa + 1)
                    print(f"\n  [RETRY {tentativa+1}/{_MAX_RETRIES-1}] HTTP {status} — aguardando {espera}s...")
                    time.sleep(espera)
                else:
                    raise

    def _put(self, endpoint: str, body: dict) -> dict:
        url = f"{BASE_URL}/{endpoint.lstrip('/')}"
        limiter = get_limiter()
        for tentativa in range(_MAX_RETRIES):
            limiter.acquire()
            headers = {
                **self._token_manager.auth_header(),
                "Content-Type": "application/json",
            }
            try:
                resp = self._session.put(url, headers=headers, json=body, timeout=30)
                if not resp.ok:
                    print(f"\n  [HTTP {resp.status_code}] PUT {endpoint} -> {resp.text[:300]}")
                resp.raise_for_status()
                return resp.json() if resp.text else {}
            except requests.exceptions.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else 0
                if status == 429:
                    ra = exc.response.headers.get("Retry-After") if exc.response is not None else None
                    espera = float(ra) if ra else 5.0
                    limiter.cooldown(espera)
                    if tentativa < _MAX_RETRIES - 1:
                        print(f"\n  [RATE LIMIT] HTTP 429 — aguardando {espera:.0f}s...")
                        time.sleep(espera)
                    else:
                        raise
                elif status in _STATUS_RETRY and tentativa < _MAX_RETRIES - 1:
                    espera = 5 * (tentativa + 1)
                    print(f"\n  [RETRY {tentativa+1}/{_MAX_RETRIES-1}] HTTP {status} — aguardando {espera}s...")
                    time.sleep(espera)
                else:
                    raise

    # ------------------------------------------------------------------ #
    #  Pedidos                                                             #
    # ------------------------------------------------------------------ #

    def listar_pedidos(self, offset: int = 0, **filtros) -> dict:
        """Lista pedidos com paginação por offset (do mais recente ao mais antigo).

        Para filtrar por período use `dataInicial`/`dataFinal` (YYYY-MM-DD), que
        funcionam no servidor. ATENCAO: `dataInicio`/`dataFim` sao aceitos e
        IGNORADOS pela API. Pagine dentro de janelas mensais: offset profundo
        devolve 400 "Sua consulta levou muito tempo".
        """
        return self._get("pedidos", params={"offset": offset, "limit": _LIMIT, **filtros})

    def obter_pedido(self, pedido_id) -> dict:
        return self._get(f"pedidos/{pedido_id}")

    def listar_produtos(self, offset: int = 0, **filtros) -> dict:
        """Lista produtos com paginação por offset (mesma estrutura de /pedidos)."""
        return self._get("produtos", params={"offset": offset, "limit": _LIMIT, **filtros})

    def obter_produto(self, produto_id) -> dict:
        return self._get(f"produtos/{produto_id}")

    def buscar_produto_por_sku(self, sku: str, debug: bool = False) -> dict | None:
        """Retorna o produto correspondente ao SKU, ou None.

        O filtro ?codigo=<sku> faz match exato server-side, mas pode devolver MAIS DE UM
        item: o mesmo SKU costuma existir duas vezes no cadastro — o produto vivo e um
        gemeo excluido logicamente (situacao 'E'; nao existe DELETE de produto usado em
        pedido). Descartar o excluido e o que faz esses SKUs serem encontrados.

        Atencao ao nome do campo: na listagem quem vem preenchido e 'sku'; 'codigo' vem
        sempre null. Comparar por 'codigo' nunca casa.
        """
        sku_alvo = str(sku).strip()
        resp = self._get("produtos", params={"codigo": sku_alvo, "limit": _LIMIT})
        if "data" in resp:
            resp = resp["data"]
        itens = resp.get("itens", [])

        if debug:
            print(f"\n    [debug] codigo={sku_alvo!r} retornou {len(itens)} item(ns)")
            for it in itens[:3]:
                print(f"      id={it.get('id')}  sku={it.get('sku')!r}  "
                      f"situacao={it.get('situacao')!r}  "
                      f"desc={str(it.get('descricao', ''))[:50]!r}")

        exatos = [it for it in itens
                  if str(it.get("sku") or it.get("codigo") or "").strip() == sku_alvo]
        candidatos = exatos or itens

        vivos = [it for it in candidatos
                 if str(it.get("situacao") or "").strip().upper() != "E"]
        if len(vivos) == 1:
            return vivos[0]
        if len(candidatos) == 1:
            return candidatos[0]
        return None   # ambiguo de verdade: dois cadastros vivos com o mesmo SKU

    def atualizar_produto(self, produto_id, body: dict) -> dict:
        """PUT /produtos/{id} — atualiza campos do produto."""
        return self._put(f"produtos/{produto_id}", body)

    def atualizar_variacao(self, produto_pai_id, variacao_id, body: dict) -> dict:
        """PUT /produtos/{idPai}/variacoes/{idVariacao} — atualiza uma variacao.

        Variacoes NAO sao atualizaveis por PUT /produtos/{id} (retorna 404) nem
        pelo array 'variacoes' do pai (o custo e ignorado). Este e o unico caminho
        que persiste precos.precoCusto de uma variacao.
        """
        return self._put(f"produtos/{produto_pai_id}/variacoes/{variacao_id}", body)

    def enviar_produto(self, produto_id, body: dict) -> dict:
        """POST /produtos/{id} — envia JSON completo do produto."""
        return self._post(f"produtos/{produto_id}", body)

    def excluir_produto(self, produto_id) -> dict:
        """DELETE /produtos/{id} — remove o cadastro (falha se houver vinculos)."""
        return self._delete(f"produtos/{produto_id}")

    # ------------------------------------------------------------------ #
    #  Categorias                                                          #
    # ------------------------------------------------------------------ #

    def criar_categoria(self, descricao: str, id_categoria_pai: int | None = None) -> dict:
        """POST /categorias — cria uma categoria (raiz ou filha de id_categoria_pai).

        Retorna o JSON da API (contendo o 'id' da categoria criada).
        """
        body = {"descricao": descricao}
        if id_categoria_pai is not None:
            body["idCategoriaPai"] = id_categoria_pai
        return self._post("categorias", body)

    def excluir_categoria(self, id_categoria) -> dict:
        """DELETE /categorias/{id} — remove a categoria.

        O Tiny pode recusar (400/403) se a categoria ainda tiver produtos
        vinculados ou subcategorias. O chamador deve tratar a excecao.
        """
        return self._delete(f"categorias/{id_categoria}")

    def listar_categorias_arvore(self) -> list:
        """GET /categorias/todas — retorna a arvore de categorias.

        Cada no tem {id, descricao, filhas:[...]} (recursivo). Desembrulha o
        eventual envelope {'data': ...} e normaliza para uma lista de raizes.
        """
        resp = self._get("categorias/todas")
        if isinstance(resp, dict) and "data" in resp:
            resp = resp["data"]
        if isinstance(resp, dict):
            # alguns endpoints devolvem {'itens': [...]} ou um unico no raiz
            if "itens" in resp:
                return resp["itens"]
            return [resp]
        return resp if isinstance(resp, list) else []

    def produtos_todos(self, on_total=None):
        """Gerador que percorre todos os produtos cadastrados via paginação por offset.

        on_total: callback(total) chamado após a primeira página.
        """
        offset = 0
        primeira = True

        while True:
            resp = None
            for tentativa in range(_MAX_RETRIES):
                try:
                    resp = self.listar_produtos(offset=offset)
                    break
                except requests.exceptions.HTTPError as exc:
                    status = exc.response.status_code if exc.response is not None else 0
                    if status == 400:
                        if tentativa < _MAX_RETRIES - 1:
                            time.sleep(2 ** tentativa)
                        else:
                            return
                    else:
                        raise

            if resp is None:
                return

            if "data" in resp:
                resp = resp["data"]

            itens = resp.get("itens", [])
            paginacao = resp.get("paginacao", {})
            total = paginacao.get("total")

            if primeira and on_total is not None:
                on_total(total)
                primeira = False

            if not itens:
                return

            for item in itens:
                yield item

            offset += _LIMIT
            if total is not None and offset >= total:
                return
            # Rate limiter controla o ritmo — sem delay fixo aqui

    def listar_vendedores(self, nome: str = None, **filtros) -> dict:
        """GET /vendedores — lista vendedores (filtra por nome parcial se informado)."""
        params = dict(filtros)
        if nome:
            params["nome"] = nome
        return self._get("vendedores", params=params)

    def listar_contatos(self, pagina: int = 1, **filtros) -> dict:
        return self._get("contatos", params={"pagina": pagina, **filtros})

    def buscar_contatos(self, **filtros) -> dict:
        """GET /contatos — busca por filtros (ex: cpfCnpj, nome, celular)."""
        return self._get("contatos", params=filtros)

    def obter_contato(self, id_contato) -> dict:
        """GET /contatos/{id} — contato completo (inclui telefone e celular)."""
        return self._get(f"contatos/{id_contato}")

    def atualizar_contato(self, id_contato, body: dict) -> dict:
        """PUT /contatos/{id} — atualiza campos do contato."""
        return self._put(f"contatos/{id_contato}", body)

    def excluir_contato(self, id_contato, nome: str, tipos_ids: list | None = None) -> dict:
        """Exclusao LOGICA de contato via PUT /contatos/{id} com situacao='E'.

        A API v3 do Tiny NAO expoe DELETE /contatos/{id} (retorna 400 'metodo
        DELETE nao e compativel'). O unico DELETE de contatos e para a
        subentidade 'pessoas' (DELETE /contatos/{id}/pessoas/{pessoaId}), que
        nao remove o cadastro. A exclusao do cadastro e feita marcando
        situacao='E' (Excluido) — os demais campos sao preservados quando
        omitidos. 'nome' e reenviado (obrigatorio); 'tipos' e reenviado quando
        informado para evitar que o Tiny adicione um tipo 'Outro' indevido.
        """
        body = {"nome": nome, "situacao": "E"}
        if tipos_ids:
            body["tipos"] = tipos_ids
        return self._put(f"contatos/{id_contato}", body)

    def contatos_todos(self, on_total=None):
        """Gerador que percorre todos os contatos via paginacao por offset."""
        offset = 0
        total = None
        primeira = True
        while True:
            resp = self._get("contatos", params={"offset": offset, "limit": _LIMIT})
            if "data" in resp:
                resp = resp["data"]
            itens = resp.get("itens", [])
            if total is None:
                total = resp.get("paginacao", {}).get("total", 0)
                if primeira and on_total is not None:
                    on_total(total)
                    primeira = False
            if not itens:
                return
            for item in itens:
                yield item
            offset += _LIMIT
            if total is not None and offset >= total:
                return

    # ------------------------------------------------------------------ #
    #  Estoque                                                             #
    # ------------------------------------------------------------------ #

    def obter_estoque(self, produto_id) -> dict:
        """GET /estoque/{id} — saldo do produto, total e por deposito.

        Retorna {id, nome, codigo, unidade, saldo, reservado, disponivel,
        localizacao, depositos:[{id, nome, desconsiderar, saldo, reservado,
        disponivel, empresa}]}.
        """
        return self._get(f"estoque/{produto_id}")

    def listar_depositos(self) -> list:
        """GET /depositos — lista os depositos da empresa.

        A doc mostra um array puro, mas a API responde {'itens': [...]};
        normalizamos para lista.
        """
        resp = self._get("depositos")
        if isinstance(resp, dict):
            return resp.get("itens") or resp.get("data") or []
        return resp if isinstance(resp, list) else []

    def atualizar_estoque(self, produto_id, tipo: str, quantidade: float,
                          preco_unitario: float, data: str | None = None,
                          observacoes: str | None = None,
                          id_deposito: int | None = None) -> dict:
        """POST /estoque/{id} — lanca uma movimentacao de estoque.

        tipo: 'B' balanco (define o saldo absoluto), 'E' entrada (soma),
        'S' saida (subtrai). quantidade e preco_unitario sao obrigatorios
        pela API. Omitir id_deposito faz o Tiny usar o deposito padrao.

        IMPORTANTE: com tipo 'B', se a quantidade pedida for IGUAL ao saldo
        atual o Tiny responde HTTP 204 sem corpo e NAO cria movimentacao —
        a observacao tambem nao e gravada. So ha lancamento (HTTP 200 +
        idLancamento) quando o saldo realmente muda. Confirmado por teste;
        por isso o retorno vazio deve ser tratado como "sem efeito", nunca
        como sucesso.

        data: string 'YYYY-MM-DD HH:MM:SS' (separador espaco, sem timezone).
        A doc so diz "string", mas a API exige data E hora nesse formato
        exato — confirmado por teste: 'YYYY-MM-DD', ISO com 'T', com 'Z' ou
        com offset ('-03:00') e 'dd/mm/aaaa' retornam todos HTTP 400
        {"campo":"data","mensagem":"Este valor nao e uma data e hora valida"}.
        Omitir o campo tambem e aceito (o Tiny usa a data/hora atual).

        ATENCAO: o lancamento e irreversivel — nao existe DELETE de
        movimentacao. Um erro so se corrige com outro lancamento.

        Retorna {'idLancamento': N}.
        """
        tipo = str(tipo).strip().upper()
        if tipo not in ("B", "E", "S"):
            raise ValueError(f"tipo invalido: {tipo!r} (esperado 'B', 'E' ou 'S')")

        body: dict = {
            "tipo": tipo,
            "quantidade": float(quantidade),
            "precoUnitario": float(preco_unitario),
        }
        if id_deposito is not None:
            body["deposito"] = {"id": int(id_deposito)}
        if data:
            body["data"] = data
        if observacoes:
            body["observacoes"] = observacoes

        return self._post(f"estoque/{produto_id}", body)
