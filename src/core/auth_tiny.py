import json
import time
import threading
import webbrowser
import requests
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import urlencode, urlparse, parse_qs

CALLBACK_HOST = "localhost"
CALLBACK_PORT = 8765
CALLBACK_PATH = "/callback"
CALLBACK_TIMEOUT_SECONDS = 300  # 5 minutos para o usuário concluir login

TOKEN_URL = "https://accounts.tiny.com.br/realms/tiny/protocol/openid-connect/token"
AUTH_URL = "https://accounts.tiny.com.br/realms/tiny/protocol/openid-connect/auth"
TENANTS_FILE = Path(__file__).resolve().parent.parent.parent / "data" / "tenants.json"
REFRESH_BUFFER_SECONDS = 60  # renova se restar menos de 60s de validade



class TokenExpiredError(Exception):
    """Refresh token expirado — re-autenticação manual necessária."""


class TinyTokenManager:
    def __init__(self, tenant_id: str):
        self.tenant_id = tenant_id
        self._token_file = Path(__file__).resolve().parent.parent.parent / "data" / f"tokens_{self.tenant_id}.json"
        self._tokens: dict = {}
        
        self._load_tenant_config()
        self._load()

    def _load_tenant_config(self) -> None:
        if not TENANTS_FILE.exists():
            raise FileNotFoundError(
                f"Arquivo de configuração {TENANTS_FILE} não encontrado. "
                "Copie data/tenants.exemplo.json para data/tenants.json e preencha.")
        
        tenants_data = json.loads(TENANTS_FILE.read_text(encoding="utf-8"))
        if self.tenant_id not in tenants_data:
            raise ValueError(f"Tenant '{self.tenant_id}' não configurado em {TENANTS_FILE.name}.")
            
        config = tenants_data[self.tenant_id]
        self.client_id = config["client_id"]
        self.client_secret = config["client_secret"]
        self.redirect_uri = config["redirect_uri"]

    # ------------------------------------------------------------------ #
    #  Persistência                                                        #
    # ------------------------------------------------------------------ #

    def _load(self) -> None:
        if self._token_file.exists():
            self._tokens = json.loads(self._token_file.read_text(encoding="utf-8"))

    def _save(self) -> None:
        self._token_file.write_text(
            json.dumps(self._tokens, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    # ------------------------------------------------------------------ #
    #  Inicialização (primeiro uso ou tokens expirados)                   #
    # ------------------------------------------------------------------ #

    def bootstrap(self, authorization_code: str) -> None:
        """Troca o authorization_code por access + refresh token e persiste."""
        payload = {
            "grant_type": "authorization_code",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "redirect_uri": self.redirect_uri,
            "code": authorization_code,
        }
        resp = requests.post(TOKEN_URL, data=payload, timeout=15)
        if not resp.ok:
            print(f"\nErro {resp.status_code} ao tentar gerar tokens (bootstrap):")
            print(f"  client_id: {self.client_id}")
            print(f"  redirect_uri: {self.redirect_uri}")
            print(f"  response: {resp.text}")
            if resp.status_code == 401:
                print("\n  -> 401 = client_id ou client_secret INCORRETOS.")
                print("    Verifique as credenciais no painel do Tiny (Configuracoes -> Aplicativos).")
            elif resp.status_code == 400:
                error_desc = resp.json().get("error_description", "")
                if "redirect_uri" in error_desc.lower():
                    print("\n  -> redirect_uri nao confere com o cadastrado no aplicativo do Tiny.")
                elif "code" in error_desc.lower() or "grant" in error_desc.lower():
                    print("\n  -> Authorization code invalido ou ja utilizado. Gere um novo.")
        resp.raise_for_status()
        self._store(resp.json())
        print(f"  Autenticacao concluida - tokens salvos em {self._token_file.name}")

    # ------------------------------------------------------------------ #
    #  Lazy refresh                                                        #
    # ------------------------------------------------------------------ #

    def _is_expired(self) -> bool:
        expires_at = self._tokens.get("expires_at", 0)
        return time.time() >= expires_at - REFRESH_BUFFER_SECONDS

    def _is_refresh_expired(self) -> bool:
        """Verifica se o refresh token expirou."""
        refresh_expires_at = self._tokens.get("refresh_expires_at", 0)
        if not refresh_expires_at:
            return False  # sem info de expiração, assume válido
        return time.time() >= refresh_expires_at - REFRESH_BUFFER_SECONDS

    def _refresh(self) -> None:
        refresh_token = self._tokens.get("refresh_token")
        if not refresh_token:
            raise TokenExpiredError("Nenhum refresh_token disponivel.")

        if self._is_refresh_expired():
            raise TokenExpiredError(
                f"Refresh token do tenant '{self.tenant_id}' expirado. "
                "Re-autenticacao necessaria."
            )

        payload = {
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "refresh_token": refresh_token,
        }
        resp = requests.post(TOKEN_URL, data=payload, timeout=15)

        if resp.status_code == 400:
            raise TokenExpiredError(
                "Refresh token invalido ou expirado. "
                "Re-autenticacao necessaria."
            )

        resp.raise_for_status()
        self._store(resp.json())
        print(f"  Access token renovado automaticamente ({self.tenant_id}).")

    def _store(self, data: dict) -> None:
        self._tokens = {
            "access_token": data["access_token"],
            "refresh_token": data.get("refresh_token", self._tokens.get("refresh_token")),
            "token_type": data.get("token_type", "Bearer"),
            "expires_at": time.time() + data["expires_in"],
            "refresh_expires_at": (
                time.time() + data["refresh_expires_in"]
                if "refresh_expires_in" in data
                else self._tokens.get("refresh_expires_at", 0)
            ),
        }
        self._save()

    # ------------------------------------------------------------------ #
    #  Diagnóstico e validação                                             #
    # ------------------------------------------------------------------ #

    def get_auth_url(self) -> str:
        """Gera a URL de autorização OAuth2 para o usuário visitar no navegador."""
        params = {
            "client_id": self.client_id,
            "redirect_uri": self.redirect_uri,
            "response_type": "code",
            "scope": "openid email offline_access",
            "prompt": "login",  # forca re-login mesmo com sessao ativa (evita usar conta errada)
        }
        return f"{AUTH_URL}?{urlencode(params)}"

    def is_healthy(self) -> tuple[bool, str]:
        """Verifica se os tokens estão válidos e utilizáveis.

        Returns:
            (healthy, message) — True se a API pode ser usada, mensagem descritiva.
        """
        if not self._tokens:
            return False, "Sem tokens - bootstrap necessario"

        if self._is_refresh_expired():
            return False, "Refresh token expirado - re-autenticacao necessaria"

        if self._is_expired():
            try:
                self._refresh()
                return True, "Access token renovado"
            except TokenExpiredError:
                return False, "Refresh token invalido - re-autenticacao necessaria"

        # Calcula tempo restante do access e refresh tokens
        access_remaining = self._tokens.get("expires_at", 0) - time.time()
        refresh_remaining = self._tokens.get("refresh_expires_at", 0) - time.time()

        access_info = f"access={self._format_time(access_remaining)}"
        refresh_info = (
            f"refresh={self._format_time(refresh_remaining)}"
            if refresh_remaining > 0
            else "refresh=sem info"
        )

        return True, f"OK ({access_info}, {refresh_info})"

    @staticmethod
    def _format_time(seconds: float) -> str:
        """Formata segundos em texto legível (ex: '3h 20min')."""
        if seconds <= 0:
            return "expirado"
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        if h > 0:
            return f"{h}h {m}min"
        return f"{m}min"

    # ------------------------------------------------------------------ #
    #  Interface pública                                                   #
    # ------------------------------------------------------------------ #

    def get_access_token(self) -> str:
        """Retorna um access token válido, renovando-o se necessário (lazy)."""
        if not self._tokens:
            raise TokenExpiredError(
                "Tokens não encontrados. Chame bootstrap(authorization_code) primeiro."
            )
        if self._is_expired():
            self._refresh()
        return self._tokens["access_token"]

    def auth_header(self) -> dict:
        return {"Authorization": f"Bearer {self.get_access_token()}"}


# ====================================================================== #
#  Captura automática do authorization_code via servidor local            #
# ====================================================================== #


class _CallbackHandler(BaseHTTPRequestHandler):
    captured: dict = {}

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path != CALLBACK_PATH:
            self.send_response(404)
            self.end_headers()
            return

        params = parse_qs(parsed.query)
        code = params.get("code", [None])[0]
        error = params.get("error", [None])[0]

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()

        if code:
            _CallbackHandler.captured["code"] = code
            body = (
                "<html><body style='font-family:sans-serif;text-align:center;padding:40px'>"
                "<h2>Autenticacao concluida</h2>"
                "<p>Voce ja pode fechar esta aba e voltar ao terminal.</p>"
                "</body></html>"
            )
        else:
            _CallbackHandler.captured["error"] = error or "code ausente"
            body = (
                f"<html><body style='font-family:sans-serif;text-align:center;padding:40px'>"
                f"<h2>Falha na autenticacao</h2><p>{error or 'code ausente'}</p>"
                f"</body></html>"
            )
        self.wfile.write(body.encode("utf-8"))

    def log_message(self, format, *args):  # silencia logs do http.server
        return


def capture_authorization_code(auth_url: str, timeout: int = CALLBACK_TIMEOUT_SECONDS) -> str:
    """Abre o navegador na auth_url e captura o `code` via servidor local.

    Sobe um HTTP server em http://localhost:8765/callback, abre o navegador
    padrão na URL de autorização e bloqueia até receber o redirect do Tiny
    (ou até o timeout). Retorna o `code` capturado.
    """
    _CallbackHandler.captured = {}
    server = HTTPServer((CALLBACK_HOST, CALLBACK_PORT), _CallbackHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    try:
        print(f"  Abrindo navegador para autorizacao...")
        print(f"  (se nao abrir, acesse manualmente: {auth_url})")
        webbrowser.open(auth_url)

        deadline = time.time() + timeout
        while time.time() < deadline:
            if "code" in _CallbackHandler.captured:
                return _CallbackHandler.captured["code"]
            if "error" in _CallbackHandler.captured:
                raise RuntimeError(
                    f"Erro retornado pelo Tiny: {_CallbackHandler.captured['error']}"
                )
            time.sleep(0.3)

        raise TimeoutError(
            f"Timeout de {timeout}s aguardando autorizacao no navegador."
        )
    finally:
        server.shutdown()
        server.server_close()


# ====================================================================== #
#  Validação de tenants (automática + interativa)                         #
# ====================================================================== #

def _carregar_nomes() -> dict[str, str]:
    """tenant_id -> nome de exibicao, lido do campo `nome` de tenants.json.

    Tenant sem `nome` usa o proprio id. Sem o arquivo devolve {} para que o
    import nao falhe; quem precisar de credencial falha depois, com mensagem.
    """
    if not TENANTS_FILE.exists():
        return {}
    dados = json.loads(TENANTS_FILE.read_text(encoding="utf-8"))
    return {t: (c.get("nome") or t) for t, c in dados.items()}


TENANT_NAMES = _carregar_nomes()


def get_all_tenant_ids() -> list[str]:
    """Retorna todos os tenant_ids configurados em tenants.json."""
    return list(TENANT_NAMES)


def resolver_tenant(tenant: str | None = None, perguntar: bool = True) -> str:
    """Valida o tenant informado ou escolhe um entre os de tenants.json.

    Sem argumento: com um unico tenant configurado usa ele; com varios,
    pergunta (se `perguntar`) mostrando as opcoes. Tenant desconhecido encerra
    com a lista dos validos -- nunca cai num default silencioso.
    """
    ids = get_all_tenant_ids()
    if not ids:
        raise SystemExit(f"Nenhum tenant em {TENANTS_FILE}. "
                         "Copie data/tenants.exemplo.json e preencha.")
    if not tenant:
        if len(ids) == 1:
            return ids[0]
        if not perguntar:
            raise SystemExit(f"Informe o tenant: {', '.join(ids)}")
        tenant = input(f"Tenant ({', '.join(ids)}): ").strip()
    if tenant not in ids:
        raise SystemExit(f"Tenant desconhecido: {tenant!r}. "
                         f"Configurados: {', '.join(ids)}")
    return tenant


def validate_tenants(tenant_ids: list[str] | None = None, interactive: bool = True) -> bool:
    """Valida e renova tokens de todos (ou dos especificados) tenants.

    - Access token expirado → renova automaticamente via refresh token
    - Refresh token expirado / sem tokens → solicita authorization_code interativamente

    Args:
        tenant_ids: lista de tenants a validar (None = todos).
        interactive: se True, oferece bootstrap interativo para tenants com problemas.

    Returns:
        True se todos os tenants estão saudáveis ao final.
    """
    if tenant_ids is None:
        tenant_ids = get_all_tenant_ids()

    print("=" * 60)
    print("  VALIDACAO DE TOKENS — Tiny API V3")
    print("=" * 60)

    all_healthy = True
    unhealthy: list[tuple[str, TinyTokenManager]] = []

    for tenant_id in tenant_ids:
        nome = TENANT_NAMES.get(tenant_id, tenant_id)
        tm = TinyTokenManager(tenant_id)
        healthy, msg = tm.is_healthy()

        status = "OK" if healthy else "!!"
        print(f"  [{status}] {nome:20s} {msg}")

        if not healthy:
            all_healthy = False
            unhealthy.append((tenant_id, tm))

    if all_healthy:
        print("\n  Todos os tenants autenticados com sucesso.\n")
        return True

    if not interactive:
        print("\n  Alguns tenants precisam de re-autenticacao.")
        print("  Execute validate_tenants(interactive=True) para corrigir.\n")
        return False

    # Fluxo interativo para cada tenant com problemas
    print(f"\n  {len(unhealthy)} tenant(s) precisam de re-autenticacao.\n")

    for tenant_id, tm in unhealthy:
        nome = TENANT_NAMES.get(tenant_id, tenant_id)
        print(f"  --- {nome.upper()} ({tenant_id}) ---")
        print(f"  Faca login no Tiny com a conta da {nome} no navegador que sera aberto.")
        resp = input(f"  Pressione Enter para abrir o navegador (ou 's' para pular): ").strip().lower()

        if resp == "s":
            print(f"  -> Pulando {nome}\n")
            continue

        try:
            code = capture_authorization_code(tm.get_auth_url())
            tm.bootstrap(code)
            print()
        except Exception as e:
            print(f"  FALHA: {e}\n")
            all_healthy = False

    # Resumo final — recalcula a saude real apos as re-autenticacoes
    print("-" * 60)
    all_healthy = True
    for tenant_id in tenant_ids:
        nome = TENANT_NAMES.get(tenant_id, tenant_id)
        tm = TinyTokenManager(tenant_id)
        healthy, msg = tm.is_healthy()
        status = "OK" if healthy else "!!"
        print(f"  [{status}] {nome:20s} {msg}")
        if not healthy:
            all_healthy = False

    print()
    return all_healthy
