"""
Rate Limiter centralizado para a API V3 do Tiny.

Estratégia simples e previsível:
  - Garante intervalo mínimo de 0.5s entre requisições (= 2 req/s = 120 req/min).
  - Thread-safe: todas as threads passam pela mesma fila.
  - Em caso de 429, aplica cooldown antes de liberar novas requisições.
"""

import time
import threading


class RateLimiter:
    """Controla a taxa de requisições com intervalo fixo entre chamadas.

    Parâmetros:
        requests_per_second: taxa máxima de requisições por segundo.
    """

    def __init__(self, requests_per_second: float = 2.0):
        self._interval = 1.0 / requests_per_second  # 0.5s para 2 req/s
        self._lock = threading.Lock()
        self._last_request: float = 0.0  # timestamp da última requisição

    def acquire(self) -> None:
        """Bloqueia até que o intervalo mínimo desde a última requisição tenha passado."""
        with self._lock:
            now = time.monotonic()
            elapsed = now - self._last_request
            if elapsed < self._interval:
                wait = self._interval - elapsed
                time.sleep(wait)
            self._last_request = time.monotonic()

    def cooldown(self, seconds: float) -> None:
        """Aplica um cooldown global (ex: após receber 429).

        Avança o timestamp da última requisição para forçar todas as threads
        a esperarem pelo menos 'seconds' antes da próxima chamada.
        """
        with self._lock:
            self._last_request = time.monotonic() + seconds - self._interval


# ------------------------------------------------------------------ #
#  Instância global compartilhada por todos os TinyClients             #
# ------------------------------------------------------------------ #

_global_limiter = RateLimiter(requests_per_second=2.0)


def get_limiter() -> RateLimiter:
    """Retorna o rate limiter global."""
    return _global_limiter
