from __future__ import annotations

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_TIMEOUT_PADRAO = (10, 180)


def criar_sessao(pool_size: int = 16) -> requests.Session:
    sessao = requests.Session()
    retry = Retry(
        total=0,
        connect=2,
        read=0,
        backoff_factor=0.5,
        status_forcelist=[],
        allowed_methods=frozenset(["POST"]),
    )
    adapter = HTTPAdapter(
        pool_connections=pool_size,
        pool_maxsize=pool_size,
        max_retries=retry,
    )
    sessao.mount("https://", adapter)
    sessao.mount("http://", adapter)
    return sessao


SESSAO = criar_sessao()
