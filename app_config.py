"""
Configuração administrada do app.

A chave OpenRouter e os modelos ficam salvos em config_app.json (local, fora
do git), protegidos por senha de administrador. Usuários finais não configuram
nada — só fazem upload dos arquivos.

Segurança (limites honestos, documentados no README):
- A senha é guardada como hash PBKDF2-SHA256 com salt (não recuperável).
- A chave da API fica em claro no JSON local: qualquer pessoa com acesso ao
  ARQUIVO no disco consegue lê-la. A senha protege contra alteração/leitura
  PELA INTERFACE, que é o cenário relevante quando o app roda num servidor
  interno e os usuários só têm o navegador.
"""
from __future__ import annotations

import hashlib
import json
import os
import secrets as _secrets

_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_app.json")

_DEFAULTS = {
    "openrouter_api_key": "",
    "modelo": "google/gemini-2.5-flash",
    "modelo_escalonamento": "google/gemini-2.5-pro",
    "usar_ia_juiz": True,
    "classificar_emails_com_ia": True,
    "admin_password_hash": "",
}

_PBKDF2_ITERS = 200_000


def carregar_config(path: str | None = None) -> dict:
    """Carrega a configuração; devolve defaults para chaves ausentes."""
    cfg = dict(_DEFAULTS)
    try:
        with open(path or _CONFIG_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if isinstance(dados, dict):
            cfg.update({k: v for k, v in dados.items() if k in _DEFAULTS})
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    return cfg


def salvar_config(cfg: dict, path: str | None = None) -> bool:
    """Persiste a configuração (somente chaves conhecidas)."""
    dados = {k: cfg.get(k, _DEFAULTS[k]) for k in _DEFAULTS}
    try:
        with open(path or _CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Senha de administrador
# ---------------------------------------------------------------------------

def _hash_senha(senha: str, salt: str) -> str:
    dk = hashlib.pbkdf2_hmac(
        "sha256", senha.encode("utf-8"), salt.encode("utf-8"), _PBKDF2_ITERS
    )
    return dk.hex()


def definir_senha(cfg: dict, senha: str) -> dict:
    """Define/troca a senha admin. Retorna o cfg atualizado (não salva)."""
    salt = _secrets.token_hex(16)
    cfg["admin_password_hash"] = f"{salt}${_hash_senha(senha, salt)}"
    return cfg


def tem_senha(cfg: dict) -> bool:
    return bool(cfg.get("admin_password_hash"))


def verificar_senha(cfg: dict, senha: str) -> bool:
    guardado = cfg.get("admin_password_hash") or ""
    if "$" not in guardado:
        return False
    salt, hash_salvo = guardado.split("$", 1)
    return _secrets.compare_digest(_hash_senha(senha or "", salt), hash_salvo)


def config_administrada(cfg: dict) -> bool:
    """True quando o admin já configurou a chave — usuários não precisam de nada."""
    return bool((cfg.get("openrouter_api_key") or "").strip())
