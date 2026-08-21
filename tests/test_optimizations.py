import sqlite3
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app_config
import db_utils
import email_classifier
from structured_extract import detectar_tipo_por_paginas


def test_modelo_principal_usa_qwen_nitro():
    assert app_config.PRIMARY_MODEL == "qwen/qwen3.7-flash:nitro"


def test_detectar_tipo_por_paginas_prefers_text_when_majority_pages_have_text():
    pages = [
        {"texto": "texto da pagina 1"},
        {"texto": "texto da pagina 2"},
        {"texto": ""},
        {"texto": ""},
    ]
    assert detectar_tipo_por_paginas(pages) == "pdf_texto"


def test_detectar_tipo_por_paginas_uses_scanned_when_text_is_sparse():
    pages = [
        {"texto": ""},
        {"texto": ""},
        {"texto": ""},
    ]
    assert detectar_tipo_por_paginas(pages) == "pdf_escaneado"


def test_classificar_emails_em_lote_aplica_guardas_de_duvida():
    parseds = [
        {"assunto": "Dúvida sobre item 10?", "corpo": "R$ 120,00", "remetente_email": "fornecedor@teste.com"},
        {"assunto": "Pedido", "corpo": "Preciso confirmar item 5", "remetente_email": "fornecedor@teste.com"},
    ]

    def fake_api(_payload):
        return {"choices": [{"message": {"content": '{"classificacoes": [{"id": 1, "tipo": "duvida", "confianca": 90, "resumo": "Pergunta", "numero_processo": null}]}'}}]}

    # monkeypatch is used directly to exercise batch output without needing a real OpenRouter call.
    monkeypatch = pytest.MonkeyPatch()
    monkeypatch.setattr(email_classifier, "_classificar_email_via_api", fake_api)
    try:
        result = email_classifier.classificar_emails_em_lote(parseds, api_key="x", model="dummy")
    finally:
        monkeypatch.undo()

    assert result[0]["tipo"] == "orcamento_recebido"
    assert result[1]["tipo"] == "outro"


def test_db_utils_supports_content_sha256_cache_lookup(tmp_path):
    db = sqlite3.connect(tmp_path / "cache.sqlite")
    try:
        db_utils._init_db(db)
        db.execute(
            "INSERT INTO arquivos (file_id, nome, empresa, cnpj, telefone, modified_time, extraction_version, processado_em, content_sha256) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            ("abc", "a.pdf", "Acme", "", "", "t", db_utils.EXTRACTION_VERSION, "2024-01-01T00:00:00Z", "hash-123"),
        )
        db.commit()
        cached = db_utils.get_cached_file(db, "abc")
        assert cached["content_sha256"] == "hash-123"
        assert db_utils.get_cached_file(db, "missing", content_sha256="hash-123")["file_id"] == "abc"
    finally:
        db.close()
