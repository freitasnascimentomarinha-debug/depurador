"""
Memória de correções de casamento (aprendizado incremental).

Cada vez que o usuário confirma ou rejeita um casamento fuzzy na interface,
a decisão é gravada aqui e aplicada automaticamente nos próximos lotes,
ANTES do fuzzy matching — exatamente como a skill registra lições de campo.

Chave da correção: par ordenado de descrições normalizadas. A normalização
deve ser idêntica à de match_utils.normalize_desc para as chaves baterem.
"""
from __future__ import annotations

import os
import sqlite3
from datetime import datetime

DECISOES_VALIDAS = {"casar", "nao_casar"}

_DEFAULT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aprendizado.db")


def normalize_desc(desc: str) -> str:
    """Mesma normalização de match_utils.normalize_desc."""
    if not desc:
        return ""
    return " ".join(desc.strip().lower().split())


def _par_chave(desc_a: str, desc_b: str) -> tuple[str, str]:
    a, b = normalize_desc(desc_a), normalize_desc(desc_b)
    return (a, b) if a <= b else (b, a)


def get_connection(path: str | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(path or _DEFAULT_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS correcoes_matching (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            desc_a      TEXT NOT NULL,
            desc_b      TEXT NOT NULL,
            decisao     TEXT NOT NULL CHECK (decisao IN ('casar', 'nao_casar')),
            origem      TEXT DEFAULT 'manual',
            criado_em   TEXT NOT NULL,
            UNIQUE(desc_a, desc_b)
        )
        """
    )
    conn.commit()
    return conn


def registrar_correcao(
    conn: sqlite3.Connection,
    desc_a: str,
    desc_b: str,
    decisao: str,
    origem: str = "manual",
) -> bool:
    """Grava (ou atualiza) uma decisão de casamento. Retorna True se gravou."""
    if decisao not in DECISOES_VALIDAS:
        return False
    a, b = _par_chave(desc_a, desc_b)
    if not a or not b or a == b:
        return False
    conn.execute(
        """
        INSERT INTO correcoes_matching (desc_a, desc_b, decisao, origem, criado_em)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(desc_a, desc_b) DO UPDATE SET
            decisao = excluded.decisao,
            origem = excluded.origem,
            criado_em = excluded.criado_em
        """,
        (a, b, decisao, origem, datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    return True


def carregar_correcoes(conn: sqlite3.Connection) -> dict[tuple[str, str], str]:
    """Retorna {(desc_a_norm, desc_b_norm): 'casar'|'nao_casar'} para uso no matching."""
    rows = conn.execute("SELECT desc_a, desc_b, decisao FROM correcoes_matching").fetchall()
    return {(r["desc_a"], r["desc_b"]): r["decisao"] for r in rows}


def contar_correcoes(conn: sqlite3.Connection) -> int:
    row = conn.execute("SELECT COUNT(*) AS n FROM correcoes_matching").fetchone()
    return int(row["n"]) if row else 0


def consultar_par(correcoes: dict, desc_a: str, desc_b: str) -> str | None:
    """Consulta uma decisão salva para o par (já com dict carregado em memória)."""
    return correcoes.get(_par_chave(desc_a, desc_b))
