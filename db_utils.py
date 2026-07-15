"""
Persistência local em SQLite: guarda os itens já extraídos de cada arquivo,
para não reprocessar (e não gastar API de novo) arquivos que não mudaram
desde a última execução.
"""
import sqlite3
from datetime import datetime, timezone


EXTRACTION_VERSION = "5"


def get_connection(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, check_same_thread=False)
    _init_db(conn)
    return conn


def _init_db(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS arquivos (
            file_id TEXT PRIMARY KEY,
            nome TEXT,
            empresa TEXT,
            cnpj TEXT,
            telefone TEXT,
            modified_time TEXT,
            extraction_version TEXT,
            processado_em TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id TEXT,
            numero_item TEXT,
            descricao TEXT,
            unidade TEXT,
            quantidade REAL,
            preco_unitario REAL,
            preco_total REAL,
            fonte_extracao TEXT,
            origem TEXT,
            FOREIGN KEY(file_id) REFERENCES arquivos(file_id)
        )
    """)
    colunas_arquivos = {
        row[1] for row in conn.execute("PRAGMA table_info(arquivos)").fetchall()
    }
    if "extraction_version" not in colunas_arquivos:
        conn.execute("ALTER TABLE arquivos ADD COLUMN extraction_version TEXT")
    if "cnpj" not in colunas_arquivos:
        conn.execute("ALTER TABLE arquivos ADD COLUMN cnpj TEXT")
    if "telefone" not in colunas_arquivos:
        conn.execute("ALTER TABLE arquivos ADD COLUMN telefone TEXT")

    colunas_itens = {
        row[1] for row in conn.execute("PRAGMA table_info(itens)").fetchall()
    }
    if "fonte_extracao" not in colunas_itens:
        conn.execute("ALTER TABLE itens ADD COLUMN fonte_extracao TEXT")
    if "origem" not in colunas_itens:
        conn.execute("ALTER TABLE itens ADD COLUMN origem TEXT")
    conn.commit()


def get_cached_file(conn: sqlite3.Connection, file_id: str):
    """Retorna {'modified_time', 'empresa', 'nome'} se o arquivo já foi processado antes, senão None."""
    cur = conn.execute(
        "SELECT modified_time, empresa, nome, extraction_version, cnpj, telefone FROM arquivos WHERE file_id = ?", (file_id,)
    )
    row = cur.fetchone()
    if row is None:
        return None
    return {
        "modified_time": row[0],
        "empresa": row[1],
        "nome": row[2],
        "extraction_version": row[3],
        "cnpj": row[4],
        "telefone": row[5],
    }


def get_items_for_file(conn: sqlite3.Connection, file_id: str):
    cur = conn.execute(
        "SELECT numero_item, descricao, unidade, quantidade, preco_unitario, preco_total, fonte_extracao, origem "
        "FROM itens WHERE file_id = ?",
        (file_id,),
    )
    return [
        {
            "numero_item": r[0],
            "descricao": r[1],
            "unidade": r[2],
            "quantidade": r[3],
            "preco_unitario": r[4],
            "preco_total": r[5],
            "fonte_extracao": r[6],
            "origem": r[7],
        }
        for r in cur.fetchall()
    ]


def save_extraction(conn: sqlite3.Connection, file_id: str, nome: str, empresa: str,
                     modified_time: str, itens: list, cnpj: str = "", telefone: str = "") -> None:
    conn.execute("DELETE FROM itens WHERE file_id = ?", (file_id,))
    conn.execute("DELETE FROM arquivos WHERE file_id = ?", (file_id,))
    conn.execute(
        "INSERT INTO arquivos (file_id, nome, empresa, cnpj, telefone, modified_time, extraction_version, processado_em) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            file_id,
            nome,
            empresa,
            cnpj or "",
            telefone or "",
            modified_time,
            EXTRACTION_VERSION,
            datetime.now(timezone.utc).isoformat(),
        ),
    )
    for item in itens:
        conn.execute(
            "INSERT INTO itens (file_id, numero_item, descricao, unidade, quantidade, "
            "preco_unitario, preco_total, fonte_extracao, origem) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                file_id,
                item.get("numero_item"),
                item.get("descricao"),
                item.get("unidade"),
                item.get("quantidade"),
                item.get("preco_unitario"),
                item.get("preco_total"),
                item.get("fonte_extracao"),
                item.get("origem"),
            ),
        )
    conn.commit()


def count_files(conn: sqlite3.Connection) -> int:
    cur = conn.execute(
        "SELECT COUNT(*) FROM arquivos WHERE COALESCE(extraction_version, '') = ?",
        (EXTRACTION_VERSION,),
    )
    return cur.fetchone()[0]


def purge_old_versions(conn: sqlite3.Connection) -> int:
    """Remove entradas de cache de versões antigas de extração."""
    cur = conn.execute(
        "SELECT file_id FROM arquivos WHERE COALESCE(extraction_version, '') != ?",
        (EXTRACTION_VERSION,),
    )
    antigos = [r[0] for r in cur.fetchall()]
    for file_id in antigos:
        conn.execute("DELETE FROM itens WHERE file_id = ?", (file_id,))
        conn.execute("DELETE FROM arquivos WHERE file_id = ?", (file_id,))
    conn.commit()
    return len(antigos)


def clear_db(conn: sqlite3.Connection) -> None:
    conn.execute("DELETE FROM itens")
    conn.execute("DELETE FROM arquivos")
    conn.commit()
