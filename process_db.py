"""
Banco de dados SQLite para processos, emails e fornecedores.
Schema separado do banco de orçamentos (orcamentos.db).
"""
from __future__ import annotations

import json
import re
import sqlite3
import unicodedata
from datetime import datetime, timezone
from typing import Optional


GENERIC_EMAIL_DOMAINS = {
    "gmail.com", "hotmail.com", "outlook.com", "live.com", "msn.com",
    "yahoo.com", "yahoo.com.br", "icloud.com", "me.com", "aol.com",
    "bol.com.br", "uol.com.br", "terra.com.br", "globo.com", "ig.com.br",
    "proton.me", "protonmail.com",
}

COMMON_DOMAIN_SUFFIXES = {
    "com", "net", "org", "gov", "edu", "mil", "co", "ac", "nom",
    "br", "uk", "ar", "cl", "uy", "py", "mx", "es", "pt", "us",
}

_REGEX_CNPJ = re.compile(r"(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\b\d{14}\b)")
_REGEX_TELEFONE = re.compile(r"(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?\d{4,5}[\s.-]?\d{4}")


def _validar_cnpj_local(cnpj_digits: str) -> bool:
    """Valida CNPJ onde `cnpj_digits` é apenas dígitos (14 caracteres)."""
    if not cnpj_digits or len(cnpj_digits) != 14:
        return False
    if cnpj_digits == cnpj_digits[0] * 14:
        return False

    def _calc(digs: str, mults: list[int]) -> int:
        s = sum(int(a) * b for a, b in zip(digs, mults))
        r = s % 11
        return 0 if r < 2 else 11 - r

    mult1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    mult2 = [6] + mult1
    d1 = _calc(cnpj_digits[:12], mult1)
    d2 = _calc(cnpj_digits[:12] + str(d1), mult2)
    return cnpj_digits.endswith(f"{d1}{d2}")


def _humanize_token(token: str) -> str:
    token = (token or "").strip().replace("_", " ").replace("-", " ")
    token = " ".join(part for part in token.split() if part)
    if not token:
        return ""
    return " ".join(p.capitalize() for p in token.split())


def _infer_company_from_email(email_addr: str) -> str:
    email_addr = (email_addr or "").strip().lower()
    if "@" not in email_addr:
        return ""
    domain = email_addr.split("@", 1)[1].strip(". ")
    if not domain or domain in GENERIC_EMAIL_DOMAINS:
        return ""

    parts = [p for p in domain.split(".") if p]
    if len(parts) < 2:
        return ""

    # Ex.: mail.nomedaempresa.com.br -> nomedaempresa
    idx = len(parts) - 2
    if parts[-1] in COMMON_DOMAIN_SUFFIXES and parts[-2] in COMMON_DOMAIN_SUFFIXES and len(parts) >= 3:
        idx = len(parts) - 3

    core = parts[idx]
    if not any(ch.isalpha() for ch in core):
        return ""
    return _humanize_token(core)


def nome_fornecedor_preferencial(email_addr: str, nome: str = "") -> str:
    """Prioriza nome da empresa pelo domínio do e-mail; fallback para nome informado."""
    by_domain = _infer_company_from_email(email_addr)
    if by_domain:
        return by_domain
    nome = (nome or "").strip()
    if nome:
        return _humanize_token(nome)
    if "@" in (email_addr or ""):
        local = email_addr.split("@", 1)[0]
        return _humanize_token(local)
    return ""


def _extrair_cnpj_do_texto(texto: str) -> str:
    if not texto:
        return ""

    # procura por rótulos explícitos primeiro (ex.: "CNPJ: 00.000.000/0000-00")
    m = re.search(r"(?i)\bcnpj\b[^\d]{0,12}([\d\.\-/]{10,30})", texto)
    if m:
        cand = re.sub(r"\D", "", m.group(1) or "")
        if len(cand) >= 14:
            cand = cand[-14:]
        if len(cand) == 14:
            if _validar_cnpj_local(cand):
                return f"{cand[:2]}.{cand[2:5]}.{cand[5:8]}/{cand[8:12]}-{cand[12:]}"

    # busca com regex original e valida (prioriza formatos escritos corretamente)
    m2 = _REGEX_CNPJ.search(texto)
    if m2:
        dig = re.sub(r"\D", "", m2.group(0))
        if len(dig) == 14 and _validar_cnpj_local(dig):
            return f"{dig[:2]}.{dig[2:5]}.{dig[5:8]}/{dig[8:12]}-{dig[12:]}"

    # procura por quaisquer sequências de dígitos longas que possam conter CNPJ
    # tenta extrair qualquer substring de 14 dígitos e validar checksum
    digits = re.sub(r"\D", "", texto)
    for i in range(0, max(0, len(digits) - 13)):
        sub = digits[i : i + 14]
        if len(sub) == 14 and _validar_cnpj_local(sub):
            return f"{sub[:2]}.{sub[2:5]}.{sub[5:8]}/{sub[8:12]}-{sub[12:]}"

    return ""


def _extrair_telefone_do_texto(texto: str) -> str:
    if not texto:
        return ""

    # DDDs válidos do Brasil
    VALID_DDDS = {
        '11','12','13','14','15','16','17','18','19',
        '21','22','24','27','28',
        '31','32','33','34','35','37','38',
        '41','42','43','44','45','46',
        '47','48','49',
        '51','53','54','55',
        '61','62','64','63','65','66','67',
        '68','69','71','73','74','75','77','79','81','82','83','84','85','86','87','88','89','91','92','93','94','95','96','97','98','99'
    }

    def normalize(num: str) -> str:
        d = re.sub(r"\D", "", num)
        if d.startswith('55') and len(d) in (12, 13):
            d = d[2:]
        return d

    # procura por rótulos explícitos (prioriza números próximos a palavras-chave)
    for m in re.finditer(r"(?i)(telefone|fone|tel|contato)[^\d]{0,12}([\d\(\)\s\+\-\.]{6,30}\d)", texto):
        cand = normalize(m.group(2) or "")
        if len(cand) in (10, 11) and cand[:2] in VALID_DDDS:
            ddd = cand[:2]
            num = cand[2:]
            if len(num) == 9:
                return f"({ddd}) {num[:5]}-{num[5:]}"
            return f"({ddd}) {num[:4]}-{num[4:]}"

    # coleta todas as sequências de dígitos longas e tenta encontrar janelas válidas
    digit_seqs = re.findall(r"\d{6,}", texto)
    candidates = []
    for seq in digit_seqs:
        s = seq
        # tenta janelas 10 e 11
        for L in (11, 10):
            if len(s) >= L:
                for i in range(0, len(s) - L + 1):
                    w = s[i : i + L]
                    norm = normalize(w)
                    if len(norm) in (10, 11) and norm[:2] in VALID_DDDS:
                        score = 0
                        # aumenta score se próximo a palavra-chave
                        idx = texto.find(w)
                        if idx >= 0:
                            window = texto[max(0, idx - 40) : idx + len(w) + 40].lower()
                            if any(k in window for k in ('telefone', 'fone', 'contato', 'cel', 'whatsapp', 'tel')):
                                score += 10
                        candidates.append((score, norm))

    # ordena por score e escolhe o melhor
    if candidates:
        candidates.sort(key=lambda x: (-x[0]))
        best = candidates[0][1]
        ddd = best[:2]
        num = best[2:]
        if len(num) == 9:
            return f"({ddd}) {num[:5]}-{num[5:]}"
        return f"({ddd}) {num[:4]}-{num[4:]}"

    # fallback: uso regex original
    for m in _REGEX_TELEFONE.finditer(texto):
        dig = re.sub(r"\D", "", m.group(0))
        if len(dig) == 13 and dig.startswith("55"):
            dig = dig[2:]
        if len(dig) not in (10, 11):
            continue
        ddd = dig[:2]
        num = dig[2:]
        if len(num) == 9:
            return f"({ddd}) {num[:5]}-{num[5:]}"
        return f"({ddd}) {num[:4]}-{num[4:]}"

    return ""


# ---------------------------------------------------------------------------
# Conexão e schema
# ---------------------------------------------------------------------------

def get_connection(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    _init_db(conn)
    return conn


def _init_db(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS processos (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            numero      TEXT    UNIQUE NOT NULL,
            titulo      TEXT,
            descricao   TEXT,
            criado_em   TEXT    NOT NULL
        );

        CREATE TABLE IF NOT EXISTS fornecedores (
            id    INTEGER PRIMARY KEY AUTOINCREMENT,
            nome  TEXT,
            email TEXT UNIQUE NOT NULL,
            cnpj  TEXT,
            telefone TEXT
        );

        CREATE TABLE IF NOT EXISTS emails (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            processo_id      INTEGER NOT NULL,
            message_id       TEXT,
            assunto          TEXT,
            remetente_nome   TEXT,
            remetente_email  TEXT,
            destinatarios    TEXT,   -- JSON: [{nome, email}, ...]
            data_envio       TEXT,
            tipo             TEXT,   -- categoria classificada
            confianca_tipo   INTEGER DEFAULT 0,
            resumo           TEXT,
            corpo            TEXT,
            tem_anexo        INTEGER DEFAULT 0,
            nomes_anexos     TEXT,   -- JSON: [nome1, nome2, ...]
            arquivo_eml_nome TEXT,
            classificado_em  TEXT,
            FOREIGN KEY(processo_id) REFERENCES processos(id)
        );

        CREATE TABLE IF NOT EXISTS email_anexos (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            email_id  INTEGER NOT NULL,
            nome      TEXT,
            conteudo  BLOB,
            tipo_mime TEXT,
            FOREIGN KEY(email_id) REFERENCES emails(id)
        );

        CREATE TABLE IF NOT EXISTS participacoes (
            id                     INTEGER PRIMARY KEY AUTOINCREMENT,
            processo_id            INTEGER NOT NULL,
            fornecedor_id          INTEGER NOT NULL,
            enviou_orcamento       INTEGER DEFAULT 0,
            recusou                INTEGER DEFAULT 0,
            confirmou_leitura      INTEGER DEFAULT 0,
            fez_pergunta           INTEGER DEFAULT 0,
            data_pedido_enviado    TEXT,
            data_primeira_resposta TEXT,
            UNIQUE(processo_id, fornecedor_id),
            FOREIGN KEY(processo_id)  REFERENCES processos(id),
            FOREIGN KEY(fornecedor_id) REFERENCES fornecedores(id)
        );

        CREATE TABLE IF NOT EXISTS processo_orcamentos (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            processo_id    INTEGER NOT NULL,
            file_id        TEXT NOT NULL,
            nome_arquivo   TEXT,
            criado_em      TEXT NOT NULL,
            UNIQUE(processo_id, file_id),
            FOREIGN KEY(processo_id) REFERENCES processos(id)
        );

        CREATE TABLE IF NOT EXISTS processo_consumo (
            id                     INTEGER PRIMARY KEY AUTOINCREMENT,
            processo_id             INTEGER UNIQUE NOT NULL,
            tokens_total            INTEGER DEFAULT 0,
            custo_total             REAL DEFAULT 0,
            prompt_tokens_total     INTEGER DEFAULT 0,
            completion_tokens_total INTEGER DEFAULT 0,
            tokens_emails_total     INTEGER DEFAULT 0,
            tokens_orcamentos_total INTEGER DEFAULT 0,
            executions_count        INTEGER DEFAULT 0,
            atualizado_em           TEXT,
            FOREIGN KEY(processo_id) REFERENCES processos(id)
        );
    """)
    # migrações incrementais
    _migrate(conn)
    conn.commit()


def _migrate(conn: sqlite3.Connection) -> None:
    """Adiciona colunas novas sem recriar tabelas."""
    emails_cols = {r[1] for r in conn.execute("PRAGMA table_info(emails)").fetchall()}
    if "confianca_tipo" not in emails_cols:
        conn.execute("ALTER TABLE emails ADD COLUMN confianca_tipo INTEGER DEFAULT 0")
    if "resumo" not in emails_cols:
        conn.execute("ALTER TABLE emails ADD COLUMN resumo TEXT")

    fornecedores_cols = {r[1] for r in conn.execute("PRAGMA table_info(fornecedores)").fetchall()}
    if "telefone" not in fornecedores_cols:
        conn.execute("ALTER TABLE fornecedores ADD COLUMN telefone TEXT")

    orcamentos_cols = {r[1] for r in conn.execute("PRAGMA table_info(processo_orcamentos)").fetchall()}
    if "remetente_email" not in orcamentos_cols:
        conn.execute("ALTER TABLE processo_orcamentos ADD COLUMN remetente_email TEXT")


# ---------------------------------------------------------------------------
# Processos
# ---------------------------------------------------------------------------

def listar_processos(conn: sqlite3.Connection) -> list[dict]:
    rows = conn.execute(
        "SELECT id, numero, titulo, descricao, criado_em FROM processos ORDER BY criado_em DESC"
    ).fetchall()
    return [dict(r) for r in rows]


def buscar_processo_por_numero(conn: sqlite3.Connection, numero: str) -> Optional[dict]:
    row = conn.execute(
        "SELECT id, numero, titulo, descricao, criado_em FROM processos WHERE numero = ?", (numero,)
    ).fetchone()
    return dict(row) if row else None


def criar_processo(conn: sqlite3.Connection, numero: str, titulo: str = "", descricao: str = "") -> int:
    """Cria novo processo e retorna o id."""
    agora = datetime.now(timezone.utc).isoformat()
    cur = conn.execute(
        "INSERT INTO processos (numero, titulo, descricao, criado_em) VALUES (?, ?, ?, ?)",
        (numero, titulo, descricao, agora),
    )
    conn.commit()
    return cur.lastrowid


def atualizar_processo(conn: sqlite3.Connection, processo_id: int, titulo: str = "", descricao: str = "") -> None:
    conn.execute(
        "UPDATE processos SET titulo = ?, descricao = ? WHERE id = ?",
        (titulo, descricao, processo_id),
    )
    conn.commit()


def deletar_processo(conn: sqlite3.Connection, processo_id: int) -> None:
    """Remove processo e todos os dados associados."""
    email_ids = [
        r[0] for r in conn.execute(
            "SELECT id FROM emails WHERE processo_id = ?", (processo_id,)
        ).fetchall()
    ]
    for eid in email_ids:
        conn.execute("DELETE FROM email_anexos WHERE email_id = ?", (eid,))
    conn.execute("DELETE FROM emails WHERE processo_id = ?", (processo_id,))
    conn.execute("DELETE FROM participacoes WHERE processo_id = ?", (processo_id,))
    conn.execute("DELETE FROM processo_orcamentos WHERE processo_id = ?", (processo_id,))
    conn.execute("DELETE FROM processo_consumo WHERE processo_id = ?", (processo_id,))
    conn.execute("DELETE FROM processos WHERE id = ?", (processo_id,))
    conn.commit()


def vincular_orcamento_ao_processo(
    conn: sqlite3.Connection,
    processo_id: int,
    file_id: str,
    nome_arquivo: str,
    remetente_email: str = "",
) -> None:
    agora = datetime.now(timezone.utc).isoformat()
    remetente_email = (remetente_email or "").strip().lower()
    conn.execute(
        """
        INSERT INTO processo_orcamentos (processo_id, file_id, nome_arquivo, remetente_email, criado_em)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(processo_id, file_id) DO UPDATE SET
            remetente_email = CASE
                WHEN excluded.remetente_email != '' THEN excluded.remetente_email
                ELSE remetente_email
            END
        """,
        (processo_id, file_id, nome_arquivo, remetente_email, agora),
    )
    conn.commit()


def listar_orcamentos_do_processo(conn: sqlite3.Connection, processo_id: int) -> list[dict]:
    rows = conn.execute(
        "SELECT file_id, nome_arquivo, remetente_email, criado_em FROM processo_orcamentos WHERE processo_id = ? ORDER BY criado_em ASC, id ASC",
        (processo_id,),
    ).fetchall()
    return [dict(r) for r in rows]


def backfill_remetente_email_orcamentos(conn: sqlite3.Connection, processo_id: Optional[int] = None) -> int:
    """Preenche remetente_email em vínculos antigos de anexos de e-mail.

    Para file_id no formato email_attachment:<email_id>:..., busca o remetente
    em emails.id e grava em processo_orcamentos.remetente_email.
    """
    sql = """
        SELECT id, processo_id, file_id
        FROM processo_orcamentos
        WHERE (remetente_email IS NULL OR TRIM(remetente_email) = '')
          AND file_id LIKE 'email_attachment:%'
    """
    params: list = []
    if processo_id is not None:
        sql += " AND processo_id = ?"
        params.append(processo_id)

    rows = conn.execute(sql, tuple(params)).fetchall()
    atualizados = 0
    for r in rows:
        file_id = str(r["file_id"] or "")
        m = re.match(r"^email_attachment:(\d+):", file_id)
        if not m:
            continue
        email_id = int(m.group(1))
        row_email = conn.execute(
            "SELECT remetente_email FROM emails WHERE id = ? AND processo_id = ?",
            (email_id, r["processo_id"]),
        ).fetchone()
        remetente = (row_email["remetente_email"] if row_email else "") or ""
        remetente = remetente.strip().lower()
        if not remetente:
            continue

        conn.execute(
            "UPDATE processo_orcamentos SET remetente_email = ? WHERE id = ?",
            (remetente, r["id"]),
        )
        atualizados += 1

    if atualizados:
        conn.commit()
    return atualizados


def atualizar_cnpjs_fornecedores(conn: sqlite3.Connection, processo_id: Optional[int] = None) -> int:
    """Preenche CNPJ/telefone de fornecedores usando corpo de e-mails já armazenados."""
    sql = """
        SELECT DISTINCT f.id AS fornecedor_id, e.corpo
        FROM fornecedores f
        JOIN emails e ON LOWER(e.remetente_email) = LOWER(f.email)
        LEFT JOIN participacoes p ON p.fornecedor_id = f.id
                WHERE ((f.cnpj IS NULL OR TRIM(f.cnpj) = '') OR (f.telefone IS NULL OR TRIM(f.telefone) = ''))
          AND (e.corpo IS NOT NULL AND TRIM(e.corpo) <> '')
    """
    params = []
    if processo_id is not None:
        sql += " AND p.processo_id = ?"
        params.append(processo_id)
    sql += " ORDER BY e.id DESC"

    rows = conn.execute(sql, tuple(params)).fetchall()
    atualizados = 0
    vistos = set()
    for r in rows:
        fid = r["fornecedor_id"]
        if fid in vistos:
            continue
        cnpj = _extrair_cnpj_do_texto(r["corpo"])
        tel = _extrair_telefone_do_texto(r["corpo"])
        if not cnpj and not tel:
            continue

        atual = conn.execute("SELECT cnpj, telefone FROM fornecedores WHERE id = ?", (fid,)).fetchone()
        if not atual:
            continue
        updates = []
        params = []
        if cnpj and not (atual["cnpj"] or "").strip():
            updates.append("cnpj = ?")
            params.append(cnpj)
        if tel and not (atual["telefone"] or "").strip():
            updates.append("telefone = ?")
            params.append(tel)
        if not updates:
            continue

        params.append(fid)
        conn.execute(f"UPDATE fornecedores SET {', '.join(updates)} WHERE id = ?", tuple(params))
        vistos.add(fid)
        atualizados += 1
    if atualizados:
        conn.commit()
    return atualizados


def atualizar_dados_fornecedor_por_nome_no_processo(
    conn: sqlite3.Connection,
    processo_id: int,
    nome_referencia: str,
    cnpj: str = "",
    telefone: str = "",
) -> int:
    """Atualiza CNPJ/telefone de fornecedor do processo por aproximação de nome."""
    nome_ref = (nome_referencia or "").strip().lower()
    cnpj = (cnpj or "").strip()
    telefone = (telefone or "").strip()
    if not nome_ref or (not cnpj and not telefone):
        return 0

    rows = conn.execute(
        """
        SELECT f.id, f.nome, f.cnpj, f.telefone
        FROM participacoes p
        JOIN fornecedores f ON f.id = p.fornecedor_id
        WHERE p.processo_id = ?
        """,
        (processo_id,),
    ).fetchall()

    def _norm_nome(v: str) -> str:
        base = unicodedata.normalize("NFKD", (v or "").lower())
        base = "".join(ch for ch in base if not unicodedata.combining(ch))
        base = re.sub(r"[^a-z0-9]+", " ", base)
        return " ".join(base.split())

    stop = {
        "ltda", "eireli", "me", "epp", "sa", "s", "a", "comercio", "servicos",
        "de", "da", "do", "das", "dos", "empresa", "sociedade", "limitada",
    }

    def _tokens(v: str) -> set[str]:
        return {t for t in _norm_nome(v).split() if len(t) >= 3 and t not in stop}

    nome_ref_n = _norm_nome(nome_ref)
    ref_tokens = _tokens(nome_ref)
    atualizados = 0
    for r in rows:
        nome_forn = _norm_nome(r["nome"] or "")
        if not nome_forn:
            continue
        forn_tokens = _tokens(r["nome"] or "")
        inter = len(ref_tokens.intersection(forn_tokens)) if ref_tokens and forn_tokens else 0
        ratio_ref = inter / max(1, len(ref_tokens)) if ref_tokens else 0.0
        ratio_forn = inter / max(1, len(forn_tokens)) if forn_tokens else 0.0

        casou = (
            nome_ref_n in nome_forn
            or nome_forn in nome_ref_n
            or ratio_ref >= 0.6
            or ratio_forn >= 0.6
            or inter >= 2
        )
        if not casou:
            continue

        updates = []
        params = []
        if cnpj and not (r["cnpj"] or "").strip():
            updates.append("cnpj = ?")
            params.append(cnpj)
        if telefone and not (r["telefone"] or "").strip():
            updates.append("telefone = ?")
            params.append(telefone)
        if not updates:
            continue

        params.append(r["id"])
        conn.execute(f"UPDATE fornecedores SET {', '.join(updates)} WHERE id = ?", tuple(params))
        atualizados += 1

    if atualizados:
        conn.commit()
    return atualizados


# ---------------------------------------------------------------------------
# Fornecedores
# ---------------------------------------------------------------------------

def upsert_fornecedor(
    conn: sqlite3.Connection,
    email_addr: str,
    nome: str = "",
    cnpj: str = "",
    telefone: str = "",
) -> int:
    """Cria ou encontra fornecedor pelo email. Retorna id."""
    email_addr = email_addr.lower().strip()
    nome_preferencial = nome_fornecedor_preferencial(email_addr, nome)
    cnpj = (cnpj or "").strip()
    telefone = (telefone or "").strip()
    row = conn.execute(
        "SELECT id, nome, cnpj, telefone FROM fornecedores WHERE email = ?", (email_addr,)
    ).fetchone()
    if row:
        fid = row["id"]
        update_cols = []
        params = []

        # Atualiza sempre que houver nome preferencial mais consistente.
        if nome_preferencial and nome_preferencial != (row["nome"] or ""):
            update_cols.append("nome = ?")
            params.append(nome_preferencial)

        # Preenche CNPJ quando ainda estiver vazio no cadastro.
        if cnpj and not (row["cnpj"] or "").strip():
            update_cols.append("cnpj = ?")
            params.append(cnpj)

        # Preenche telefone quando ainda estiver vazio no cadastro.
        if telefone and not (row["telefone"] or "").strip():
            update_cols.append("telefone = ?")
            params.append(telefone)

        if update_cols:
            params.append(fid)
            conn.execute(f"UPDATE fornecedores SET {', '.join(update_cols)} WHERE id = ?", tuple(params))
            conn.commit()
        return fid
    cur = conn.execute(
        "INSERT INTO fornecedores (email, nome, cnpj, telefone) VALUES (?, ?, ?, ?)",
        (email_addr, nome_preferencial, cnpj or "", telefone or ""),
    )
    conn.commit()
    return cur.lastrowid


def listar_fornecedores(conn: sqlite3.Connection) -> list[dict]:
    """Lista fornecedores com estatísticas agregadas."""
    rows = conn.execute("""
        SELECT
            f.id,
            f.nome,
            f.email,
            f.cnpj,
            f.telefone,
            COUNT(DISTINCT p.processo_id)                           AS total_processos,
            SUM(p.enviou_orcamento)                                 AS total_orcamentos,
            SUM(p.recusou)                                          AS total_recusas,
            SUM(p.confirmou_leitura)                                AS total_leituras
        FROM fornecedores f
        LEFT JOIN participacoes p ON p.fornecedor_id = f.id
        GROUP BY f.id
        ORDER BY f.nome COLLATE NOCASE
    """).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d["nome"] = nome_fornecedor_preferencial(d.get("email", ""), d.get("nome", ""))
        result.append(d)
    return result


def listar_fornecedores_processo(conn: sqlite3.Connection, processo_id: int) -> list[dict]:
    """Lista fornecedores vinculados a um processo com suas métricas no processo."""
    rows = conn.execute(
        """
        SELECT
            f.id,
            f.nome,
            f.email,
            f.cnpj,
            f.telefone,
            p.enviou_orcamento,
            p.recusou,
            p.confirmou_leitura,
            p.fez_pergunta,
            p.data_pedido_enviado,
            p.data_primeira_resposta
        FROM participacoes p
        JOIN fornecedores f ON f.id = p.fornecedor_id
        WHERE p.processo_id = ?
        ORDER BY f.nome COLLATE NOCASE
        """,
        (processo_id,),
    ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d["nome"] = nome_fornecedor_preferencial(d.get("email", ""), d.get("nome", ""))
        result.append(d)
    return result


def clear_all(conn: sqlite3.Connection) -> None:
    """Limpa todo o histórico de processos, e-mails, anexos e fornecedores."""
    conn.execute("DELETE FROM email_anexos")
    conn.execute("DELETE FROM emails")
    conn.execute("DELETE FROM participacoes")
    conn.execute("DELETE FROM processo_orcamentos")
    conn.execute("DELETE FROM processo_consumo")
    conn.execute("DELETE FROM processos")
    conn.execute("DELETE FROM fornecedores")
    conn.commit()


# ---------------------------------------------------------------------------
# Emails
# ---------------------------------------------------------------------------

def email_ja_importado(conn: sqlite3.Connection, processo_id: int, message_id: str) -> bool:
    """Evita duplicatas: verifica se o message_id já foi salvo para este processo."""
    row = conn.execute(
        "SELECT 1 FROM emails WHERE processo_id = ? AND message_id = ?",
        (processo_id, message_id),
    ).fetchone()
    return row is not None


def salvar_email(
    conn: sqlite3.Connection,
    processo_id: int,
    parsed: dict,
    tipo: str,
    confianca: int = 0,
    resumo: str = "",
) -> int:
    """Persiste um email classificado. Retorna o id inserido."""
    agora = datetime.now(timezone.utc).isoformat()
    cur = conn.execute(
        """
        INSERT INTO emails (
            processo_id, message_id, assunto, remetente_nome, remetente_email,
            destinatarios, data_envio, tipo, confianca_tipo, resumo, corpo,
            tem_anexo, nomes_anexos, arquivo_eml_nome, classificado_em
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            processo_id,
            parsed.get("message_id"),
            parsed.get("assunto"),
            parsed.get("remetente_nome"),
            parsed.get("remetente_email"),
            json.dumps(parsed.get("destinatarios", []), ensure_ascii=False),
            parsed.get("data_envio"),
            tipo,
            confianca,
            resumo,
            parsed.get("corpo"),
            1 if parsed.get("tem_anexo") else 0,
            json.dumps(parsed.get("nomes_anexos", []), ensure_ascii=False),
            parsed.get("arquivo_eml_nome"),
            agora,
        ),
    )
    conn.commit()
    return cur.lastrowid


def salvar_anexo(
    conn: sqlite3.Connection, email_id: int, nome: str, conteudo: bytes, tipo_mime: str
) -> int:
    cur = conn.execute(
        "INSERT INTO email_anexos (email_id, nome, conteudo, tipo_mime) VALUES (?, ?, ?, ?)",
        (email_id, nome, conteudo, tipo_mime),
    )
    conn.commit()
    return cur.lastrowid


def listar_emails_processo(conn: sqlite3.Connection, processo_id: int) -> list[dict]:
    rows = conn.execute(
        """
        SELECT id, message_id, assunto, remetente_nome, remetente_email,
               destinatarios, data_envio, tipo, confianca_tipo, resumo,
               tem_anexo, nomes_anexos, arquivo_eml_nome, classificado_em
        FROM emails
        WHERE processo_id = ?
        ORDER BY data_envio ASC, id ASC
        """,
        (processo_id,),
    ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d["destinatarios"] = json.loads(d["destinatarios"] or "[]")
        d["nomes_anexos"] = json.loads(d["nomes_anexos"] or "[]")
        result.append(d)
    return result


def get_email_corpo(conn: sqlite3.Connection, email_id: int) -> str:
    row = conn.execute("SELECT corpo FROM emails WHERE id = ?", (email_id,)).fetchone()
    return row["corpo"] if row else ""


def listar_anexos_email(conn: sqlite3.Connection, email_id: int) -> list[dict]:
    rows = conn.execute(
        "SELECT id, nome, tipo_mime FROM email_anexos WHERE email_id = ?", (email_id,)
    ).fetchall()
    return [dict(r) for r in rows]


def get_anexo_conteudo(conn: sqlite3.Connection, anexo_id: int) -> Optional[bytes]:
    row = conn.execute("SELECT conteudo FROM email_anexos WHERE id = ?", (anexo_id,)).fetchone()
    return row["conteudo"] if row else None


# ---------------------------------------------------------------------------
# Participações (estatísticas por fornecedor x processo)
# ---------------------------------------------------------------------------

def atualizar_participacao(
    conn: sqlite3.Connection,
    processo_id: int,
    fornecedor_id: int,
    tipo_email: str,
    data_envio: Optional[str] = None,
    data_pedido_inferida: Optional[str] = None,
) -> None:
    """Atualiza os flags de participação de um fornecedor em um processo."""
    # Garante que o registro existe
    conn.execute(
        "INSERT OR IGNORE INTO participacoes (processo_id, fornecedor_id) VALUES (?, ?)",
        (processo_id, fornecedor_id),
    )
    if tipo_email == "pedido_orcamento" and data_envio:
        conn.execute(
            """UPDATE participacoes
               SET data_pedido_enviado = COALESCE(data_pedido_enviado, ?)
               WHERE processo_id = ? AND fornecedor_id = ?""",
            (data_envio, processo_id, fornecedor_id),
        )
    elif tipo_email == "orcamento_recebido":
        conn.execute(
            """UPDATE participacoes
               SET enviou_orcamento = 1,
                   data_primeira_resposta = COALESCE(data_primeira_resposta, ?),
                   data_pedido_enviado = COALESCE(data_pedido_enviado, ?)
               WHERE processo_id = ? AND fornecedor_id = ?""",
            (data_envio, data_pedido_inferida, processo_id, fornecedor_id),
        )
    elif tipo_email == "declinio":
        conn.execute(
            """UPDATE participacoes
               SET recusou = 1,
                   data_primeira_resposta = COALESCE(data_primeira_resposta, ?),
                   data_pedido_enviado = COALESCE(data_pedido_enviado, ?)
               WHERE processo_id = ? AND fornecedor_id = ?""",
            (data_envio, data_pedido_inferida, processo_id, fornecedor_id),
        )
    elif tipo_email == "confirmacao_leitura":
        conn.execute(
            """UPDATE participacoes
               SET confirmou_leitura = 1
               WHERE processo_id = ? AND fornecedor_id = ?""",
            (processo_id, fornecedor_id),
        )
    elif tipo_email == "duvida":
        conn.execute(
            """UPDATE participacoes
               SET fez_pergunta = 1,
                   data_primeira_resposta = COALESCE(data_primeira_resposta, ?),
                   data_pedido_enviado = COALESCE(data_pedido_enviado, ?)
               WHERE processo_id = ? AND fornecedor_id = ?""",
            (data_envio, data_pedido_inferida, processo_id, fornecedor_id),
        )
    conn.commit()


def get_participacoes_processo(conn: sqlite3.Connection, processo_id: int) -> list[dict]:
    rows = conn.execute(
        """
        SELECT
            f.id         AS fornecedor_id,
            f.nome,
            f.email,
            p.enviou_orcamento,
            p.recusou,
            p.confirmou_leitura,
            p.fez_pergunta,
            p.data_pedido_enviado,
            p.data_primeira_resposta
        FROM participacoes p
        JOIN fornecedores f ON f.id = p.fornecedor_id
        WHERE p.processo_id = ?
        ORDER BY f.nome COLLATE NOCASE
        """,
        (processo_id,),
    ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d["nome"] = nome_fornecedor_preferencial(d.get("email", ""), d.get("nome", ""))
        result.append(d)
    return result


def registrar_consumo_processo(
    conn: sqlite3.Connection,
    processo_id: int,
    tokens_total: int,
    custo_total: float,
    prompt_tokens_total: int = 0,
    completion_tokens_total: int = 0,
    tokens_emails_total: int = 0,
    tokens_orcamentos_total: int = 0,
) -> None:
    agora = datetime.now(timezone.utc).isoformat()
    conn.execute(
        """
        INSERT INTO processo_consumo (
            processo_id, tokens_total, custo_total,
            prompt_tokens_total, completion_tokens_total,
            tokens_emails_total, tokens_orcamentos_total,
            executions_count, atualizado_em
        ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)
        ON CONFLICT(processo_id) DO UPDATE SET
            tokens_total = tokens_total + excluded.tokens_total,
            custo_total = custo_total + excluded.custo_total,
            prompt_tokens_total = prompt_tokens_total + excluded.prompt_tokens_total,
            completion_tokens_total = completion_tokens_total + excluded.completion_tokens_total,
            tokens_emails_total = tokens_emails_total + excluded.tokens_emails_total,
            tokens_orcamentos_total = tokens_orcamentos_total + excluded.tokens_orcamentos_total,
            executions_count = executions_count + 1,
            atualizado_em = excluded.atualizado_em
        """,
        (
            processo_id,
            int(tokens_total or 0),
            float(custo_total or 0.0),
            int(prompt_tokens_total or 0),
            int(completion_tokens_total or 0),
            int(tokens_emails_total or 0),
            int(tokens_orcamentos_total or 0),
            agora,
        ),
    )
    conn.commit()


def obter_consumo_processo(conn: sqlite3.Connection, processo_id: int) -> dict:
    row = conn.execute(
        """
        SELECT tokens_total, custo_total, prompt_tokens_total, completion_tokens_total,
               tokens_emails_total, tokens_orcamentos_total, executions_count, atualizado_em
        FROM processo_consumo
        WHERE processo_id = ?
        """,
        (processo_id,),
    ).fetchone()
    if not row:
        return {
            "tokens_total": 0,
            "custo_total": 0.0,
            "prompt_tokens_total": 0,
            "completion_tokens_total": 0,
            "tokens_emails_total": 0,
            "tokens_orcamentos_total": 0,
            "executions_count": 0,
            "atualizado_em": None,
        }
    return dict(row)


# ---------------------------------------------------------------------------
# Relatório resumido de um processo
# ---------------------------------------------------------------------------

def get_resumo_processo(conn: sqlite3.Connection, processo_id: int) -> dict:
    """Retorna estatísticas consolidadas de um processo."""
    emails = listar_emails_processo(conn, processo_id)
    participacoes = get_participacoes_processo(conn, processo_id)

    contagem = {}
    for e in emails:
        t = e.get("tipo") or "outro"
        contagem[t] = contagem.get(t, 0) + 1

    # Tempo médio de resposta (pedido -> primeira resposta do fornecedor)
    tempos_resposta = []

    def _parse_iso_para_delta(valor: str | None):
        if not valor:
            return None
        d = datetime.fromisoformat(valor)
        # Normaliza para datetime "naive" em UTC para evitar erro de aware vs naive.
        if d.tzinfo is not None:
            d = d.astimezone(timezone.utc).replace(tzinfo=None)
        return d

    for p in participacoes:
        if p.get("data_pedido_enviado") and p.get("data_primeira_resposta"):
            try:
                t_pedido = _parse_iso_para_delta(p["data_pedido_enviado"])
                t_resposta = _parse_iso_para_delta(p["data_primeira_resposta"])
                if not t_pedido or not t_resposta:
                    continue
                delta_h = (t_resposta - t_pedido).total_seconds() / 3600
                if delta_h >= 0:
                    tempos_resposta.append(delta_h)
            except Exception:
                pass

    tempo_medio_h = sum(tempos_resposta) / len(tempos_resposta) if tempos_resposta else None

    return {
        "total_emails": len(emails),
        "por_tipo": contagem,
        "total_fornecedores": len(participacoes),
        "enviaram_orcamento": sum(1 for p in participacoes if p["enviou_orcamento"]),
        "recusaram": sum(1 for p in participacoes if p["recusou"]),
        "confirmaram_leitura": sum(1 for p in participacoes if p["confirmou_leitura"]),
        "sem_resposta": sum(
            1 for p in participacoes
            if not p["enviou_orcamento"]
            and not p["recusou"]
            and not p["confirmou_leitura"]
            and not p["fez_pergunta"]
            and not p.get("data_primeira_resposta")
        ),
        "tempo_medio_resposta_h": tempo_medio_h,
        "participacoes": participacoes,
    }
