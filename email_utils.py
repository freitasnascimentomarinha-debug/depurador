"""
Utilitários para processar arquivos .eml e extrair emails de ZIPs.
"""
from __future__ import annotations

import email
import email.policy
import hashlib
import io
import os
import re
import tarfile
import zipfile
from datetime import datetime
from email import message_from_bytes
from email.header import decode_header as _decode_header_raw
from email.utils import parsedate_to_datetime, getaddresses
from typing import Optional


SUPPORTED_BUDGET_EXTENSIONS = (
    ".pdf", ".docx", ".doc", ".xlsx", ".xls",
    ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp",
)


# ---------------------------------------------------------------------------
# Helpers de decodificação
# ---------------------------------------------------------------------------

def _decode_str(value: str) -> str:
    """Decodifica valores de cabeçalho RFC 2047."""
    if not value:
        return ""
    try:
        parts = _decode_header_raw(value)
        decoded = []
        for part, charset in parts:
            if isinstance(part, bytes):
                decoded.append(part.decode(charset or "utf-8", errors="replace"))
            else:
                decoded.append(str(part))
        return "".join(decoded).strip()
    except Exception:
        return str(value).strip()


def _parse_addr(raw: str) -> list[dict]:
    """Parseia endereços no formato 'Nome <email@x.com>'.
    Retorna lista de {'nome': ..., 'email': ...}.
    """
    decoded = _decode_str(raw or "")
    result = []
    for name, addr in getaddresses([decoded]):
        addr = addr.lower().strip()
        if addr and "@" in addr:
            result.append({"nome": name.strip().strip('"'), "email": addr})
    return result


# ---------------------------------------------------------------------------
# Extração de corpo e anexos
# ---------------------------------------------------------------------------

def _get_body(msg) -> str:
    """Extrai corpo de texto plano do email (prioriza text/plain)."""
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            disp = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in disp:
                charset = part.get_content_charset() or "utf-8"
                payload = part.get_payload(decode=True)
                if payload:
                    try:
                        parts.append(payload.decode(charset, errors="replace"))
                    except Exception:
                        parts.append(payload.decode("latin-1", errors="replace"))
        # fallback: text/html se não houver plain
        if not parts:
            for part in msg.walk():
                ct = part.get_content_type()
                disp = str(part.get("Content-Disposition", ""))
                if ct == "text/html" and "attachment" not in disp:
                    charset = part.get_content_charset() or "utf-8"
                    payload = part.get_payload(decode=True)
                    if payload:
                        try:
                            text = payload.decode(charset, errors="replace")
                        except Exception:
                            text = payload.decode("latin-1", errors="replace")
                        # remove tags HTML básicas para leitura
                        text = re.sub(r"<[^>]+>", " ", text)
                        text = re.sub(r"&nbsp;", " ", text)
                        text = re.sub(r"&amp;", "&", text)
                        text = re.sub(r"\s{3,}", "\n", text)
                        parts.append(text.strip())
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            try:
                parts.append(payload.decode(charset, errors="replace"))
            except Exception:
                parts.append(payload.decode("latin-1", errors="replace"))
    return "\n".join(parts).strip()


def _get_attachments(msg) -> list[dict]:
    """Retorna lista de anexos: {'nome', 'conteudo_bytes', 'tipo_mime'}."""
    anexos = []
    if not msg.is_multipart():
        return anexos
    for part in msg.walk():
        disp = str(part.get("Content-Disposition", ""))
        filename = part.get_filename()
        # inclui partes com filename ou com Content-Disposition: attachment
        if filename or "attachment" in disp:
            if filename:
                filename = _decode_str(filename)
            else:
                filename = f"anexo_{len(anexos) + 1}"
            payload = part.get_payload(decode=True)
            if payload:
                anexos.append({
                    "nome": filename,
                    "conteudo_bytes": payload,
                    "tipo_mime": part.get_content_type() or "application/octet-stream",
                })
    return anexos


# ---------------------------------------------------------------------------
# Parser principal de .eml
# ---------------------------------------------------------------------------

def parse_eml(eml_bytes: bytes, filename: str = "") -> dict:
    """
    Parseia bytes de um arquivo .eml.
    Retorna dict com todos os campos relevantes.
    """
    msg = message_from_bytes(eml_bytes, policy=email.policy.compat32)

    assunto = _decode_str(msg.get("Subject", ""))

    remetente_lista = _parse_addr(msg.get("From", ""))
    remetente_nome = remetente_lista[0]["nome"] if remetente_lista else ""
    remetente_email = remetente_lista[0]["email"] if remetente_lista else ""

    to_addrs = _parse_addr(msg.get("To", ""))
    cc_addrs = _parse_addr(msg.get("Cc", ""))
    bcc_addrs = _parse_addr(msg.get("Bcc", ""))
    destinatarios = to_addrs + cc_addrs + bcc_addrs

    # Data de envio
    data_str = msg.get("Date", "")
    data_envio: Optional[str] = None
    try:
        if data_str:
            data_envio = parsedate_to_datetime(data_str).isoformat()
    except Exception:
        data_envio = data_str or None

    # Message-ID (identificador único do email)
    message_id = (msg.get("Message-ID", "") or "").strip()
    if not message_id:
        message_id = hashlib.md5(eml_bytes[:512]).hexdigest()

    corpo = _get_body(msg)
    anexos = _get_attachments(msg)

    # Heurística: confirmação de leitura (MDN)
    content_type_base = msg.get_content_type() or ""
    is_mdn = "disposition-notification" in content_type_base
    if not is_mdn and msg.is_multipart():
        for part in msg.walk():
            if "disposition-notification" in (part.get_content_type() or ""):
                is_mdn = True
                break

    return {
        "message_id": message_id,
        "assunto": assunto,
        "remetente_nome": remetente_nome,
        "remetente_email": remetente_email,
        "destinatarios": destinatarios,      # lista de {nome, email}
        "data_envio": data_envio,
        "corpo": corpo,
        "anexos": anexos,                    # lista de {nome, conteudo_bytes, tipo_mime}
        "tem_anexo": len(anexos) > 0,
        "nomes_anexos": [a["nome"] for a in anexos],
        "arquivo_eml_nome": filename,
        "is_mdn": is_mdn,
    }


_MESES_PT = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}


def _sem_acento(txt: str) -> str:
    # Evita dependência externa; substituição suficiente para nomes de meses.
    return (
        (txt or "")
        .replace("á", "a")
        .replace("à", "a")
        .replace("â", "a")
        .replace("ã", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
        .replace("Á", "A")
        .replace("À", "A")
        .replace("Â", "A")
        .replace("Ã", "A")
        .replace("É", "E")
        .replace("Ê", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ô", "O")
        .replace("Õ", "O")
        .replace("Ú", "U")
        .replace("Ç", "C")
    )


def _parse_data_textual(candidato: str) -> Optional[datetime]:
    txt = (candidato or "").strip()
    if not txt:
        return None

    # Tenta parser RFC (ex.: Tue, 23 Jan 2024 14:37:00 -0300)
    try:
        return parsedate_to_datetime(txt)
    except Exception:
        pass

    # Formatos numéricos comuns: 23/01/2024 14:37, 23/01/24 14:37:10, etc.
    m_num = re.search(
        r"(\d{1,2}/\d{1,2}/\d{2,4})(?:\s*(?:,|-|as|às)?\s*)(\d{1,2}:\d{2}(?::\d{2})?\s*(?:[AaPp][Mm])?)?",
        txt,
        flags=re.IGNORECASE,
    )
    if m_num:
        data = m_num.group(1)
        hora = (m_num.group(2) or "00:00").strip()
        blob = f"{data} {hora}".strip()
        for fmt in (
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y %I:%M %p",
            "%d/%m/%Y %I:%M:%S %p",
            "%d/%m/%y %H:%M:%S",
            "%d/%m/%y %H:%M",
            "%d/%m/%y %I:%M %p",
            "%d/%m/%y %I:%M:%S %p",
        ):
            try:
                return datetime.strptime(blob, fmt)
            except Exception:
                continue

    # Formato textual PT: 23 de janeiro de 2024 14:37
    txt_norm = _sem_acento(txt.lower())
    m_pt = re.search(
        r"(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})(?:\s*(?:,|-|as|às)?\s*(\d{1,2}:\d{2}(?::\d{2})?))?",
        txt_norm,
        flags=re.IGNORECASE,
    )
    if m_pt:
        dia = int(m_pt.group(1))
        mes_nome = m_pt.group(2).strip().lower()
        ano = int(m_pt.group(3))
        mes = _MESES_PT.get(mes_nome)
        if mes:
            hora = m_pt.group(4) or "00:00"
            partes_h = hora.split(":")
            hh = int(partes_h[0]) if len(partes_h) >= 1 else 0
            mm = int(partes_h[1]) if len(partes_h) >= 2 else 0
            ss = int(partes_h[2]) if len(partes_h) >= 3 else 0
            try:
                return datetime(ano, mes, dia, hh, mm, ss)
            except Exception:
                return None

    return None


def extrair_data_pedido_do_historico(corpo: str, data_resposta_iso: Optional[str] = None) -> Optional[str]:
    """Tenta inferir data do pedido a partir de histórico citado no corpo de resposta."""
    if not corpo:
        return None

    resposta_dt = None
    if data_resposta_iso:
        try:
            resposta_dt = datetime.fromisoformat(data_resposta_iso)
        except Exception:
            resposta_dt = None

    candidatos_txt = []
    padroes = [
        # Gmail/Outlook PT: Em 23/01/2026 10:30, Fulano escreveu:
        re.compile(r"^\s*Em\s+(.{6,120}?),\s*.+escreveu\s*:\s*$", re.IGNORECASE | re.MULTILINE),
        # Gmail EN: On Tue, Jan 23, 2026 at 10:30 AM ... wrote:
        re.compile(r"^\s*On\s+(.{6,140}?)\s+wrote\s*:\s*$", re.IGNORECASE | re.MULTILINE),
        # Outlook style headers in quoted message
        re.compile(r"^\s*(?:Date|Data|Sent|Enviado\s+em)\s*:\s*(.+)\s*$", re.IGNORECASE | re.MULTILINE),
    ]
    for p in padroes:
        for m in p.finditer(corpo):
            candidatos_txt.append((m.group(1) or "").strip())

    candidatos_dt = []
    for c in candidatos_txt:
        dt = _parse_data_textual(c)
        if dt is not None:
            candidatos_dt.append(dt)

    if not candidatos_dt:
        return None

    if resposta_dt is not None:
        anteriores = []
        for dt in candidatos_dt:
            try:
                if dt.timestamp() <= resposta_dt.timestamp():
                    anteriores.append(dt)
            except Exception:
                continue
        if anteriores:
            return max(anteriores, key=lambda x: x.timestamp()).isoformat()

    # Fallback: usa o candidato mais recente encontrado.
    return max(candidatos_dt, key=lambda x: x.timestamp()).isoformat()


# ---------------------------------------------------------------------------
# Extração de .eml de dentro de um ZIP
# ---------------------------------------------------------------------------

def extrair_emls_do_zip(zip_bytes: bytes) -> list[dict]:
    """
    Extrai todos os arquivos .eml de um ZIP (inclusive subpastas).
    Retorna lista de {'nome', 'conteudo_bytes'}.
    """
    result = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for name in sorted(zf.namelist()):
            if name.lower().endswith(".eml"):
                with zf.open(name) as f:
                    conteudo = f.read()
                basename = os.path.basename(name) or name
                result.append({"nome": basename, "conteudo_bytes": conteudo})
    return result


def extrair_emls_do_tgz(tgz_bytes: bytes) -> list[dict]:
    """
    Extrai todos os arquivos .eml de um .tgz/.tar.gz (inclusive subpastas).
    Retorna lista de {'nome', 'conteudo_bytes'}.
    """
    result = []
    with tarfile.open(fileobj=io.BytesIO(tgz_bytes), mode="r:gz") as tf:
        for member in tf.getmembers():
            if not member.isfile():
                continue
            name = member.name or ""
            if name.lower().endswith(".eml"):
                f = tf.extractfile(member)
                if f is None:
                    continue
                conteudo = f.read()
                basename = os.path.basename(name) or name
                result.append({"nome": basename, "conteudo_bytes": conteudo})
    return result


def extrair_conteudo_compactado(file_name: str, file_bytes: bytes) -> dict:
    """
    Extrai conteúdo de arquivos compactados suportados.

    Retorno:
      {
        'emails': [{'nome', 'conteudo_bytes'}, ...],
        'orcamentos': [{'nome', 'conteudo_bytes'}, ...],
      }
    """
    nome = (file_name or "").lower()
    emails = []
    orcamentos = []

    def _is_budget(name: str) -> bool:
        return name.lower().endswith(SUPPORTED_BUDGET_EXTENSIONS)

    if nome.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            for entry in sorted(zf.namelist()):
                if entry.endswith("/"):
                    continue
                lower = entry.lower()
                with zf.open(entry) as f:
                    payload = f.read()
                base = os.path.basename(entry) or entry
                if lower.endswith(".eml"):
                    emails.append({"nome": base, "conteudo_bytes": payload})
                elif _is_budget(lower):
                    orcamentos.append({"nome": base, "conteudo_bytes": payload})
        return {"emails": emails, "orcamentos": orcamentos}

    if nome.endswith(".tgz") or nome.endswith(".tar.gz"):
        with tarfile.open(fileobj=io.BytesIO(file_bytes), mode="r:gz") as tf:
            for member in tf.getmembers():
                if not member.isfile():
                    continue
                lower = (member.name or "").lower()
                f = tf.extractfile(member)
                if f is None:
                    continue
                payload = f.read()
                base = os.path.basename(member.name) or member.name
                if lower.endswith(".eml"):
                    emails.append({"nome": base, "conteudo_bytes": payload})
                elif _is_budget(lower):
                    orcamentos.append({"nome": base, "conteudo_bytes": payload})
        return {"emails": emails, "orcamentos": orcamentos}

    raise ValueError(f"Formato compactado não suportado: {file_name}")


# ---------------------------------------------------------------------------
# Extração de número de processo do assunto
# ---------------------------------------------------------------------------

_PADROES_PROCESSO = [
    # NUP / SIPAC: 60650.000123/2024-01
    re.compile(r"\b(\d{5,6}\.\d{5,6}/\d{4}-\d{2})\b"),
    # Pregão / TP / RDC + número/ano
    re.compile(
        r"(?:Preg[aã]o|Tomada\s+de\s+Pre[cç]os?|TP|RDC|Concorr[eê]ncia|Ades[aã]o|"
        r"Dispensa|Inexigibilidade)[^\d]{0,15}(\d{1,4}[./]\d{4})",
        re.IGNORECASE,
    ),
    # Processo XXXXXX/AAAA ou Proc. XXX/AAAA
    re.compile(r"(?:Processo|Proc\.?)\s*([\w.\-]{3,30}/\d{4})", re.IGNORECASE),
    # Referência genérica NN/AAAA
    re.compile(r"\b(\d{1,4}/\d{4})\b"),
]


def extrair_numero_processo(texto: str) -> Optional[str]:
    """Tenta extrair número de processo/licitação do texto (normalmente o assunto)."""
    for padrao in _PADROES_PROCESSO:
        m = padrao.search(texto)
        if m:
            return m.group(1).strip()
    return None
