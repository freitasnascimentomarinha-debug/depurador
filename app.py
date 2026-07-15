"""
Depurador de Orçamentos — Sistema Integrado de Pesquisa de Preços
=================================================================
Aplicação única para ingestão de dados espalhados:
1) Orçamentos avulsos (PDF/Word/Excel)
2) Pacotes de e-mails em ZIP/TGZ/TAR.GZ com arquivos .eml

O sistema classifica e-mails, organiza por processo e fornecedor,
extrai itens dos orçamentos e gera mapa comparativo + relatórios.
"""
from __future__ import annotations

import csv
import io
import os
import tarfile
import traceback
import tempfile
import time
import zipfile
import re
import unicodedata
import sys
import html
from difflib import SequenceMatcher
from functools import lru_cache
from datetime import datetime
from io import BytesIO

import streamlit as st

IMPORT_ERROR = None
IMPORT_ERROR_TRACE = ""
try:
    import db_utils
    import file_utils
    import email_utils
    import email_classifier
    import normalize_utils
    import process_db
    import report_docx
    import learning_db
    import ai_judge
    import sanity_check
    import app_config
    import openpyxl
    from export_utils import build_excel
    from match_utils import (
        build_comparison_table,
        build_master_from_consensus,
        encontrar_pares_zona_cinzenta,
        fundir_linhas_julgadas,
    )
except Exception as exc:
    IMPORT_ERROR = exc
    IMPORT_ERROR_TRACE = traceback.format_exc()


SUPPORTED_BUDGET_EXTENSIONS = (".pdf", ".docx", ".doc", ".xlsx", ".xls", ".csv")
SUPPORTED_ARCHIVE_EXTENSIONS = (".zip", ".tgz", ".tar.gz")

_ATTACHMENT_POSITIVE_HINTS = {
    "orcamento",
    "orc",
    "cotacao",
    "cot",
    "proposta",
    "proposta",
    "preco",
    "precos",
    "precificacao",
    "planilha",
    "plan",
    "itens",
    "lote",
    "mapa",
    "valor",
    "valores",
    "comercial",
    "fornecedor",
}

_ATTACHMENT_NEGATIVE_HINTS = {
    "catalogo",
    "brochura",
    "portfolio",
    "manual",
    "foto",
    "fotos",
    "imagem",
    "imagens",
    "produto",
    "produtos",
    "logo",
    "banner",
    "assinatura",
    "rodape",
    "xml",
    "nfe",
    "danfe",
    "boleto",
    "comprovante",
    "nota",
    "pedido",
    "ordem",
    "romaneio",
    "certidao",
    "certificado",
    "ficha",
}


def _default_budget_db_path() -> str:
    if os.path.exists("/mount/src"):
        return "/tmp/orcamentos.db"
    return "orcamentos.db"


def _default_process_db_path() -> str:
    if os.path.exists("/mount/src"):
        return "/tmp/processos_emails.db"
    return "processos_emails.db"


@lru_cache(maxsize=1)
def _get_extract_runtime():
    """Carrega extract_utils sob demanda para reduzir risco de crash no bootstrap."""
    import extract_utils

    return extract_utils.extrair_orcamento_em_camadas, extract_utils.ocr_runtime_status


def _load_openrouter_api_key() -> str:
    # 1) Configuração administrada (config_app.json, aba Configurações → Administração)
    try:
        cfg = app_config.carregar_config()
        if (cfg.get("openrouter_api_key") or "").strip():
            return str(cfg["openrouter_api_key"]).strip()
    except Exception:
        pass
    # 2) Secrets do Streamlit
    for secret_name in ("OPENROUTER_API_KEY", "openrouter_api_key"):
        api_key = st.secrets.get(secret_name)
        if api_key:
            return str(api_key).strip()
    # 3) Variável de ambiente
    return os.getenv("OPENROUTER_API_KEY", "").strip()


def _to_int(value) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _to_float(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _fmt_data(iso: str | None) -> str:
    if not iso:
        return "-"
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(iso)


def _delta_horas(iso_inicio: str | None, iso_fim: str | None) -> str:
    if not iso_inicio or not iso_fim:
        return "-"
    try:
        d1 = datetime.fromisoformat(iso_inicio)
        d2 = datetime.fromisoformat(iso_fim)
        if d1.tzinfo is not None:
            d1 = d1.astimezone().replace(tzinfo=None)
        if d2.tzinfo is not None:
            d2 = d2.astimezone().replace(tzinfo=None)
        delta = d2 - d1
        h = int(delta.total_seconds() // 3600)
        m = int((delta.total_seconds() % 3600) // 60)
        if h >= 24:
            return f"{h // 24}d {h % 24}h"
        return f"{h}h {m}min"
    except Exception:
        return "-"


def _leitura_efetiva(participacao: dict) -> bool:
    """Considera leitura confirmada explicitamente ou implícita por resposta."""
    return bool(
        participacao.get("confirmou_leitura")
        or participacao.get("enviou_orcamento")
        or participacao.get("recusou")
        or participacao.get("fez_pergunta")
        or participacao.get("data_primeira_resposta")
    )


def _norm_header(texto) -> str:
    if texto is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(texto))
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = txt.lower().strip()
    txt = txt.replace("º", " ").replace("°", " ")
    txt = re.sub(r"[^a-z0-9]+", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


_MASTER_NUM_ALIASES = {
    "numero item", "numero_item", "numero", "n", "n item", "num", "item"
}
_MASTER_DESC_ALIASES = {
    "descricao", "nomenclatura", "nome", "material", "objeto", "especificacao", "produto"
}
# Código único do material (PI/NSN/Part Number) — critério mais forte de casamento
_MASTER_COD_ALIASES = {
    "pi", "nsn", "part number", "partnumber", "pn", "codigo", "cod",
    "codigo do item", "codigo item", "cod item", "n estoque", "numero de estoque",
    "num estoque", "referencia", "ref"
}


def _pick_master_column(headers: list[str], aliases: set[str]) -> int | None:
    norm_headers = [_norm_header(h) for h in headers]
    # Prioriza match exato por alias.
    for i, h in enumerate(norm_headers):
        if h in aliases:
            return i
    # Fallback por tokens para casos como "item do edital".
    for i, h in enumerate(norm_headers):
        tokens = set(h.split())
        for alias in aliases:
            alias_tokens = set(alias.split())
            if alias_tokens and alias_tokens.issubset(tokens):
                return i
    return None


def _ler_master_items(master_file) -> list[dict]:
    nome = (master_file.name or "").lower()
    if nome.endswith(".csv"):
        conteudo = master_file.getvalue().decode("utf-8", errors="ignore")
        reader = csv.DictReader(conteudo.splitlines())
        headers = list(reader.fieldnames or [])
        idx_num = _pick_master_column(headers, _MASTER_NUM_ALIASES)
        idx_desc = _pick_master_column(headers, _MASTER_DESC_ALIASES)
        idx_cod = _pick_master_column(headers, _MASTER_COD_ALIASES)
        if idx_num is None and idx_desc is None:
            return []

        out = []
        for r in reader:
            vals = [r.get(h) for h in headers]
            out.append({
                "numero_item": vals[idx_num] if idx_num is not None and idx_num < len(vals) else None,
                "codigo": vals[idx_cod] if idx_cod is not None and idx_cod < len(vals) else None,
                "descricao": vals[idx_desc] if idx_desc is not None and idx_desc < len(vals) else None,
            })
        return out

    if nome.endswith(".xls") and not nome.endswith(".xlsx"):
        try:
            import xlrd
        except Exception as exc:
            raise RuntimeError("Leitura de .xls indisponível. Instale a dependência 'xlrd'.") from exc

        wb = xlrd.open_workbook(file_contents=master_file.getvalue())
        if wb.nsheets <= 0:
            return []
        ws = wb.sheet_by_index(0)
        if ws.nrows <= 0:
            return []

        headers = [str(h).strip() if h is not None else "" for h in ws.row_values(0)]
        idx_num = _pick_master_column(headers, _MASTER_NUM_ALIASES)
        idx_desc = _pick_master_column(headers, _MASTER_DESC_ALIASES)
        idx_cod = _pick_master_column(headers, _MASTER_COD_ALIASES)
        if idx_num is None and idx_desc is None:
            return []

        out = []
        for ridx in range(1, ws.nrows):
            row = ws.row_values(ridx)
            out.append({
                "numero_item": row[idx_num] if idx_num is not None and idx_num < len(row) else None,
                "codigo": row[idx_cod] if idx_cod is not None and idx_cod < len(row) else None,
                "descricao": row[idx_desc] if idx_desc is not None and idx_desc < len(row) else None,
            })
        return out

    wb = openpyxl.load_workbook(master_file, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    idx_num = _pick_master_column(headers, _MASTER_NUM_ALIASES)
    idx_desc = _pick_master_column(headers, _MASTER_DESC_ALIASES)
    idx_cod = _pick_master_column(headers, _MASTER_COD_ALIASES)
    if idx_num is None and idx_desc is None:
        return []

    out = []
    for row in rows[1:]:
        out.append({
            "numero_item": row[idx_num] if idx_num is not None and idx_num < len(row) else None,
            "codigo": row[idx_cod] if idx_cod is not None and idx_cod < len(row) else None,
            "descricao": row[idx_desc] if idx_desc is not None and idx_desc < len(row) else None,
        })
    return out


def _is_budget_file(name: str) -> bool:
    return (name or "").lower().endswith(SUPPORTED_BUDGET_EXTENSIONS)


def _norm_filename_tokens(name: str) -> set[str]:
    low = (name or "").lower()
    no_ext, _ = os.path.splitext(low)
    no_acc = unicodedata.normalize("NFKD", no_ext).encode("ascii", "ignore").decode("ascii")
    parts = re.split(r"[^a-z0-9]+", no_acc)
    return {p for p in parts if p}


def _token_similaridade(token: str, termos: set[str]) -> float:
    if not token or not termos:
        return 0.0
    melhor = 0.0
    for termo in termos:
        ratio = SequenceMatcher(None, token, termo).ratio()
        if ratio > melhor:
            melhor = ratio
    return melhor


# Prefixos de nome de arquivo que indicam claramente não ser orçamento
_ATTACHMENT_NEGATIVE_PREFIXES = (
    "img", "image", "imagem", "foto", "whatsapp",
    "assinatura", "logo", "banner", "icon",
)


def _score_email_attachment_candidate(name: str, mime: str, email_tipo: str) -> tuple[int, str]:
    """Pontua chance do anexo ser cotacao/orcamento real.
    Retorna score > 0 apenas para candidatos válidos; <= 0 rejeita.
    """
    low_name = (name or "").lower()

    # Extensão não suportada como documento de orçamento
    if not _is_budget_file(low_name):
        return -99, "extensao_nao_suportada"

    mime_l = (mime or "").lower()

    # Imagem pelo MIME → sempre rejeitar
    if mime_l.startswith("image/"):
        return -10, "mime_imagem"

    # Prefixo claramente não-orçamento
    if any(low_name.startswith(p) for p in _ATTACHMENT_NEGATIVE_PREFIXES):
        return -10, "prefixo_negativo"

    ext = os.path.splitext(low_name)[1]
    tokens = _norm_filename_tokens(low_name)

    # Contar sinais negativos e positivos no nome do arquivo
    neg_exatos = tokens & _ATTACHMENT_NEGATIVE_HINTS
    neg_fuzzy = sum(
        1 for tk in tokens
        if tk not in neg_exatos
        and _token_similaridade(tk, _ATTACHMENT_NEGATIVE_HINTS) >= 0.85
    )
    neg_total = len(neg_exatos) + neg_fuzzy

    pos_exatos = tokens & _ATTACHMENT_POSITIVE_HINTS
    pos_fuzzy = sum(
        1 for tk in tokens
        if tk not in pos_exatos
        and _token_similaridade(tk, _ATTACHMENT_POSITIVE_HINTS) >= 0.82
    )
    pos_total = len(pos_exatos) + pos_fuzzy

    # Sinais negativos dominam: rejeitar quando negativo >= positivo (e há pelo menos 1 negativo)
    if neg_total > 0 and neg_total >= pos_total:
        return -(neg_total * 3), "sinal_negativo"

    # --- Calcular pontuação positiva ---
    score = 0

    if ext in {".pdf", ".xlsx", ".xls", ".csv"}:
        score += 3
    elif ext in {".docx", ".doc"}:
        score += 2

    score += pos_total * 3

    # PDF é preferido quando o nome do arquivo contém palavras-chave de cotação
    if ext == ".pdf" and (tokens & {"cotacao", "proposta", "orcamento"}):
        score += 2

    if email_tipo == "orcamento_recebido":
        score += 3
    elif email_tipo == "pedido_orcamento":
        score -= 2
    # Lição de campo: e-mails de dúvida/declínio às vezes trazem anexo com
    # cotação válida (caso Nexbolt, PE-90043/2026). A classificação do e-mail
    # e o aproveitamento do anexo são decisões independentes — não penalizar.
    # O template em branco devolvido nesses e-mails é barrado pelo hash
    # (template_hashes), não pela classificação.

    # Arquivo genérico (sem nenhum sinal positivo no nome) só passa se o
    # email foi explicitamente classificado como orçamento recebido.
    if pos_total == 0 and email_tipo != "orcamento_recebido":
        return -1, "sem_indicio_textual"

    # Pontuação final negativa ou zero → rejeitar
    if score <= 0:
        return score, "baixo_sinal_de_orcamento"

    return score, "candidato_valido"


def _is_archive_file(name: str) -> bool:
    low = (name or "").lower()
    return low.endswith(SUPPORTED_ARCHIVE_EXTENSIONS)


def _persist_bytes(tmpdir: str, base_name: str, data: bytes, prefix: str = "") -> str:
    safe_name = os.path.basename(base_name).replace("..", "_")
    if prefix:
        safe_name = f"{prefix}_{safe_name}"
    out_path = os.path.join(tmpdir, safe_name)
    with open(out_path, "wb") as fh:
        fh.write(data)
    return out_path


def _extrair_compactado_compat(file_name: str, file_bytes: bytes) -> dict:
    """
    Compatibilidade entre versões do módulo email_utils.
    Se existir extrair_conteudo_compactado no módulo, usa a implementação oficial.
    Caso contrário, usa fallback local para ZIP/TGZ/TAR.GZ.
    """
    if hasattr(email_utils, "extrair_conteudo_compactado"):
        return email_utils.extrair_conteudo_compactado(file_name, file_bytes)

    nome = (file_name or "").lower()
    emails = []
    orcamentos = []

    if nome.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            for entry in sorted(zf.namelist()):
                if entry.endswith("/"):
                    continue
                with zf.open(entry) as f:
                    payload = f.read()
                lower = entry.lower()
                base = os.path.basename(entry) or entry
                if lower.endswith(".eml"):
                    emails.append({"nome": base, "conteudo_bytes": payload})
                elif _is_budget_file(lower):
                    orcamentos.append({"nome": base, "conteudo_bytes": payload})
        return {"emails": emails, "orcamentos": orcamentos}

    if nome.endswith(".tgz") or nome.endswith(".tar.gz"):
        with tarfile.open(fileobj=io.BytesIO(file_bytes), mode="r:gz") as tf:
            for member in tf.getmembers():
                if not member.isfile():
                    continue
                f = tf.extractfile(member)
                if f is None:
                    continue
                payload = f.read()
                lower = (member.name or "").lower()
                base = os.path.basename(member.name) or member.name
                if lower.endswith(".eml"):
                    emails.append({"nome": base, "conteudo_bytes": payload})
                elif _is_budget_file(lower):
                    orcamentos.append({"nome": base, "conteudo_bytes": payload})
        return {"emails": emails, "orcamentos": orcamentos}

    raise ValueError(f"Formato compactado não suportado: {file_name}")


def _atualizar_cnpjs_fornecedores_compat(conn_proc, processo_id: int | None = None) -> int:
    """Compatibilidade: executa backfill de CNPJ se a função existir no módulo atual."""
    fn = getattr(process_db, "atualizar_cnpjs_fornecedores", None)
    if not callable(fn):
        return 0
    try:
        if processo_id is None:
            return int(fn(conn_proc) or 0)
        return int(fn(conn_proc, processo_id) or 0)
    except Exception:
        return 0


def _atualizar_fornecedor_dados_por_nome_compat(
    conn_proc,
    processo_id: int,
    nome_referencia: str,
    cnpj: str = "",
    telefone: str = "",
) -> int:
    """Compatibilidade para enriquecer fornecedor por nome com CNPJ/telefone vindo de orçamento."""
    fn = getattr(process_db, "atualizar_dados_fornecedor_por_nome_no_processo", None)
    if not callable(fn):
        return 0
    try:
        return int(fn(conn_proc, processo_id, nome_referencia, cnpj, telefone) or 0)
    except Exception:
        return 0


def _extrair_cnpj_telefone_de_bytes(conteudo_bytes: bytes, nome_arquivo: str = "") -> tuple[str, str, str]:
    """Tenta extrair CNPJ e telefone de bytes de arquivo (prioriza PDF)."""
    texto = ""
    lower = (nome_arquivo or "").lower()
    try:
        if lower.endswith(".pdf"):
            try:
                import pdfplumber
                from io import BytesIO

                with pdfplumber.open(BytesIO(conteudo_bytes)) as pdf:
                    if len(pdf.pages) >= 1:
                        p0 = pdf.pages[0]
                        t0 = p0.extract_text() or ""
                        texto += t0[:2000]
                    if len(pdf.pages) >= 2:
                        plast = pdf.pages[-1]
                        tlast = plast.extract_text() or ""
                        texto += "\n" + tlast[-2000:]
            except Exception:
                texto = ""
        elif lower.endswith(('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp')):
            try:
                from PIL import Image
                import pytesseract
                from io import BytesIO

                img = Image.open(BytesIO(conteudo_bytes))
                texto = pytesseract.image_to_string(img)
            except Exception:
                texto = ""
        else:
            # Tentativa simples para DOCX/HTML/texto
            try:
                texto = conteudo_bytes.decode('utf-8', errors='ignore')[:4000]
            except Exception:
                texto = ""
    except Exception:
        texto = ""

    # Usa funções de process_db para extrair padrões
    cnpj = process_db._extrair_cnpj_do_texto(texto)
    tel = process_db._extrair_telefone_do_texto(texto)
    return cnpj, tel, texto


def _validar_cnpj(cnpj: str) -> bool:
    if not cnpj:
        return False
    dig = re.sub(r"\D", "", cnpj)
    if len(dig) != 14:
        return False
    if dig == dig[0] * 14:
        return False

    def _calc(digs: str, mults: list[int]) -> int:
        s = sum(int(a) * b for a, b in zip(digs, mults))
        r = s % 11
        return 0 if r < 2 else 11 - r

    mult1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    mult2 = [6] + mult1
    d1 = _calc(dig[:12], mult1)
    d2 = _calc(dig[:12] + str(d1), mult2)
    return dig.endswith(f"{d1}{d2}")


VALID_DDDS = {
    '11','12','13','14','15','16','17','18','19',
    '21','22','24',
    '27','28',
    '31','32','33','34','35','37','38',
    '41','42','43','44','45','46',
    '47','48','49',
    '51','53','54','55',
    '61','62','64','63','65','66','67',
    '68','69','71','73','74','75','77','79','81','82','83','84','85','86','87','88','89','91','92','93','94','95','96','97','98','99'
}


def _validar_telefone(numero: str) -> bool:
    if not numero:
        return False
    dig = re.sub(r"\D", "", numero)
    if dig.startswith('55') and len(dig) in (12, 13):
        dig = dig[2:]
    if len(dig) not in (10, 11):
        return False
    ddd = dig[:2]
    return ddd in VALID_DDDS


def _formatar_telefone(numero: str) -> str:
    if not numero:
        return ""
    dig = re.sub(r"\D", "", numero)
    if dig.startswith('55'):
        dig = dig[2:]
    if len(dig) == 11:
        return f"({dig[:2]}) {dig[2:7]}-{dig[7:]}"
    if len(dig) == 10:
        return f"({dig[:2]}) {dig[2:6]}-{dig[6:]}"
    return numero


def _call_ai_extract(corpus_text: str, empresa: str, api_key_local: str, model_name: str) -> tuple[str, str]:
    """Chama o modelo via OpenRouter para extrair CNPJ e telefone a partir de um texto.
    Retorna (cnpj, telefone) ou ('','') em falha.
    """
    if not api_key_local:
        return "", ""
    try:
        url = "https://api.openrouter.ai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key_local}",
            "Content-Type": "application/json",
        }
        prompt = (
            f"Extraia somente o CNPJ e o telefone da empresa {empresa} a partir do texto fornecido. "
            "Responda estritamente em JSON com chaves 'cnpj' e 'telefone'. Se não encontrar, deixe valor vazio."
            "\n\nTexto:\n" + corpus_text[:6000]
        )
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": "Você é um assistente que extrai CNPJ e telefone brasileiros."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.0,
            "max_tokens": 200,
        }
        import requests
        resp = requests.post(url, headers=headers, json=payload, timeout=15)
        if resp.status_code != 200:
            return "", ""
        data = resp.json()
        text = ""
        # extrai texto da resposta
        try:
            text = data.get('choices', [])[0].get('message', {}).get('content', '')
        except Exception:
            text = data.get('choices', [])[0].get('text', '') if data.get('choices') else ''
        # busca cnpj/telefone em texto retornado
        cnpj = process_db._extrair_cnpj_do_texto(text)
        tel = process_db._extrair_telefone_do_texto(text)
        return cnpj or "", tel or ""
    except Exception:
        return "", ""


def _enriquecer_fornecedores_por_orcamentos(conn_proc, conn_budget, processo_id: int) -> int:
    """Procura CNPJ/telefone nos caches de orçamentos e em anexos de e-mail vinculados.
    Atualiza fornecedores no processo por aproximação de nome quando encontrar dados.
    Retorna o número de fornecedores atualizados (estimado).
    """
    atualizados = 0
    try:
        orcamentos = process_db.listar_orcamentos_do_processo(conn_proc, processo_id)
        for reg in orcamentos:
            file_id = reg.get('file_id')
            nome_arquivo = reg.get('nome_arquivo') or ''
            # 1) checa cache do orçamento (db_utils)
            cached = None
            try:
                cached = db_utils.get_cached_file(conn_budget, file_id) or {}
            except Exception:
                cached = {}
            cnpj = (cached.get('cnpj') or '').strip() if cached else ''
            tel = (cached.get('telefone') or '').strip() if cached else ''
            empresa = (cached.get('empresa') or nome_arquivo) if cached else nome_arquivo
            # valida cache primeiro
            valid_cnpj = cnpj and _validar_cnpj(cnpj)
            valid_tel = tel and _validar_telefone(tel)
            if valid_cnpj or valid_tel:
                cnpj_use = cnpj if valid_cnpj else ""
                tel_use = _formatar_telefone(tel) if valid_tel else ""
                n = _atualizar_fornecedor_dados_por_nome_compat(conn_proc, processo_id, empresa, cnpj_use, tel_use)
                atualizados += int(n)
                continue

            # 2) se não encontrou no cache, tenta extrair do anexo do e-mail se houver
            if isinstance(file_id, str) and file_id.startswith('email_attachment:'):
                import re
                m = re.match(r'^email_attachment:(\d+):', file_id)
                if m:
                    email_id = int(m.group(1))
                    anexos = process_db.listar_anexos_email(conn_proc, email_id)
                    for anx in anexos:
                        nome_anx = (anx.get('nome') or '').strip()
                        if nome_arquivo and nome_arquivo not in nome_anx and nome_anx not in nome_arquivo:
                            # tenta permanecer permissivo: continue apenas se nomes muito diferentes
                            pass
                        conteudo = process_db.get_anexo_conteudo(conn_proc, anx['id'])
                        if not conteudo:
                            continue
                        cnpj2, tel2, texto = _extrair_cnpj_telefone_de_bytes(conteudo, nome_anx)
                        valid_cnpj2 = cnpj2 and _validar_cnpj(cnpj2)
                        valid_tel2 = tel2 and _validar_telefone(tel2)

                        final_cnpj = cnpj2 if valid_cnpj2 else ""
                        final_tel = _formatar_telefone(tel2) if valid_tel2 else ""

                        # Detecta múltiplos candidatos conflitantes no texto
                        cand_cnpjs = list(dict.fromkeys(re.findall(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}", texto)))
                        cand_tels = list(dict.fromkeys(re.findall(r"(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?\d{4,5}[\s.-]?\d{4}", texto)))
                        conflict_cnpj = len([c for c in cand_cnpjs if c and c.strip()]) > 1
                        conflict_tel = len([t for t in cand_tels if t and t.strip()]) > 1

                        # Política ajustada: chamar IA apenas quando
                        # - ambos CNPJ e telefone estiverem ausentes/inválidos, OU
                        # - houver múltiplos candidatos conflitantes para CNPJ ou telefone
                        need_ai = False
                        if (not valid_cnpj2 and not valid_tel2):
                            need_ai = True
                        if conflict_cnpj or conflict_tel:
                            need_ai = True

                        ai_used = False
                        if need_ai:
                            try:
                                model_name = st.session_state.get('cfg_model', 'google/gemini-2.5-flash')
                                ai_cnpj, ai_tel = _call_ai_extract(texto, empresa or nome_anx, api_key, model_name)
                                if ai_cnpj and not final_cnpj and _validar_cnpj(ai_cnpj):
                                    final_cnpj = ai_cnpj
                                if ai_tel and not final_tel and _validar_telefone(ai_tel):
                                    final_tel = _formatar_telefone(ai_tel)
                                ai_used = True
                            except Exception:
                                ai_used = False

                        if final_cnpj or final_tel:
                            n = _atualizar_fornecedor_dados_por_nome_compat(conn_proc, processo_id, nome_arquivo or nome_anx, final_cnpj, final_tel)
                            atualizados += int(n)
                            # registra auditoria local se IA foi usada para chegar ao resultado
                            try:
                                if ai_used:
                                    with open('ia_updates.log', 'a', encoding='utf-8') as f:
                                        from datetime import datetime
                                        f.write(f"{datetime.utcnow().isoformat()}Z\tproc:{processo_id}\tfile:{nome_anx or nome_arquivo}\tcnpj:{final_cnpj}\ttel:{final_tel}\n")
                            except Exception:
                                pass
                            break
    except Exception:
        pass
    return atualizados


def _itens_orcados_por_fornecedor(conn_proc, conn_budget, processo_id: int) -> dict[str, int]:
    """Conta itens orçados por fornecedor usando o e-mail do remetente salvo diretamente
    na tabela processo_orcamentos.  Para registros antigos sem esse campo, usa fallback
    por similaridade de nome da empresa extraída do orçamento.
    """
    fn_backfill = getattr(process_db, "backfill_remetente_email_orcamentos", None)
    if callable(fn_backfill):
        try:
            fn_backfill(conn_proc, processo_id)
        except Exception:
            pass

    regs = process_db.listar_orcamentos_do_processo(conn_proc, processo_id)
    if not regs:
        return {}

    def _norm(v: str) -> str:
        txt = unicodedata.normalize("NFKD", (v or "").lower())
        txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
        txt = re.sub(r"[^a-z0-9]+", " ", txt)
        return " ".join(txt.split())

    stop = {
        "ltda", "eireli", "me", "epp", "sa", "s", "a", "comercio", "servicos",
        "de", "da", "do", "das", "dos", "empresa", "sociedade", "limitada",
    }

    def _tokens(v: str) -> set[str]:
        return {t for t in _norm(v).split() if len(t) >= 3 and t not in stop}

    # Passo 1: agrupa itens por e-mail diretamente do campo remetente_email.
    itens_por_email: dict[str, int] = {}
    regs_sem_email: list[dict] = []
    for reg in regs:
        remetente = (reg.get("remetente_email") or "").strip().lower()
        if not remetente:
            regs_sem_email.append(reg)
            continue
        qtd = len(db_utils.get_items_for_file(conn_budget, reg["file_id"]))
        itens_por_email[remetente] = itens_por_email.get(remetente, 0) + qtd

    # Passo 2: fallback por nome de empresa para registros sem e-mail vinculado.
    itens_por_empresa: dict[str, int] = {}
    for reg in regs_sem_email:
        cached = db_utils.get_cached_file(conn_budget, reg["file_id"]) or {}
        empresa = (cached.get("empresa") or reg.get("nome_arquivo") or "").strip()
        if not empresa:
            continue
        qtd = len(db_utils.get_items_for_file(conn_budget, reg["file_id"]))
        itens_por_empresa[empresa] = itens_por_empresa.get(empresa, 0) + qtd

    # Monta resultado final por e-mail de fornecedor.
    out: dict[str, int] = {}
    fornecedores = process_db.listar_fornecedores_processo(conn_proc, processo_id)
    for f in fornecedores:
        email_f = (f.get("email") or "").strip().lower()
        # Prioridade: vínculo exato por e-mail.
        total = int(itens_por_email.get(email_f, 0)) if email_f else 0

        # Fallback: similaridade de nome apenas para registros sem e-mail.
        nome_f = (f.get("nome") or "").strip()
        if nome_f and itens_por_empresa:
            nome_f_n = _norm(nome_f)
            tok_f = _tokens(nome_f)
            for emp, qtd in itens_por_empresa.items():
                emp_n = _norm(emp)
                tok_e = _tokens(emp)
                inter = len(tok_f.intersection(tok_e)) if tok_f and tok_e else 0
                ratio_f = inter / max(1, len(tok_f)) if tok_f else 0.0
                ratio_e = inter / max(1, len(tok_e)) if tok_e else 0.0
                if (
                    nome_f_n in emp_n
                    or emp_n in nome_f_n
                    or ratio_f >= 0.6
                    or ratio_e >= 0.6
                    or inter >= 2
                ):
                    total += int(qtd)

        if total > 0 and email_f:
            out[email_f] = total

    # Garante que remetentes com itens mas sem registro em fornecedores também apareçam.
    for email, qtd in itens_por_email.items():
        if qtd > 0 and email not in out:
            out[email] = int(qtd)

    return out


def _build_item_cobertura(rows: list[dict], matrix: dict, base_meta: int = 3) -> list[dict]:
    """Gera visão de cobertura por item: qtd de orçamentos e % (base >=3 = 100%)."""
    out = []
    base = max(1, int(base_meta or 3))
    for r in rows or []:
        key = r.get("_key")
        if key is None:
            key = ("num", str(r.get("numero_item")).strip()) if r.get("numero_item") else (
                "desc", " ".join(str(r.get("descricao") or "").strip().lower().split())
            )
        dados = matrix.get(key, {}) or {}
        qtd_orc = sum(1 for _emp, info in dados.items() if info.get("preco_unitario") is not None)
        perc = min(qtd_orc, base) / float(base) * 100.0
        out.append({
            "Item": r.get("numero_item") or "-",
            "Descrição": r.get("descricao") or "-",
            "Orçamentos": int(qtd_orc),
            "% conclusão": f"{perc:.1f}%",
        })
    return out


def _resumo_cobertura_itens(item_cobertura: list[dict]) -> dict[str, float | int]:
    """Resume a cobertura de itens por faixas de quantidade de orçamentos."""
    total_itens = len(item_cobertura or [])
    itens_zero = sum(1 for r in (item_cobertura or []) if int(r.get("Orçamentos") or 0) == 0)
    itens_1_a_2 = sum(1 for r in (item_cobertura or []) if 1 <= int(r.get("Orçamentos") or 0) <= 2)
    itens_3_ou_mais = sum(1 for r in (item_cobertura or []) if int(r.get("Orçamentos") or 0) >= 3)
    perc_zero = (itens_zero / total_itens * 100.0) if total_itens else 0.0
    perc_1_a_2 = (itens_1_a_2 / total_itens * 100.0) if total_itens else 0.0
    perc_3_ou_mais = (itens_3_ou_mais / total_itens * 100.0) if total_itens else 0.0
    return {
        "total_itens": int(total_itens),
        "itens_zero": int(itens_zero),
        "itens_1_a_2": int(itens_1_a_2),
        "itens_3_ou_mais": int(itens_3_ou_mais),
        "perc_zero": float(perc_zero),
        "perc_1_a_2": float(perc_1_a_2),
        "perc_3_ou_mais": float(perc_3_ou_mais),
    }


def _categorizar_itens_por_orcamentos(item_cobertura: list[dict]) -> dict[int, list[tuple[str, str]]]:
    """Retorna dicionário com chaves 0,1,2,3 (3 significa 3 ou mais) mapeando para
    lista de tuplas (numero_item, descricao).
    """
    cat = {0: [], 1: [], 2: [], 3: []}
    for r in (item_cobertura or []):
        try:
            qtd = int(r.get("Orçamentos") or 0)
        except Exception:
            qtd = 0
        chave = 3 if qtd >= 3 else qtd
        num = r.get("Item") or r.get("numero_item") or "-"
        desc = r.get("Descrição") or r.get("descricao") or "-"
        cat.setdefault(chave, []).append((str(num), str(desc)))
    return cat


def _safe_dataframe(rows: list[dict], *, empty_message: str | None = None) -> None:
    """Renderiza tabela priorizando estabilidade em runtimes 3.14+ no Streamlit Cloud."""
    if not rows:
        if empty_message:
            st.info(empty_message)
        return

    if sys.version_info < (3, 14):
        st.dataframe(rows, use_container_width=True)
        return

    cols = list(rows[0].keys())
    head = "".join(f"<th>{html.escape(str(c))}</th>" for c in cols)
    body_rows = []
    for r in rows:
        tds = "".join(f"<td>{html.escape(str(r.get(c, '')))}</td>" for c in cols)
        body_rows.append(f"<tr>{tds}</tr>")

    table_html = (
        "<div style='overflow-x:auto;'>"
        "<table style='width:100%; border-collapse:collapse; font-size:0.90rem;'>"
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def _select_processo(conn_proc, key: str) -> dict | None:
    processos = process_db.listar_processos(conn_proc)
    if not processos:
        st.info("Nenhum processo cadastrado.")
        return None

    opcoes = [f"{p['numero']} - {p.get('titulo') or 'sem titulo'}" for p in processos]
    idx = st.selectbox("Processo", range(len(opcoes)), format_func=lambda i: opcoes[i], key=key)
    return processos[idx]


def _rebuild_map_for_process(conn_proc, conn_budget, processo_id: int, fuzzy_threshold: int,
                             sanity_threshold: int, usar_uf: bool, usar_qtd: bool,
                             bloquear_numero_incoerente: bool, gerar_master_auto: bool,
                             min_agree: int, consensus_threshold: int):
    vinculados = process_db.listar_orcamentos_do_processo(conn_proc, processo_id)
    if not vinculados:
        return None

    all_extractions = []
    sources = []
    for reg in vinculados:
        itens = db_utils.get_items_for_file(conn_budget, reg["file_id"])
        if not itens:
            continue
        cached = db_utils.get_cached_file(conn_budget, reg["file_id"])
        empresa = (cached or {}).get("empresa") or reg.get("nome_arquivo") or reg["file_id"]
        arquivo = reg.get("nome_arquivo") or reg["file_id"]
        all_extractions.append({"empresa": empresa, "arquivo": arquivo, "itens": itens})
        for item in itens:
            sources.append({
                "empresa": empresa,
                "arquivo": arquivo,
                "fonte_extracao": item.get("fonte_extracao", "cache"),
                "origem": item.get("origem", "desconhecida"),
                "numero_item": item.get("numero_item"),
                "descricao": item.get("descricao"),
            })

    if not all_extractions:
        return None

    master_items = None
    if gerar_master_auto:
        auto_master = build_master_from_consensus(
            all_extractions,
            min_agree=int(min_agree),
            consensus_threshold=int(consensus_threshold),
        )
        if auto_master:
            master_items = auto_master

    try:
        correcoes_salvas = learning_db.carregar_correcoes(learning_db.get_connection())
    except Exception:
        correcoes_salvas = {}
    rows, _row_index, matrix, review = build_comparison_table(
        all_extractions,
        master_items=master_items,
        fuzzy_threshold=int(fuzzy_threshold),
        sanity_threshold=int(sanity_threshold),
        usar_uf=bool(usar_uf),
        usar_qtd=bool(usar_qtd),
        bloquear_numero_incoerente=bool(bloquear_numero_incoerente),
        correcoes=correcoes_salvas,
    )
    item_cobertura = _build_item_cobertura(rows, matrix, base_meta=3)
    empresas = sorted({e["empresa"] for e in all_extractions})
    excel_buffer = build_excel(rows, matrix, empresas, review, sources) if rows else None
    return {
        "excel_buffer": excel_buffer,
        "rows_count": len(rows),
        "empresas_count": len(empresas),
        "review_count": len(review),
        "item_cobertura": item_cobertura,
        "processo_id": processo_id,
    }


def _build_report_pdf(processo: dict, resumo: dict) -> bytes:
    """Gera um relatório PDF visual com métricas e barras simples."""
    os.environ.setdefault("RL_NOACCEL", "1")
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.pdfgen import canvas
    except Exception as exc:
        raise RuntimeError(
            "Dependência de PDF ausente. Instale reportlab."
        ) from exc

    def _label_sem_icone(label: str) -> str:
        for pref in ("📤 ", "💰 ", "👁️ ", "🚫 ", "❓ ", "🏢 ", "📄 "):
            if label.startswith(pref):
                return label[len(pref):]
        return label

    def _draw_bar_block(cnv, x, y, width, title, items, color_rgb):
        cnv.setFont("Helvetica-Bold", 10)
        cnv.drawString(x, y, title)
        y -= 0.45 * cm
        if not items:
            cnv.setFont("Helvetica", 9)
            cnv.drawString(x, y, "Sem dados")
            return y - 0.4 * cm

        max_value = max(v for _, v in items) or 1
        bar_max = width - 5.5 * cm
        cnv.setFont("Helvetica", 8)
        for label, value in items:
            label = str(label)[:34]
            cnv.setFillColorRGB(0, 0, 0)
            cnv.drawString(x, y, label)
            bar_w = (float(value) / float(max_value)) * bar_max if max_value else 0
            cnv.setFillColorRGB(*color_rgb)
            cnv.rect(x + 4.7 * cm, y - 0.12 * cm, max(bar_w, 0.15 * cm), 0.22 * cm, stroke=0, fill=1)
            cnv.setFillColorRGB(0, 0, 0)
            cnv.drawRightString(x + width, y, str(value))
            y -= 0.42 * cm
        return y - 0.25 * cm

    def _short(txt: str, max_len: int) -> str:
        t = str(txt or "").strip()
        if len(t) <= max_len:
            return t
        return t[: max(0, max_len - 3)] + "..."

    def _draw_fornecedor_header(cnv, x, y):
        cnv.setFont("Helvetica-Bold", 8)
        cnv.drawString(x, y, "Fornecedor")
        cnv.drawString(x + 6.4 * cm, y, "Itens")
        cnv.drawString(x + 7.7 * cm, y, "Status")
        cnv.drawString(x + 10.6 * cm, y, "Pedido")
        cnv.drawString(x + 13.3 * cm, y, "Resposta")
        cnv.drawString(x + 16.0 * cm, y, "Tempo")
        cnv.line(x, y - 0.08 * cm, x + 18.1 * cm, y - 0.08 * cm)

    por_tipo = resumo.get("por_tipo") or {}
    chart_tipo = [
        (_label_sem_icone(email_classifier.rotulo_tipo(k)), int(v))
        for k, v in sorted(por_tipo.items(), key=lambda x: -x[1])
    ]

    part = resumo.get("participacoes") or []
    itens_orcados_map = resumo.get("itens_orcados_por_fornecedor") or {}
    cobertura_resumo = _resumo_cobertura_itens(resumo.get("item_cobertura") or [])
    status_map = {
        "Enviaram orçamento": sum(1 for p in part if p.get("enviou_orcamento")),
        "Recusaram": sum(1 for p in part if p.get("recusou")),
        "Leitura (efetiva)": sum(1 for p in part if _leitura_efetiva(p)),
        "Com dúvida": sum(1 for p in part if p.get("fez_pergunta")),
        "Sem resposta": resumo.get("sem_resposta") or 0,
    }
    chart_status = [(k, max(0, int(v))) for k, v in status_map.items()]

    # PDF
    out = BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    width, height = A4

    # Cabeçalho
    c.setFillColorRGB(0.08, 0.25, 0.55)
    c.rect(0, height - 3.0 * cm, width, 3.0 * cm, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(1.2 * cm, height - 1.6 * cm, "Relatório de Pesquisa de Preços")
    c.setFont("Helvetica", 10)
    c.drawString(1.2 * cm, height - 2.4 * cm, f"Processo: {processo.get('numero')}")

    y = height - 3.8 * cm
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica", 9)
    c.drawString(1.2 * cm, y, f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    y -= 0.6 * cm
    c.drawString(1.2 * cm, y, f"Título: {processo.get('titulo') or '-'}")
    y -= 0.9 * cm

    # KPIs
    kpis = [
        f"Total de e-mails: {resumo.get('total_emails', 0)}",
        f"Fornecedores contatados: {resumo.get('total_fornecedores', 0)}",
        f"Enviaram orçamento: {resumo.get('enviaram_orcamento', 0)}",
        f"Recusaram: {resumo.get('recusaram', 0)}",
        f"Sem resposta: {resumo.get('sem_resposta', 0)}",
        f"Quantidade total de itens: {cobertura_resumo['total_itens']}",
        (
            "Itens com zero orcamentos: "
            f"{cobertura_resumo['itens_zero']} ({cobertura_resumo['perc_zero']:.1f}%)"
        ),
        (
            "Itens com 1 a 2 orçamentos: "
            f"{cobertura_resumo['itens_1_a_2']} ({cobertura_resumo['perc_1_a_2']:.1f}%)"
        ),
        (
            "Itens com 3 ou mais orçamentos: "
            f"{cobertura_resumo['itens_3_ou_mais']} ({cobertura_resumo['perc_3_ou_mais']:.1f}%)"
        ),
    ]
    tmr = resumo.get("tempo_medio_resposta_h")
    if tmr is not None:
        kpis.append(f"Tempo médio de resposta: {float(tmr):.1f}h ({float(tmr)/24.0:.2f} dias)")

    c.setFont("Helvetica-Bold", 11)
    c.drawString(1.2 * cm, y, "Indicadores")
    y -= 0.55 * cm
    c.setFont("Helvetica", 10)
    for item in kpis:
        c.drawString(1.4 * cm, y, f"- {item}")
        y -= 0.5 * cm

    y -= 0.2 * cm
    y = _draw_bar_block(c, 1.2 * cm, y, 17.0 * cm, "Distribuição de E-mails por Tipo", chart_tipo, (0.08, 0.25, 0.55))
    y = _draw_bar_block(c, 1.2 * cm, y, 17.0 * cm, "Status dos Fornecedores", chart_status, (0.18, 0.49, 0.20))

    # Adiciona listas de itens por número de orçamentos
    categorias = _categorizar_itens_por_orcamentos(resumo.get("item_cobertura") or [])
    if y < 6.0 * cm:
        c.showPage()
        y = height - 1.8 * cm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(1.2 * cm, y, "Itens classificados por número de orçamentos")
    y -= 0.6 * cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.2 * cm, y, "Itens com zero orçamento")
    y -= 0.45 * cm
    c.setFont("Helvetica", 9)
    if categorias.get(0):
        for num, desc in categorias[0]:
            c.drawString(1.2 * cm, y, f"- {num}: {desc}")
            y -= 0.38 * cm
            if y < 1.6 * cm:
                c.showPage()
                y = height - 1.8 * cm
    else:
        c.drawString(1.2 * cm, y, "- (nenhum)")
        y -= 0.38 * cm

    y -= 0.1 * cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.2 * cm, y, "Itens com 1 orçamento")
    y -= 0.45 * cm
    c.setFont("Helvetica", 9)
    if categorias.get(1):
        for num, desc in categorias[1]:
            c.drawString(1.2 * cm, y, f"- {num}: {desc}")
            y -= 0.38 * cm
            if y < 1.6 * cm:
                c.showPage()
                y = height - 1.8 * cm
    else:
        c.drawString(1.2 * cm, y, "- (nenhum)")
        y -= 0.38 * cm

    y -= 0.1 * cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.2 * cm, y, "Itens com 2 orçamentos")
    y -= 0.45 * cm
    c.setFont("Helvetica", 9)
    if categorias.get(2):
        for num, desc in categorias[2]:
            c.drawString(1.2 * cm, y, f"- {num}: {desc}")
            y -= 0.38 * cm
            if y < 1.6 * cm:
                c.showPage()
                y = height - 1.8 * cm
    else:
        c.drawString(1.2 * cm, y, "- (nenhum)")
        y -= 0.38 * cm

    y -= 0.1 * cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.2 * cm, y, "Itens com 3 ou mais orçamentos")
    y -= 0.45 * cm
    c.setFont("Helvetica", 9)
    if categorias.get(3):
        for num, desc in categorias[3]:
            c.drawString(1.2 * cm, y, f"- {num}: {desc}")
            y -= 0.38 * cm
            if y < 1.6 * cm:
                c.showPage()
                y = height - 1.8 * cm
    else:
        c.drawString(1.2 * cm, y, "- (nenhum)")
        y -= 0.38 * cm

    # Mantém a análise na primeira página quando houver espaço útil.
    if y < 7.0 * cm:
        c.showPage()
        y = height - 1.8 * cm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(1.2 * cm, y, "Análise de Participação dos Fornecedores")
    y -= 0.55 * cm
    c.setFont("Helvetica", 9)
    c.drawString(1.2 * cm, y, f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    y -= 0.55 * cm

    c.setFont("Helvetica-Bold", 10)
    c.drawString(1.2 * cm, y, "Resumo por fornecedor")
    y -= 0.45 * cm

    _draw_fornecedor_header(c, 1.2 * cm, y)
    y -= 0.45 * cm
    c.setFont("Helvetica", 8)

    for p in (part if part else []):
        status = []
        if p.get("enviou_orcamento"):
            status.append("orcamento")
        if p.get("recusou"):
            status.append("recusa")
        if _leitura_efetiva(p):
            status.append("leitura")
        if p.get("fez_pergunta"):
            status.append("duvida")
        if not status:
            status.append("sem resposta")

        fornecedor = _short((p.get("nome") or p.get("email") or "-"), 44)
        itens_qtd = int(itens_orcados_map.get((p.get("email") or "").lower(), 0))
        status_txt = _short(", ".join(status), 20)
        pedido_txt = _short(_fmt_data(p.get("data_pedido_enviado")), 16)
        resp_txt = _short(_fmt_data(p.get("data_primeira_resposta")), 16)
        tempo_txt = _short(_delta_horas(p.get("data_pedido_enviado"), p.get("data_primeira_resposta")), 10)

        c.drawString(1.2 * cm, y, fornecedor)
        c.drawRightString(8.0 * cm, y, str(itens_qtd))
        c.drawString(8.2 * cm, y, status_txt)
        c.drawString(11.0 * cm, y, pedido_txt)
        c.drawString(13.7 * cm, y, resp_txt)
        c.drawString(16.6 * cm, y, tempo_txt)
        y -= 0.38 * cm

        if y < 1.6 * cm:
            c.showPage()
            y = height - 1.8 * cm
            c.setFont("Helvetica-Bold", 10)
            c.drawString(1.2 * cm, y, "Análise de Participação dos Fornecedores (continuação)")
            y -= 0.5 * cm
            _draw_fornecedor_header(c, 1.2 * cm, y)
            y -= 0.45 * cm
            c.setFont("Helvetica", 8)

    c.save()
    pdf_bytes = out.getvalue()
    if not pdf_bytes:
        raise RuntimeError("PDF gerado vazio.")
    return pdf_bytes


st.set_page_config(page_title="Depurador de Orçamentos", page_icon="🤖", layout="wide")
st.title("🤖 Depurador de Orçamentos")
st.caption(
    "Sistema unificado para ingestão e depuração de orçamentos e e-mails de fornecedores "
    "(PDF/Word/Excel/CSV/ZIP/TGZ com .eml), com organização por processo e fornecedores."
)

if IMPORT_ERROR is not None:
    st.error("Falha ao iniciar o app por erro de importação.")
    with st.expander("Detalhes técnicos", expanded=True):
        st.code(IMPORT_ERROR_TRACE)
    st.stop()

api_key = _load_openrouter_api_key()

# Configuração administrada: defaults definidos pelo admin valem para todos
try:
    _cfg_admin = app_config.carregar_config()
    _config_administrada = app_config.config_administrada(_cfg_admin)
except Exception:
    _cfg_admin = {}
    _config_administrada = False
if _config_administrada:
    if "cfg_model" not in st.session_state:
        st.session_state.cfg_model = _cfg_admin.get("modelo") or "google/gemini-2.5-flash"
    if "cfg_usar_ia_juiz" not in st.session_state:
        st.session_state.cfg_usar_ia_juiz = bool(_cfg_admin.get("usar_ia_juiz", True))
    if "cfg_classificar_emails_com_ia" not in st.session_state:
        st.session_state.cfg_classificar_emails_com_ia = bool(_cfg_admin.get("classificar_emails_com_ia", True))
    try:
        import extract_utils as _eu
        if _cfg_admin.get("modelo_escalonamento"):
            _eu.ESCALATION_MODEL = _cfg_admin["modelo_escalonamento"]
    except Exception:
        pass
ocr_ok, ocr_err = True, None

if "last_result" not in st.session_state:
    st.session_state.last_result = {
        "excel_buffer": None,
        "rows_count": 0,
        "empresas_count": 0,
        "review_count": 0,
        "processo_id": None,
        "tokens_total": 0,
        "custo_total": 0.0,
    }
if "selected_processo_view_id" not in st.session_state:
    st.session_state.selected_processo_view_id = None
if "confirmar_deletar_processo" not in st.session_state:
    st.session_state.confirmar_deletar_processo = False
if "force_modo_novo" not in st.session_state:
    st.session_state.force_modo_novo = False
if "budget_db_path" not in st.session_state:
    st.session_state.budget_db_path = _default_budget_db_path()
if "forcar_reprocessamento" not in st.session_state:
    st.session_state.forcar_reprocessamento = False
if "show_sidebar_item3" not in st.session_state:
    st.session_state.show_sidebar_item3 = False
if "show_sidebar_item4" not in st.session_state:
    st.session_state.show_sidebar_item4 = False
if "cfg_model" not in st.session_state:
    st.session_state.cfg_model = "google/gemini-2.5-flash"
if "cfg_classificar_emails_com_ia" not in st.session_state:
    st.session_state.cfg_classificar_emails_com_ia = True
if "cfg_usar_ia_juiz" not in st.session_state:
    st.session_state.cfg_usar_ia_juiz = True
if "cfg_salvar_binario_anexos" not in st.session_state:
    st.session_state.cfg_salvar_binario_anexos = True
if "cfg_processar_orcamentos_de_anexos" not in st.session_state:
    st.session_state.cfg_processar_orcamentos_de_anexos = True
if "cfg_limiar_confianca_alta" not in st.session_state:
    st.session_state.cfg_limiar_confianca_alta = 85
if "cfg_limiar_confianca_baixa" not in st.session_state:
    st.session_state.cfg_limiar_confianca_baixa = 40
if "cfg_fuzzy_threshold" not in st.session_state:
    st.session_state.cfg_fuzzy_threshold = 85
if "cfg_pre_filtrar" not in st.session_state:
    st.session_state.cfg_pre_filtrar = True
if "cfg_sanity_threshold" not in st.session_state:
    st.session_state.cfg_sanity_threshold = 50
if "cfg_bloquear_numero_incoerente" not in st.session_state:
    st.session_state.cfg_bloquear_numero_incoerente = True
if "cfg_usar_uf_casamento" not in st.session_state:
    st.session_state.cfg_usar_uf_casamento = True
if "cfg_usar_qtd_casamento" not in st.session_state:
    st.session_state.cfg_usar_qtd_casamento = True

with st.sidebar:
    st.header("1. Entrada de dados")
    modo_entrada = st.radio(
        "Origem dos arquivos",
        ["Upload", "Pasta local"],
        help=(
            "Upload: envie arquivos avulsos e/ou pacotes ZIP/TGZ. "
            "Pasta local: informa um diretório contendo esses arquivos."
        ),
    )

    uploaded_files = None
    local_folder = None
    if modo_entrada == "Upload":
        uploaded_files = st.file_uploader(
            "Arquivos de entrada",
            type=[
                "pdf", "docx", "doc", "xlsx", "xls", "csv", "zip", "tgz", "gz",
            ],
            accept_multiple_files=True,
            help="Você pode enviar arquivos avulsos e arquivos compactados com .eml.",
        )
    else:
        local_folder = st.text_input(
            "Caminho da pasta",
            placeholder=r"Ex: /home/voce/entrada_orcamentos",
        )

    st.header("2. Processo")
    process_db_path = st.text_input("Banco de processos/emails", value=_default_process_db_path())
    conn_proc = process_db.get_connection(process_db_path)
    processos_existentes = process_db.listar_processos(conn_proc)
    opcoes_modo_processo = ["Novo processo"]
    if processos_existentes:
        opcoes_modo_processo.append("Processo existente")
    idx_modo = 0
    if (
        not st.session_state.force_modo_novo
        and len(opcoes_modo_processo) > 1
        and st.session_state.get("modo_processo") == "Processo existente"
    ):
        idx_modo = 1
    modo_processo = st.radio(
        "Vincular ingestão em",
        opcoes_modo_processo,
        index=idx_modo,
        key="modo_processo",
    )
    st.session_state.force_modo_novo = False

    processo_existente_id = None
    processo_numero_novo = ""
    processo_titulo_novo = ""
    auto_detectar_processo = st.checkbox(
        "Auto detectar numero do processo pelo assunto dos e-mails",
        value=True,
        help="Quando houver .eml, tenta identificar o processo automaticamente.",
    )

    if modo_processo == "Processo existente":
        if processos_existentes:
            opcoes = [f"{p['numero']} - {p.get('titulo') or 'sem titulo'}" for p in processos_existentes]
            idx_proc = st.selectbox("Selecionar", range(len(opcoes)), format_func=lambda i: opcoes[i])
            processo_existente_id = processos_existentes[idx_proc]["id"]
            st.session_state.selected_processo_view_id = processo_existente_id
        else:
            st.warning("Nenhum processo existente disponível no banco.")
            modo_processo = "Novo processo"

    if modo_processo == "Novo processo":
        processo_numero_novo = st.text_input("Numero do processo (opcional)", value="")
        processo_titulo_novo = st.text_input("Titulo/descricao do processo", value="")

    st.header("3. IA e classificação")
    model_options = ["google/gemini-2.5-flash", "deepseek/deepseek-chat-v3.2", "openai/gpt-5-mini"]

    if _config_administrada:
        st.markdown(
            "<p style='font-size:0.78rem;color:#2e7d32;font-weight:700;margin:0;'>"
            "✅ Configuração administrada ativa — nada a configurar aqui.<br>"
            f"Modelo em uso: {st.session_state.get('cfg_model', model_options[0])}</p>",
            unsafe_allow_html=True,
        )
        st.caption("Para alterar chave/modelos: aba Configurações → Administração (senha).")
        st.session_state.show_sidebar_item3 = False

    label_item3 = "Ocultar opções do item 3" if st.session_state.show_sidebar_item3 else "Mostrar opções do item 3"
    if not _config_administrada and st.button(label_item3, key="toggle_sidebar_item3", use_container_width=True):
        st.session_state.show_sidebar_item3 = not st.session_state.show_sidebar_item3

    if st.session_state.show_sidebar_item3:
        if api_key:
            st.markdown(
                "<p style='font-size:0.78rem;color:#2e7d32;font-weight:700;margin:0;'>"
                "Chave OpenRouter carregada via secrets.</p>",
                unsafe_allow_html=True,
            )
        else:
            st.warning("Chave OpenRouter não encontrada em secrets.")

        idx_model = model_options.index(st.session_state.cfg_model) if st.session_state.cfg_model in model_options else 0
        st.selectbox("Modelo", model_options, index=idx_model, key="cfg_model")
        st.checkbox(
            "IA juiz: decidir casamentos na zona cinzenta (score 60-84)",
            key="cfg_usar_ia_juiz",
            value=bool(st.session_state.get("cfg_usar_ia_juiz", True)),
            help="Pares de itens parecidos demais para ignorar e diferentes demais "
                 "para casar automaticamente são julgados por IA em lote (custo de centavos).",
        )
        st.checkbox(
            "Classificar e-mails com IA quando heurística não for suficiente",
            key="cfg_classificar_emails_com_ia",
            value=bool(st.session_state.get("cfg_classificar_emails_com_ia", True)),
        )
        st.checkbox(
            "Salvar anexos de e-mail no banco (download posterior)",
            key="cfg_salvar_binario_anexos",
            value=bool(st.session_state.get("cfg_salvar_binario_anexos", True)),
        )
        st.checkbox(
            "Processar anexos de e-mail como orçamento quando suportados",
            key="cfg_processar_orcamentos_de_anexos",
            value=bool(st.session_state.get("cfg_processar_orcamentos_de_anexos", True)),
        )
    else:
        st.caption("Opções do item 3 ocultas.")

    st.header("4. Regras de extração/comparação")
    label_item4 = "Ocultar opções do item 4" if st.session_state.show_sidebar_item4 else "Mostrar opções do item 4"
    if st.button(label_item4, key="toggle_sidebar_item4", use_container_width=True):
        st.session_state.show_sidebar_item4 = not st.session_state.show_sidebar_item4

    if st.session_state.show_sidebar_item4:
        try:
            _extrair_runtime_fn, _ocr_runtime_fn = _get_extract_runtime()
            ocr_ok, ocr_err = _ocr_runtime_fn()
        except Exception as exc:
            ocr_ok, ocr_err = False, str(exc)
        if not ocr_ok:
            st.error(f"OCR indisponível neste ambiente: {ocr_err}")

        st.slider("Limiar de confiança alta (PDF estrutural)", 60, 100, value=85, key="cfg_limiar_confianca_alta")
        st.slider("Limiar de confiança baixa (PDF estrutural)", 0, 60, value=40, key="cfg_limiar_confianca_baixa")
        st.slider("Sensibilidade do casamento por descrição", 70, 100, value=85, key="cfg_fuzzy_threshold")
        st.checkbox("Pré-filtrar texto antes de enviar à IA", value=True, key="cfg_pre_filtrar")
        st.slider("Alerta número igual x descrição divergente", 0, 80, value=50, key="cfg_sanity_threshold")
        st.checkbox(
            "Bloquear casamento automático quando número bate e descrição diverge",
            value=True,
            key="cfg_bloquear_numero_incoerente",
        )
        st.checkbox("Usar UF na identificação dos itens", value=True, key="cfg_usar_uf_casamento")
        st.checkbox("Usar quantidade na validação dos casamentos", value=True, key="cfg_usar_qtd_casamento")
    else:
        st.caption("Opções do item 4 ocultas.")

    model = st.session_state.cfg_model
    classificar_emails_com_ia = bool(st.session_state.cfg_classificar_emails_com_ia)
    salvar_binario_anexos = bool(st.session_state.cfg_salvar_binario_anexos)
    processar_orcamentos_de_anexos = bool(st.session_state.cfg_processar_orcamentos_de_anexos)
    limiar_confianca_alta = int(st.session_state.cfg_limiar_confianca_alta)
    limiar_confianca_baixa = int(st.session_state.cfg_limiar_confianca_baixa)
    fuzzy_threshold = int(st.session_state.cfg_fuzzy_threshold)
    pre_filtrar = bool(st.session_state.cfg_pre_filtrar)
    sanity_threshold = int(st.session_state.cfg_sanity_threshold)
    bloquear_numero_incoerente = bool(st.session_state.cfg_bloquear_numero_incoerente)
    usar_uf_casamento = bool(st.session_state.cfg_usar_uf_casamento)
    usar_qtd_casamento = bool(st.session_state.cfg_usar_qtd_casamento)

    st.header("5. Mapa comparativo")
    st.caption("Lista mestra opcional: colunas ITEM e NOMENCLATURA")
    master_file = st.file_uploader("Lista mestra", type=["xls", "xlsx", "csv"], label_visibility="collapsed")
    gerar_master_auto = st.checkbox("Gerar lista mestra automática por consenso", value=True)
    min_agree = st.number_input("Mínimo de orçamentos concordando", min_value=2, max_value=10, value=3)
    consensus_threshold = st.slider("Confiança mínima do consenso", min_value=60, max_value=100, value=80)

    processar = st.button("🚀 Processar ingestão completa", type="primary", use_container_width=True)


budget_db_path = st.session_state.budget_db_path
forcar_reprocessamento = bool(st.session_state.forcar_reprocessamento)

aba_pipeline, aba_emails, aba_relatorio, aba_fornecedores, aba_mapa, aba_configuracoes = st.tabs([
    "Pipeline",
    "Emails do Processo",
    "Relatório do Processo",
    "Fornecedores",
    "Mapa Comparativo",
    "Configurações",
])

processos_view = process_db.listar_processos(conn_proc)
if processos_view:
    ids_view = [p["id"] for p in processos_view]
    if modo_processo != "Novo processo" and st.session_state.selected_processo_view_id not in ids_view:
        st.session_state.selected_processo_view_id = ids_view[0]
elif modo_processo != "Novo processo":
    st.session_state.selected_processo_view_id = None


with aba_pipeline:
    processo_pipeline_id = None
    if modo_processo == "Processo existente" and processo_existente_id:
        processo_pipeline_id = processo_existente_id
    elif st.session_state.selected_processo_view_id is not None:
        processo_pipeline_id = st.session_state.selected_processo_view_id

    if processo_pipeline_id:
        proc_info = next((p for p in processos_view if p["id"] == processo_pipeline_id), None)
        consumo_acumulado = process_db.obter_consumo_processo(conn_proc, processo_pipeline_id)
        titulo_proc = (proc_info or {}).get("numero") or f"ID {processo_pipeline_id}"
        st.markdown(
            "<p style='font-size:0.80rem;color:#0d47a1;font-weight:700;margin-top:0.2rem;margin-bottom:0.5rem;'>"
            f"Consumo acumulado do processo {titulo_proc}: "
            f"{_to_int(consumo_acumulado.get('tokens_total')):,} tokens | "
            f"US$ {_to_float(consumo_acumulado.get('custo_total')):.6f}"
            "</p>",
            unsafe_allow_html=True,
        )
    else:
        st.caption("Selecione um processo para visualizar consumo acumulado de tokens e custo.")


if processar:
    if not api_key:
        st.error(
            "Chave OPENROUTER_API_KEY ausente. Configure em secrets para classificação de e-mails "
            "e extração de orçamentos com IA."
        )
        st.stop()

    if modo_entrada == "Upload" and not uploaded_files:
        st.error("Envie pelo menos um arquivo.")
        st.stop()
    if modo_entrada == "Pasta local" and not local_folder:
        st.error("Informe o caminho da pasta.")
        st.stop()

    try:
        extrair_orcamento_em_camadas, _ocr_runtime_fn = _get_extract_runtime()
    except Exception:
        st.error("Falha ao carregar o motor de extração de orçamentos neste ambiente.")
        with st.expander("Detalhes técnicos", expanded=False):
            st.code(traceback.format_exc())
        st.stop()

    with aba_pipeline:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                log_expander = st.expander("Logs de processamento", expanded=True)
                log_placeholder = log_expander.empty()
                log_lines: list[str] = []

                def add_log(msg: str):
                    stamp = time.strftime("%H:%M:%S")
                    log_lines.append(f"[{stamp}] {msg}")
                    log_placeholder.code("\n".join(log_lines[-300:]), language="text")

                add_log("Iniciando pipeline integrado.")

                # -----------------------------------------------------------
                # 1) Coleta de entradas
                # -----------------------------------------------------------
                budget_candidates: list[dict] = []
                eml_entries: list[dict] = []

                if modo_entrada == "Upload":
                    for uf in uploaded_files:
                        data = uf.getbuffer().tobytes()
                        nome = uf.name
                        hash_data = file_utils.hash_bytes(data)
                        low = nome.lower()

                        if _is_budget_file(low):
                            path = _persist_bytes(tmpdir, nome, data, prefix="upload")
                            budget_candidates.append({
                                "name": nome,
                                "path": path,
                                "modified_time": hash_data,
                                "file_id": f"upload:{nome}:{hash_data[:12]}",
                                "origem": "upload_direto",
                            })
                            add_log(f"Arquivo de orçamento detectado: {nome}")
                            continue

                        if _is_archive_file(low):
                            add_log(f"Extraindo pacote compactado: {nome}")
                            extracted = _extrair_compactado_compat(nome, data)
                            for eml in extracted["emails"]:
                                eml_entries.append({
                                    "nome": eml["nome"],
                                    "conteudo_bytes": eml["conteudo_bytes"],
                                    "arquivo_origem": nome,
                                })

                            for b in extracted["orcamentos"]:
                                h = file_utils.hash_bytes(b["conteudo_bytes"])
                                path = _persist_bytes(
                                    tmpdir,
                                    b["nome"],
                                    b["conteudo_bytes"],
                                    prefix=f"arc_{hash_data[:8]}",
                                )
                                budget_candidates.append({
                                    "name": f"{nome}/{b['nome']}",
                                    "path": path,
                                    "modified_time": h,
                                    "file_id": f"archive:{nome}:{b['nome']}:{h[:12]}",
                                    "origem": "arquivo_compactado",
                                })
                            add_log(
                                f"Pacote {nome}: {len(extracted['emails'])} e-mail(s), "
                                f"{len(extracted['orcamentos'])} orçamento(s) interno(s)."
                            )
                            continue

                        add_log(f"Ignorado (formato não suportado): {nome}")
                else:
                    if not os.path.isdir(local_folder):
                        st.error(f"Pasta não encontrada: {local_folder}")
                        st.stop()

                    for name in sorted(os.listdir(local_folder)):
                        full = os.path.join(local_folder, name)
                        if not os.path.isfile(full):
                            continue
                        low = name.lower()

                        if _is_budget_file(low):
                            budget_candidates.append({
                                "name": name,
                                "path": full,
                                "modified_time": str(os.path.getmtime(full)),
                                "file_id": full,
                                "origem": "pasta_local",
                            })
                            continue

                        if _is_archive_file(low):
                            with open(full, "rb") as fh:
                                data = fh.read()
                            add_log(f"Extraindo pacote compactado local: {name}")
                            extracted = _extrair_compactado_compat(name, data)
                            for eml in extracted["emails"]:
                                eml_entries.append({
                                    "nome": eml["nome"],
                                    "conteudo_bytes": eml["conteudo_bytes"],
                                    "arquivo_origem": name,
                                })
                            for b in extracted["orcamentos"]:
                                h = file_utils.hash_bytes(b["conteudo_bytes"])
                                path = _persist_bytes(
                                    tmpdir,
                                    b["nome"],
                                    b["conteudo_bytes"],
                                    prefix=f"arc_local_{h[:8]}",
                                )
                                budget_candidates.append({
                                    "name": f"{name}/{b['nome']}",
                                    "path": path,
                                    "modified_time": h,
                                    "file_id": f"archive_local:{name}:{b['nome']}:{h[:12]}",
                                    "origem": "arquivo_compactado_local",
                                })

                if not budget_candidates and not eml_entries:
                    st.warning("Nenhum arquivo processável encontrado.")
                    st.stop()

                add_log(
                    f"Entrada consolidada: {len(eml_entries)} e-mail(s) e "
                    f"{len(budget_candidates)} arquivo(s) de orçamento."
                )

                # -----------------------------------------------------------
                # 2) Resolve processo para persistir e-mails
                # -----------------------------------------------------------
                processo_id = None
                processo_numero = None
                processo_criado_agora = False

                if modo_processo == "Processo existente":
                    processo_id = processo_existente_id
                    proc = next((p for p in processos_existentes if p["id"] == processo_id), None)
                    processo_numero = proc["numero"] if proc else None
                    add_log(f"Usando processo existente: {processo_numero}")
                elif processo_numero_novo.strip():
                    numero = processo_numero_novo.strip()
                    existe = process_db.buscar_processo_por_numero(conn_proc, numero)
                    if existe:
                        processo_id = existe["id"]
                        processo_numero = existe["numero"]
                        add_log(f"Processo já existente encontrado: {processo_numero}")
                    else:
                        processo_id = process_db.criar_processo(
                            conn_proc, numero, processo_titulo_novo.strip(), ""
                        )
                        processo_numero = numero
                        processo_criado_agora = True
                        add_log(f"Processo criado: {processo_numero}")

                # -----------------------------------------------------------
                # 3) Processa e-mails (.eml) e captura anexos de orçamento
                # -----------------------------------------------------------
                n_emails_novos = 0
                n_emails_dup = 0
                n_emails_ia = 0
                n_pedidos_inferidos = 0
                n_cnpjs_identificados = 0
                n_telefones_identificados = 0
                custo_emails = 0.0
                tokens_emails = 0
                n_anexos_total = 0
                n_anexos_orc_aprovados = 0
                n_anexos_filtrados = 0
                n_anexos_duplicados = 0
                # Chave: (hash_conteudo, remetente_email) — evita processar o mesmo
                # arquivo do mesmo remetente duas vezes, mas permite arquivos com
                # bytes idênticos vindos de fornecedores diferentes.
                anexos_orc_hashes: set[tuple[str, str]] = set()

                if eml_entries:
                    add_log("Iniciando processamento de e-mails.")

                    # --- Pré-passagem: hashes dos anexos institucionais -------
                    # Lição de campo: fornecedores às vezes devolvem o próprio
                    # arquivo-modelo do pedido de cotação sem preencher preço.
                    # Hash de conteúdo igual ao do anexo enviado pelo órgão é o
                    # sinal inequívoco de "template em branco, não é orçamento",
                    # mesmo que o nome do arquivo pareça uma proposta.
                    _REMETENTES_INSTITUCIONAIS_PRE = {"sobressalentes.comrj@gmail.com"}
                    _RE_INSTITUCIONAL = re.compile(r"\.mil\.br$|\.mar\.mil\.br$|\.marinha\.mil\.br$")
                    parsed_emls = []
                    template_hashes: set[str] = set()
                    for eml in eml_entries:
                        _parsed_pre = email_utils.parse_eml(eml["conteudo_bytes"], filename=eml["nome"])
                        parsed_emls.append(_parsed_pre)
                        _rem_pre = (_parsed_pre.get("remetente_email") or "").lower()
                        if _RE_INSTITUCIONAL.search(_rem_pre) or _rem_pre in _REMETENTES_INSTITUCIONAIS_PRE:
                            for _anx_pre in _parsed_pre.get("anexos", []):
                                try:
                                    template_hashes.add(file_utils.hash_bytes(_anx_pre["conteudo_bytes"]))
                                except Exception:
                                    pass
                    if template_hashes:
                        add_log(
                            f"Pré-passagem: {len(template_hashes)} hash(es) de template institucional registrados."
                        )

                    for parsed in parsed_emls:
                        rem_email = parsed.get("remetente_email", "")
                        rem_nome = parsed.get("remetente_nome", "")

                        # Descobre processo automaticamente se ainda não estiver definido
                        if processo_id is None:
                            detected = None
                            if auto_detectar_processo:
                                detected = email_utils.extrair_numero_processo(parsed.get("assunto", ""))
                            if not detected:
                                detected = f"SEM_NUMERO_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                            existe = process_db.buscar_processo_por_numero(conn_proc, detected)
                            if existe:
                                processo_id = existe["id"]
                                processo_numero = existe["numero"]
                                add_log(f"Processo detectado automaticamente: {processo_numero}")
                            else:
                                processo_id = process_db.criar_processo(
                                    conn_proc,
                                    detected,
                                    processo_titulo_novo.strip() or "Processo detectado via e-mail",
                                    "",
                                )
                                processo_numero = detected
                                processo_criado_agora = True
                                add_log(f"Processo criado automaticamente: {processo_numero}")

                        if process_db.email_ja_importado(conn_proc, processo_id, parsed["message_id"]):
                            n_emails_dup += 1
                            continue

                        if classificar_emails_com_ia:
                            clf = email_classifier.classificar_email(parsed, api_key, model)
                        else:
                            tipo_heur = email_classifier._heuristica(parsed) or "outro"
                            clf = {
                                "tipo": tipo_heur,
                                "confianca": 80,
                                "resumo": "",
                                "numero_processo": None,
                                "uso_ia": False,
                                "usage": {},
                            }

                        if clf.get("uso_ia"):
                            n_emails_ia += 1
                            usage = clf.get("usage") or {}
                            tokens_emails += _to_int(usage.get("total_tokens"))
                            custo_emails += _to_float(usage.get("cost_usd"))

                        email_id = process_db.salvar_email(
                            conn_proc,
                            processo_id,
                            parsed,
                            clf.get("tipo") or "outro",
                            int(clf.get("confianca") or 0),
                            clf.get("resumo") or "",
                        )
                        n_emails_novos += 1

                        tipo_email = clf.get("tipo") or "outro"

                        # Anexos binários no banco
                        melhor_anexo = None
                        melhor_pontuacao = -10_000
                        for anx in parsed.get("anexos", []):
                            n_anexos_total += 1
                            if salvar_binario_anexos:
                                process_db.salvar_anexo(
                                    conn_proc,
                                    email_id,
                                    anx["nome"],
                                    anx["conteudo_bytes"],
                                    anx.get("tipo_mime") or "application/octet-stream",
                                )
                            if processar_orcamentos_de_anexos:
                                pontuacao, motivo_anx = _score_email_attachment_candidate(
                                    anx.get("nome", ""),
                                    anx.get("tipo_mime") or "",
                                    tipo_email,
                                )
                                if pontuacao > 0 and pontuacao > melhor_pontuacao:
                                    melhor_pontuacao = pontuacao
                                    melhor_anexo = anx

                        # Preferência explícita por PDF quando há outro anexo PDF
                        # com "cotação" no nome e o melhor selecionado não é PDF.
                        if processar_orcamentos_de_anexos and melhor_anexo is not None:
                            nome_melhor = (melhor_anexo.get("nome") or "").lower()
                            if not nome_melhor.endswith(".pdf") and "cota" in nome_melhor:
                                for anx in parsed.get("anexos", []):
                                    nome_anx = (anx.get("nome") or "").lower()
                                    if nome_anx.endswith(".pdf") and "cota" in nome_anx:
                                        p, _ = _score_email_attachment_candidate(
                                            anx.get("nome", ""), anx.get("tipo_mime") or "", tipo_email
                                        )
                                        if p > 0:
                                            melhor_anexo = anx
                                            break

                        if processar_orcamentos_de_anexos:
                            if melhor_anexo is None:
                                n_anexos_filtrados += len(parsed.get("anexos", []))
                            else:
                                n_anexos_filtrados += max(0, len(parsed.get("anexos", [])) - 1)
                                h = file_utils.hash_bytes(melhor_anexo["conteudo_bytes"])
                                dedup_key = (h, (rem_email or "").lower())
                                if dedup_key in anexos_orc_hashes:
                                    n_anexos_duplicados += 1
                                    add_log(
                                        f"Anexo ignorado (mesmo arquivo já processado deste remetente): {melhor_anexo['nome']}"
                                    )
                                else:
                                    anexos_orc_hashes.add(dedup_key)
                                    path = _persist_bytes(
                                        tmpdir,
                                        melhor_anexo["nome"],
                                        melhor_anexo["conteudo_bytes"],
                                        prefix=f"email_{email_id}",
                                    )
                                    # Se o remetente aparenta ser a Marinha / domínio .mil.br
                                    # ou for endereço institucional conhecido,
                                    # não trate o anexo como orçamento de fornecedor —
                                    # é possivelmente o pedido enviado às empresas.
                                    _REMETENTES_INSTITUCIONAIS = {"sobressalentes.comrj@gmail.com"}
                                    _rem_lower = (rem_email or "").lower()
                                    if re.search(r"\.mil\.br$|\.mar\.mil\.br$|\.marinha\.mil\.br$", _rem_lower) or _rem_lower in _REMETENTES_INSTITUCIONAIS:
                                        add_log(
                                            f"Anexo ignorado (remetente institucional): {melhor_anexo['nome']} — {rem_email}"
                                        )
                                    elif h in template_hashes:
                                        # Lição de campo: mesmo hash do template do órgão
                                        # = arquivo devolvido em branco, não é orçamento.
                                        n_anexos_filtrados += 1
                                        add_log(
                                            f"Anexo ignorado (template em branco do pedido, devolvido sem preços): "
                                            f"{melhor_anexo['nome']} — {rem_email}"
                                        )
                                    else:
                                        budget_candidates.append({
                                            "name": f"email:{parsed.get('arquivo_eml_nome')}/{melhor_anexo['nome']}",
                                            "path": path,
                                            "modified_time": h,
                                            "file_id": f"email_attachment:{email_id}:{melhor_anexo['nome']}:{h[:12]}",
                                            "origem": "anexo_email",
                                            "fornecedor_email_hint": rem_email,
                                            "fornecedor_nome_hint": rem_nome,
                                        })
                                        n_anexos_orc_aprovados += 1
                                        rem_label = rem_nome or rem_email or "remetente desconhecido"
                                        add_log(
                                            f"Anexo da {rem_label} <{rem_email}>: {melhor_anexo['nome']}"
                                        )

                        # Atualiza participação dos fornecedores
                        rem_cnpj = normalize_utils.extrair_cnpj(parsed.get("corpo") or "")
                        rem_telefone = normalize_utils.extrair_telefone(parsed.get("corpo") or "")
                        data_pedido_hist = None
                        if tipo_email in {"orcamento_recebido", "declinio", "duvida"}:
                            data_pedido_hist = email_utils.extrair_data_pedido_do_historico(
                                parsed.get("corpo") or "",
                                parsed.get("data_envio"),
                            )
                            if data_pedido_hist:
                                n_pedidos_inferidos += 1
                        if rem_email:
                            if rem_cnpj:
                                n_cnpjs_identificados += 1
                            if rem_telefone:
                                n_telefones_identificados += 1
                            fid = process_db.upsert_fornecedor(
                                conn_proc,
                                rem_email,
                                rem_nome,
                                rem_cnpj or "",
                                rem_telefone or "",
                            )
                            process_db.atualizar_participacao(
                                conn_proc,
                                processo_id,
                                fid,
                                tipo_email,
                                parsed.get("data_envio"),
                                data_pedido_hist,
                            )

                        if tipo_email == "pedido_orcamento":
                            for dest in parsed.get("destinatarios", []):
                                de = (dest.get("email") or "").strip()
                                dn = dest.get("nome") or ""
                                if de:
                                    fid_dest = process_db.upsert_fornecedor(conn_proc, de, dn)
                                    process_db.atualizar_participacao(
                                        conn_proc,
                                        processo_id,
                                        fid_dest,
                                        "pedido_orcamento",
                                        parsed.get("data_envio"),
                                    )

                add_log(
                    f"E-mails: {n_emails_novos} novo(s), {n_emails_dup} duplicado(s), "
                    f"{n_emails_ia} com IA."
                )
                if processar_orcamentos_de_anexos and n_anexos_total:
                    add_log(
                        "Anexos de e-mail: "
                        f"{n_anexos_total} total, "
                        f"{n_anexos_orc_aprovados} candidato(s) selecionado(s), "
                        f"{n_anexos_filtrados} filtrado(s), "
                        f"{n_anexos_duplicados} duplicado(s) por hash."
                    )
                if n_pedidos_inferidos:
                    add_log(
                        f"Histórico detectado em respostas: {n_pedidos_inferidos} data(s) de pedido inferida(s)."
                    )
                if n_cnpjs_identificados:
                    add_log(f"CNPJ identificado em e-mails de fornecedores: {n_cnpjs_identificados} ocorrência(s).")
                if n_telefones_identificados:
                    add_log(
                        f"Telefone identificado em e-mails de fornecedores: {n_telefones_identificados} ocorrência(s)."
                    )
                if processo_id:
                    n_cnpjs_backfill = _atualizar_cnpjs_fornecedores_compat(conn_proc, processo_id)
                    if n_cnpjs_backfill:
                        add_log(f"CNPJ atualizado em fornecedores já cadastrados: {n_cnpjs_backfill} fornecedor(es).")

                # -----------------------------------------------------------
                # 4) Processa arquivos de orçamento
                # -----------------------------------------------------------
                conn_budget = db_utils.get_connection(budget_db_path)
                all_extractions = []
                sources = []
                review_extracao = []
                falhas = []
                avisos_truncamento = []
                diagnostico_arquivos = []

                n_budget_cache = 0
                n_budget_novos = 0
                tokens_orc = 0
                custo_orc = 0.0
                prompt_tokens_total = 0
                completion_tokens_total = 0
                arquivos_com_ia = 0
                custo_estimado = False
                n_fornecedores_enriquecidos_por_orc = 0

                progress = st.progress(0)
                status = st.empty()

                for i, f in enumerate(budget_candidates):
                    status.text(f"Processando orçamento {i + 1}/{len(budget_candidates)}: {f['name']}")
                    # Mostra remetente quando disponível para facilitar rastreabilidade
                    fornecedor_email_hint = (f.get("fornecedor_email_hint") or "").strip().lower()
                    remetente_info = fornecedor_email_hint or f.get("origem") or "desconhecido"
                    add_log(f"Iniciando processamento do arquivo: {f['name']} (remetente: {remetente_info})")
                    file_id = f["file_id"]
                    modified_time = f["modified_time"]

                    cached = None if forcar_reprocessamento else db_utils.get_cached_file(conn_budget, file_id)
                    cached_sem_dados_fornecedor = bool(
                        cached
                        and not (cached.get("cnpj") or "").strip()
                        and not (cached.get("telefone") or "").strip()
                    )
                    if cached_sem_dados_fornecedor:
                        add_log(
                            f"{f['name']}: cache antigo sem CNPJ/telefone, reextraindo para enriquecer fornecedor."
                        )
                        cached = None
                    if (
                        cached
                        and cached["modified_time"] == modified_time
                        and cached.get("extraction_version") == db_utils.EXTRACTION_VERSION
                    ):
                        itens = db_utils.get_items_for_file(conn_budget, file_id)
                        empresa = cached["empresa"]
                        cnpj_orc = (cached.get("cnpj") or "").strip()
                        telefone_orc = (cached.get("telefone") or "").strip()
                        fornecedor_email_hint = (f.get("fornecedor_email_hint") or "").strip().lower()
                        remetente_info = fornecedor_email_hint or f.get("origem") or "desconhecido"
                        add_log(f"{f['name']}: já processado anteriormente, usando cache local (remetente: {remetente_info}).")
                        all_extractions.append({"empresa": empresa, "arquivo": f["name"], "itens": itens})
                        if processo_id:
                            fornecedor_email_hint = (f.get("fornecedor_email_hint") or "").strip().lower()
                            fornecedor_nome_hint = (f.get("fornecedor_nome_hint") or "").strip()
                            process_db.vincular_orcamento_ao_processo(
                                conn_proc, processo_id, file_id, f["name"],
                                remetente_email=fornecedor_email_hint,
                            )
                            if fornecedor_email_hint and (cnpj_orc or telefone_orc):
                                process_db.upsert_fornecedor(
                                    conn_proc,
                                    fornecedor_email_hint,
                                    fornecedor_nome_hint,
                                    cnpj_orc,
                                    telefone_orc,
                                )
                                n_fornecedores_enriquecidos_por_orc += 1
                            if cnpj_orc or telefone_orc:
                                n_fornecedores_enriquecidos_por_orc += _atualizar_fornecedor_dados_por_nome_compat(
                                    conn_proc,
                                    processo_id,
                                    empresa,
                                    cnpj_orc,
                                    telefone_orc,
                                )
                        for item in itens:
                            sources.append({
                                "empresa": empresa,
                                "arquivo": f["name"],
                                "fonte_extracao": item.get("fonte_extracao", "cache"),
                                "origem": item.get("origem", f.get("origem", "desconhecida")),
                                "numero_item": item.get("numero_item"),
                                "descricao": item.get("descricao"),
                            })
                        n_budget_cache += 1
                        progress.progress((i + 1) / max(1, len(budget_candidates)))
                        continue

                    try:
                        result = extrair_orcamento_em_camadas(
                            path=f["path"],
                            api_key=api_key,
                            model=model,
                            pre_filtrar=pre_filtrar,
                            limiar_alto=int(limiar_confianca_alta),
                            limiar_baixo=int(limiar_confianca_baixa),
                        )
                        if result.get("erro") and not result.get("itens"):
                            falhas.append(f"{f['name']}: {result['erro']}")
                            progress.progress((i + 1) / max(1, len(budget_candidates)))
                            continue

                        empresa = result.get("empresa") or os.path.splitext(os.path.basename(f["name"]))[0]
                        cnpj_orc = (result.get("cnpj") or "").strip()
                        telefone_orc = (result.get("telefone") or "").strip()
                        itens = result.get("itens", [])
                        if result.get("texto_truncado"):
                            avisos_truncamento.append(f["name"])
                        all_extractions.append({"empresa": empresa, "arquivo": f["name"], "itens": itens})
                        review_extracao.extend(result.get("review", []))

                        usage = result.get("usage") or {}
                        prompt_tokens_total += _to_int(usage.get("prompt_tokens"))
                        completion_tokens_total += _to_int(usage.get("completion_tokens"))
                        tokens_orc += _to_int(usage.get("total_tokens"))
                        custo_orc += _to_float(usage.get("cost_usd"))
                        custo_estimado = custo_estimado or bool(usage.get("estimated"))
                        if _to_int(usage.get("total_tokens")) > 0:
                            arquivos_com_ia += 1

                        diagnostico_arquivos.append({
                            "arquivo": f["name"],
                            "empresa": empresa,
                            "itens": len(itens),
                            "fonte": result.get("fonte_processamento", "ia"),
                            "confianca": result.get("confianca_estrutural"),
                        })

                        for item in itens:
                            sources.append({
                                "empresa": empresa,
                                "arquivo": f["name"],
                                "fonte_extracao": item.get("fonte_extracao", result.get("fonte_processamento", "ia")),
                                "origem": item.get("origem", f.get("origem", "desconhecida")),
                                "numero_item": item.get("numero_item"),
                                "descricao": item.get("descricao"),
                            })

                        db_utils.save_extraction(
                            conn_budget,
                            file_id,
                            f["name"],
                            empresa,
                            modified_time,
                            itens,
                            cnpj_orc,
                            telefone_orc,
                        )
                        if processo_id:
                            fornecedor_email_hint = (f.get("fornecedor_email_hint") or "").strip().lower()
                            fornecedor_nome_hint = (f.get("fornecedor_nome_hint") or "").strip()
                            process_db.vincular_orcamento_ao_processo(
                                conn_proc, processo_id, file_id, f["name"],
                                remetente_email=fornecedor_email_hint,
                            )
                            if fornecedor_email_hint and (cnpj_orc or telefone_orc):
                                process_db.upsert_fornecedor(
                                    conn_proc,
                                    fornecedor_email_hint,
                                    fornecedor_nome_hint,
                                    cnpj_orc,
                                    telefone_orc,
                                )
                                n_fornecedores_enriquecidos_por_orc += 1
                            if cnpj_orc or telefone_orc:
                                n_fornecedores_enriquecidos_por_orc += _atualizar_fornecedor_dados_por_nome_compat(
                                    conn_proc,
                                    processo_id,
                                    empresa,
                                    cnpj_orc,
                                    telefone_orc,
                                )
                        n_budget_novos += 1
                    except Exception as exc:
                        falhas.append(f"{f['name']}: {exc}")

                    progress.progress((i + 1) / max(1, len(budget_candidates)))

                # -----------------------------------------------------------
                # 5) Gera mapa comparativo (se houver orçamentos)
                # -----------------------------------------------------------
                master_items = _ler_master_items(master_file) if master_file else None
                rows = []
                review = []
                excel_buffer = None
                empresas = []

                if all_extractions:
                    if master_items is None and gerar_master_auto:
                        auto_master = build_master_from_consensus(
                            all_extractions,
                            min_agree=int(min_agree),
                            consensus_threshold=int(consensus_threshold),
                        )
                        if auto_master:
                            master_items = auto_master
                            add_log(f"Lista mestra por consenso gerada com {len(auto_master)} item(ns).")

                    try:
                        conn_learn = learning_db.get_connection()
                        correcoes_salvas = learning_db.carregar_correcoes(conn_learn)
                    except Exception:
                        correcoes_salvas = {}
                    if correcoes_salvas:
                        add_log(f"Memória de correções: {len(correcoes_salvas)} decisão(ões) salva(s) aplicada(s) ao matching.")

                    rows, _row_index, matrix, review = build_comparison_table(
                        all_extractions,
                        master_items=master_items,
                        fuzzy_threshold=int(fuzzy_threshold),
                        sanity_threshold=int(sanity_threshold),
                        usar_uf=bool(usar_uf_casamento),
                        usar_qtd=bool(usar_qtd_casamento),
                        bloquear_numero_incoerente=bool(bloquear_numero_incoerente),
                        correcoes=correcoes_salvas,
                    )
                    review.extend(review_extracao)
                    empresas = sorted({e["empresa"] for e in all_extractions})

                    # --- IA juiz: zona cinzenta do casamento -----------------
                    usar_ia_juiz = bool(st.session_state.get("cfg_usar_ia_juiz", True))
                    if usar_ia_juiz and api_key:
                        pares_cinzentos = encontrar_pares_zona_cinzenta(
                            rows, matrix,
                            fuzzy_threshold=int(fuzzy_threshold),
                            zona_min=60,
                            correcoes=correcoes_salvas,
                        )
                        if pares_cinzentos:
                            add_log(f"IA juiz: {len(pares_cinzentos)} par(es) na zona cinzenta enviados para julgamento.")
                            decisoes_juiz, usage_juiz = ai_judge.julgar_pares(pares_cinzentos, api_key, model)
                            rows, n_fusoes = fundir_linhas_julgadas(rows, matrix, review, pares_cinzentos, decisoes_juiz)
                            tokens_orc += _to_int(usage_juiz.get("total_tokens"))
                            custo_orc += _to_float(usage_juiz.get("cost_usd"))
                            add_log(
                                f"IA juiz: {n_fusoes} fusão(ões) aplicadas em "
                                f"{usage_juiz.get('chamadas', 0)} chamada(s)."
                            )

                    item_cobertura = _build_item_cobertura(rows, matrix, base_meta=3)

                    # --- Autoverificação pré-entrega --------------------------
                    try:
                        avisos_sanidade = sanity_check.verificar_mapa(rows, matrix, empresas, all_extractions)
                    except Exception as _exc_san:
                        avisos_sanidade = [f"Autoverificação falhou: {_exc_san}"]
                    for aviso in avisos_sanidade:
                        add_log(f"⚠️ {aviso}")
                    if avisos_sanidade:
                        st.warning(
                            "Autoverificação encontrou "
                            f"{len(avisos_sanidade)} ponto(s) de atenção — veja o log da execução."
                        )

                    if rows:
                        excel_buffer = build_excel(rows, matrix, empresas, review, sources)
                else:
                    item_cobertura = []

                tokens_total = tokens_emails + tokens_orc
                custo_total = custo_emails + custo_orc
                acumulado_processo = None
                if processo_id:
                    process_db.registrar_consumo_processo(
                        conn_proc,
                        processo_id,
                        tokens_total=tokens_total,
                        custo_total=custo_total,
                        prompt_tokens_total=prompt_tokens_total,
                        completion_tokens_total=completion_tokens_total,
                        tokens_emails_total=tokens_emails,
                        tokens_orcamentos_total=tokens_orc,
                    )
                    acumulado_processo = process_db.obter_consumo_processo(conn_proc, processo_id)
                st.session_state.last_result = {
                    "excel_buffer": excel_buffer,
                    "rows_count": len(rows),
                    "empresas_count": len(empresas),
                    "review_count": len(review),
                    "review": review,
                    "item_cobertura": item_cobertura,
                    "processo_id": processo_id,
                    "tokens_total": tokens_total,
                    "custo_total": custo_total,
                }
                if processo_id:
                    st.session_state.selected_processo_view_id = processo_id

                # -----------------------------------------------------------
                # 6) Resumo da execução
                # -----------------------------------------------------------
                st.success("Processamento concluído.")
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("E-mails novos", n_emails_novos)
                c2.metric("Orçamentos processados", n_budget_novos)
                c3.metric("Orçamentos do cache", n_budget_cache)
                c4.metric("Itens no mapa", len(rows))
                c5.metric("Empresas no mapa", len(empresas))

                suo = " (estimado)" if custo_estimado else ""
                st.markdown(
                    "<p style='font-size:0.80rem;color:#b71c1c;font-weight:700;margin-top:0.45rem;'>"
                    f"Consumo IA total: {tokens_total:,} tokens | "
                    f"Gasto total: US$ {custo_total:.6f}{suo} | "
                    f"Orçamentos IA: {tokens_orc:,} tokens (entrada {prompt_tokens_total:,} | "
                    f"saída {completion_tokens_total:,})"
                    "</p>",
                    unsafe_allow_html=True,
                )

                if processo_numero:
                    st.info(f"Processo da execução: {processo_numero}")
                if n_fornecedores_enriquecidos_por_orc:
                    st.info(
                        "Fornecedores enriquecidos por dados dos orçamentos "
                        f"(CNPJ/telefone): {n_fornecedores_enriquecidos_por_orc}."
                    )
                if acumulado_processo is not None:
                    st.info(
                        "Consumo acumulado deste processo: "
                        f"{_to_int(acumulado_processo.get('tokens_total')):,} tokens | "
                        f"US$ {_to_float(acumulado_processo.get('custo_total')):.6f} "
                        f"em {_to_int(acumulado_processo.get('executions_count'))} execução(ões)."
                    )

                if diagnostico_arquivos:
                    with st.expander("Diagnóstico da extração de orçamentos", expanded=False):
                        for info in diagnostico_arquivos:
                            st.write(
                                f"- {info['arquivo']}: {info['itens']} item(ns), "
                                f"fonte={info['fonte']}, confiança={info['confianca']}, empresa={info['empresa']}"
                            )

                if falhas:
                    with st.expander(f"{len(falhas)} arquivo(s) com problema"):
                        for msg in falhas:
                            st.write(f"- {msg}")

                if avisos_truncamento:
                    with st.expander(f"{len(avisos_truncamento)} arquivo(s) com texto truncado"):
                        for nome in avisos_truncamento:
                            st.write(f"- {nome}")
        except Exception:
            st.error("Falha não tratada no pipeline.")
            with st.expander("Traceback técnico", expanded=True):
                st.code(traceback.format_exc())
            # Limpa processo vazio criado antes da falha
            if processo_id is not None and processo_criado_agora:
                try:
                    emails_salvos = process_db.listar_emails_processo(conn_proc, processo_id)
                    orcamentos_salvos = process_db.listar_orcamentos_do_processo(conn_proc, processo_id)
                    if not emails_salvos and not orcamentos_salvos:
                        process_db.deletar_processo(conn_proc, processo_id)
                        add_log(f"Processo {processo_numero} removido (nenhum dado processado).")
                except Exception:
                    pass


with aba_emails:
    st.subheader("Emails por processo")
    processo_sel = None
    if st.session_state.selected_processo_view_id is not None:
        processo_sel = next(
            (p for p in process_db.listar_processos(conn_proc)
             if p["id"] == st.session_state.selected_processo_view_id),
            None,
        )
    if processo_sel:
        st.caption(f"Processo ativo: {processo_sel['numero']} - {processo_sel.get('titulo') or 'sem titulo'}")
    else:
        st.info("Selecione um processo na sidebar para visualizar esta aba.")

    if processo_sel:
        emails = process_db.listar_emails_processo(conn_proc, processo_sel["id"])
        st.caption(f"Total de e-mails: {len(emails)}")
        if not emails:
            st.info("Sem e-mails neste processo.")
        else:
            tipos_disponiveis = sorted({e["tipo"] for e in emails if e.get("tipo")})
            filtro_tipo = st.multiselect(
                "Filtrar por tipo",
                options=tipos_disponiveis,
                default=tipos_disponiveis,
                format_func=email_classifier.rotulo_tipo,
            )
            busca = st.text_input("Buscar em assunto/remetente", "")

            emails_filtrados = [
                e for e in emails
                if (not filtro_tipo or e["tipo"] in filtro_tipo)
                and (
                    not busca
                    or busca.lower() in (e.get("assunto") or "").lower()
                    or busca.lower() in (e.get("remetente_email") or "").lower()
                    or busca.lower() in (e.get("remetente_nome") or "").lower()
                )
            ]
            st.caption(f"Exibindo {len(emails_filtrados)} e-mail(s) filtrado(s).")

            for e in emails_filtrados:
                anexos = e.get("nomes_anexos") or []
                anexos_txt = f" | Anexos: {', '.join(anexos)}" if anexos else ""
                header = (
                    f"{_fmt_data(e.get('data_envio'))} | "
                    f"{e.get('remetente_nome') or e.get('remetente_email') or '?'} | "
                    f"{e.get('assunto') or '(sem assunto)'}{anexos_txt}"
                )
                with st.expander(header):
                    st.write(f"Tipo: {email_classifier.rotulo_tipo(e['tipo'])}")
                    st.write(f"Confiança: {e.get('confianca_tipo') or 0}%")
                    if e.get("resumo"):
                        st.write(f"Resumo: {e['resumo']}")

                    dests = e.get("destinatarios") or []
                    if dests:
                        st.write(
                            "Destinatários: " + "; ".join(
                                f"{d.get('nome') or ''} <{d.get('email') or ''}>".strip() for d in dests[:8]
                            )
                        )

                    corpo = process_db.get_email_corpo(conn_proc, e["id"])
                    st.text_area(
                        "Corpo",
                        value=corpo[:3000] + ("..." if len(corpo) > 3000 else ""),
                        height=220,
                        disabled=True,
                        key=f"body_{e['id']}",
                    )

                    anexos_db = process_db.listar_anexos_email(conn_proc, e["id"])
                    if anexos_db:
                        st.write("Downloads de anexos:")
                        for anx in anexos_db:
                            payload = process_db.get_anexo_conteudo(conn_proc, anx["id"])
                            if payload:
                                st.download_button(
                                    label=f"Baixar {anx['nome']}",
                                    data=payload,
                                    file_name=anx["nome"],
                                    mime=anx.get("tipo_mime") or "application/octet-stream",
                                    key=f"dl_{anx['id']}",
                                )


with aba_relatorio:
    st.subheader("Relatório consolidado por processo")
    processo_sel = None
    if st.session_state.selected_processo_view_id is not None:
        processo_sel = next(
            (p for p in process_db.listar_processos(conn_proc)
             if p["id"] == st.session_state.selected_processo_view_id),
            None,
        )
    if processo_sel:
        st.caption(f"Processo ativo: {processo_sel['numero']} - {processo_sel.get('titulo') or 'sem titulo'}")
    else:
        st.info("Selecione um processo na sidebar para visualizar esta aba.")

    if processo_sel:
        resumo = process_db.get_resumo_processo(conn_proc, processo_sel["id"])
        conn_budget_rel = db_utils.get_connection(budget_db_path)
        itens_orcados_map = _itens_orcados_por_fornecedor(conn_proc, conn_budget_rel, processo_sel["id"])
        resumo["itens_orcados_por_fornecedor"] = itens_orcados_map

        item_cobertura_rel = []
        lr_rel = st.session_state.last_result or {}
        if lr_rel.get("processo_id") == processo_sel["id"] and lr_rel.get("item_cobertura"):
            item_cobertura_rel = lr_rel.get("item_cobertura") or []
        else:
            rebuilt_rel = _rebuild_map_for_process(
                conn_proc,
                conn_budget_rel,
                processo_sel["id"],
                fuzzy_threshold,
                sanity_threshold,
                usar_uf_casamento,
                usar_qtd_casamento,
                bloquear_numero_incoerente,
                gerar_master_auto,
                min_agree,
                consensus_threshold,
            )
            if rebuilt_rel:
                item_cobertura_rel = rebuilt_rel.get("item_cobertura") or []
        resumo["item_cobertura"] = item_cobertura_rel
        cobertura_resumo = _resumo_cobertura_itens(item_cobertura_rel)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total e-mails", resumo["total_emails"])
        c2.metric("Fornecedores", resumo["total_fornecedores"])
        c3.metric("Enviaram orçamento", resumo["enviaram_orcamento"])
        c4.metric("Recusaram", resumo["recusaram"])
        c5.metric("Sem resposta", resumo["sem_resposta"])

        st.markdown("### Indicadores de Itens")
        i1, i2, i3, i4 = st.columns(4)
        i1.metric("Quantidade total de itens", int(cobertura_resumo["total_itens"]))
        i2.metric(
            "Itens com zero orçamento",
            f"{int(cobertura_resumo['itens_zero'])} ({float(cobertura_resumo['perc_zero']):.1f}%)",
        )
        i3.metric(
            "Itens com 1 a 2 orçamentos",
            f"{int(cobertura_resumo['itens_1_a_2'])} ({float(cobertura_resumo['perc_1_a_2']):.1f}%)",
        )
        i4.metric(
            "Itens com 3 ou mais orçamentos",
            f"{int(cobertura_resumo['itens_3_ou_mais'])} ({float(cobertura_resumo['perc_3_ou_mais']):.1f}%)",
        )

        if resumo.get("tempo_medio_resposta_h") is not None:
            h = float(resumo["tempo_medio_resposta_h"])
            st.metric("Tempo médio de resposta", f"{h:.1f}h ({h / 24.0:.2f} dias)")

        por_tipo = resumo.get("por_tipo") or {}
        if por_tipo:
            rows_tipo = [
                {"Tipo": email_classifier.rotulo_tipo(k), "Quantidade": v}
                for k, v in sorted(por_tipo.items(), key=lambda x: -x[1])
            ]
            _safe_dataframe(rows_tipo)

        participacoes = resumo.get("participacoes") or []
        if participacoes:
            rows_part = []
            for p in participacoes:
                leitura_ok = _leitura_efetiva(p)
                itens_orcados = itens_orcados_map.get((p.get("email") or "").lower(), 0)
                rows_part.append({
                    "Fornecedor": p.get("nome") or p.get("email"),
                    "Email": p.get("email"),
                    "Itens orçados": int(itens_orcados),
                    "Orçamento": "SIM" if p.get("enviou_orcamento") else "NAO",
                    "Recusa": "SIM" if p.get("recusou") else "NAO",
                    "Leitura": "SIM" if leitura_ok else "NAO",
                    "Dúvida": "SIM" if p.get("fez_pergunta") else "NAO",
                    "Pedido enviado": _fmt_data(p.get("data_pedido_enviado")),
                    "Primeira resposta": _fmt_data(p.get("data_primeira_resposta")),
                    "Tempo resposta": _delta_horas(
                        p.get("data_pedido_enviado"),
                        p.get("data_primeira_resposta"),
                    ),
                })
            _safe_dataframe(rows_part)

        if st.button("Gerar relatório detalhado (.txt)", key="btn_relatorio_txt"):
            linhas = [
                f"RELATORIO DO PROCESSO: {processo_sel['numero']}",
                f"Titulo: {processo_sel.get('titulo') or '-'}",
                f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                "=" * 70,
                "",
                f"Total de emails: {resumo['total_emails']}",
                "Distribuicao por tipo:",
            ]
            for tipo, qtd in sorted((resumo.get("por_tipo") or {}).items(), key=lambda x: -x[1]):
                linhas.append(f"- {email_classifier.rotulo_tipo(tipo)}: {qtd}")
            linhas.extend([
                "",
                f"Fornecedores contatados: {resumo['total_fornecedores']}",
                f"Enviaram orcamento: {resumo['enviaram_orcamento']}",
                f"Recusaram: {resumo['recusaram']}",
                f"Confirmaram leitura: {resumo['confirmaram_leitura']}",
                f"Sem resposta: {resumo['sem_resposta']}",
                f"Quantidade total de itens: {int(cobertura_resumo['total_itens'])}",
                f"Itens com zero orcamentos: {int(cobertura_resumo['itens_zero'])} ({float(cobertura_resumo['perc_zero']):.1f}%)",
                (
                    "Itens com 1 a 2 orcamentos: "
                    f"{int(cobertura_resumo['itens_1_a_2'])} ({float(cobertura_resumo['perc_1_a_2']):.1f}%)"
                ),
                (
                    "Itens com 3 ou mais orcamentos: "
                    f"{int(cobertura_resumo['itens_3_ou_mais'])} ({float(cobertura_resumo['perc_3_ou_mais']):.1f}%)"
                ),
                "",
                "Detalhe por fornecedor:",
            ])
            for p in resumo.get("participacoes") or []:
                linhas.append(f"- {p.get('nome') or p.get('email')} <{p.get('email')}>")
                leitura_ok = _leitura_efetiva(p)
                itens_orcados = itens_orcados_map.get((p.get("email") or "").lower(), 0)
                linhas.append(f"  itens_orcados: {int(itens_orcados)}")
                linhas.append(
                    "  status: "
                    + ", ".join(
                        x for x, ok in [
                            ("enviou_orcamento", p.get("enviou_orcamento")),
                            ("recusou", p.get("recusou")),
                            ("leitura", leitura_ok),
                            ("fez_duvida", p.get("fez_pergunta")),
                        ] if ok
                    )
                )
                linhas.append(f"  pedido: {_fmt_data(p.get('data_pedido_enviado'))}")
                linhas.append(f"  resposta: {_fmt_data(p.get('data_primeira_resposta'))}")
                linhas.append(
                    "  tempo: "
                    + _delta_horas(p.get("data_pedido_enviado"), p.get("data_primeira_resposta"))
                )

            # Adiciona listas de itens por número de orçamentos
            cobertura = resumo.get("item_cobertura") or []
            categorias = _categorizar_itens_por_orcamentos(cobertura)
            linhas.append("")
            linhas.append("Itens com zero orcamento:")
            if categorias.get(0):
                for num, desc in sorted(categorias[0], key=lambda x: (x[0], x[1])):
                    linhas.append(f"- {num}: {desc}")
            else:
                linhas.append("- (nenhum)")

            linhas.append("")
            linhas.append("Itens com 1 orcamento:")
            if categorias.get(1):
                for num, desc in sorted(categorias[1], key=lambda x: (x[0], x[1])):
                    linhas.append(f"- {num}: {desc}")
            else:
                linhas.append("- (nenhum)")

            linhas.append("")
            linhas.append("Itens com 2 orcamentos:")
            if categorias.get(2):
                for num, desc in sorted(categorias[2], key=lambda x: (x[0], x[1])):
                    linhas.append(f"- {num}: {desc}")
            else:
                linhas.append("- (nenhum)")

            linhas.append("")
            linhas.append("Itens com 3 ou mais orcamentos:")
            if categorias.get(3):
                for num, desc in sorted(categorias[3], key=lambda x: (x[0], x[1])):
                    linhas.append(f"- {num}: {desc}")
            else:
                linhas.append("- (nenhum)")

            payload = "\n".join(linhas)
            st.download_button(
                "Baixar relatório txt",
                data=payload.encode("utf-8"),
                file_name=f"relatorio_{processo_sel['numero'].replace('/', '-')}.txt",
                mime="text/plain",
            )
            st.text_area("Prévia", payload, height=320)

        if st.button("Gerar relatório PDF visual", key="btn_relatorio_pdf"):
            try:
                pdf_bytes = _build_report_pdf(processo_sel, resumo)
                st.download_button(
                    "Baixar relatório PDF",
                    data=pdf_bytes,
                    file_name=f"relatorio_{processo_sel['numero'].replace('/', '-')}.pdf",
                    mime="application/pdf",
                    key="download_pdf_relatorio",
                )
                st.success("PDF gerado com sucesso.")
            except Exception as exc:
                st.error(f"Falha ao gerar PDF: {exc}")

        if st.button("Gerar Relatório Gerencial (.docx)", key="btn_relatorio_docx"):
            try:
                categorias_docx = _categorizar_itens_por_orcamentos(
                    resumo.get("item_cobertura") or []
                )
                docx_bytes = report_docx.build_report_docx(
                    processo_sel,
                    resumo,
                    cobertura_resumo,
                    categorias_docx,
                    rotulo_tipo_fn=email_classifier.rotulo_tipo,
                )
                st.download_button(
                    "Baixar Relatório Gerencial (.docx)",
                    data=docx_bytes,
                    file_name=f"Relatorio_Gerencial_{processo_sel['numero'].replace('/', '-')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_docx_relatorio",
                )
                st.success("Relatório Gerencial .docx gerado com sucesso.")
            except Exception as exc:
                st.error(f"Falha ao gerar o Relatório Gerencial: {exc}")


with aba_fornecedores:
    st.subheader("Fornecedores do processo ativo")
    fornecedores = []
    processo_sel = None
    if st.session_state.selected_processo_view_id is not None:
        processo_sel = next(
            (p for p in process_db.listar_processos(conn_proc)
             if p["id"] == st.session_state.selected_processo_view_id),
            None,
        )
    if processo_sel:
        st.caption(f"Processo ativo: {processo_sel['numero']} - {processo_sel.get('titulo') or 'sem titulo'}")
        _atualizar_cnpjs_fornecedores_compat(conn_proc, processo_sel["id"])
        # Enriquecer fornecedores com dados extraídos dos orçamentos vinculados
        conn_budget_local = db_utils.get_connection(st.session_state.budget_db_path)
        try:
            _enriquecer_fornecedores_por_orcamentos(conn_proc, conn_budget_local, processo_sel["id"])
        except Exception:
            pass
        fornecedores = process_db.listar_fornecedores_processo(conn_proc, processo_sel["id"])
    else:
        st.info("Selecione um processo na sidebar para visualizar esta aba.")

    if not fornecedores:
        st.info("Nenhum fornecedor registrado para o processo selecionado.")
    else:
        busca = st.text_input("Buscar por nome/email", "")
        filtrados = [
            f for f in fornecedores
            if not busca
            or busca.lower() in (f.get("nome") or "").lower()
            or busca.lower() in (f.get("email") or "").lower()
            or busca.lower() in (f.get("cnpj") or "").lower()
            or busca.lower() in (f.get("telefone") or "").lower()
        ]
        rows_forn = [
            {
                "Nome": f.get("nome") or "-",
                "Email": f.get("email") or "-",
                "CNPJ": f.get("cnpj") or "-",
                "Telefone": f.get("telefone") or "-",
                "Orçamento": "SIM" if f.get("enviou_orcamento") else "NAO",
                "Recusa": "SIM" if f.get("recusou") else "NAO",
                "Leitura": "SIM" if _leitura_efetiva(f) else "NAO",
                "Dúvida": "SIM" if f.get("fez_pergunta") else "NAO",
                "Pedido enviado": _fmt_data(f.get("data_pedido_enviado")),
                "Primeira resposta": _fmt_data(f.get("data_primeira_resposta")),
            }
            for f in filtrados
        ]
        _safe_dataframe(rows_forn)


with aba_mapa:
    st.subheader("Mapa comparativo do processo")
    lr = st.session_state.last_result
    processo_mapa_id = st.session_state.selected_processo_view_id
    if processo_mapa_id and lr.get("processo_id") != processo_mapa_id:
        conn_budget_map = db_utils.get_connection(budget_db_path)
        rebuilt = _rebuild_map_for_process(
            conn_proc,
            conn_budget_map,
            processo_mapa_id,
            fuzzy_threshold,
            sanity_threshold,
            usar_uf_casamento,
            usar_qtd_casamento,
            bloquear_numero_incoerente,
            gerar_master_auto,
            min_agree,
            consensus_threshold,
        )
        if rebuilt:
            rebuilt["tokens_total"] = lr.get("tokens_total", 0)
            rebuilt["custo_total"] = lr.get("custo_total", 0.0)
            st.session_state.last_result = rebuilt
            lr = rebuilt
    if processo_mapa_id is None:
        st.info("Nenhum processo selecionado para visualização do mapa.")
    elif not lr.get("excel_buffer"):
        st.info("Nenhum mapa comparativo gerado nesta sessão ainda.")
    else:
        st.success(
            f"Mapa pronto: {lr.get('rows_count', 0)} item(ns), "
            f"{lr.get('empresas_count', 0)} empresa(s), "
            f"{lr.get('review_count', 0)} ponto(s) para revisão."
        )
        st.download_button(
            "Baixar Mapa Comparativo (.xlsx)",
            data=lr["excel_buffer"],
            file_name="mapa_comparativo_orcamentos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        cobertura = lr.get("item_cobertura") or []
        if cobertura:
            st.markdown("### Cobertura de Itens")
            st.caption("Meta de conclusão por item: 3 ou mais orçamentos = 100%.")
            _safe_dataframe(cobertura)

        # ------------------------------------------------------------------
        # Revisão de casamentos + memória de correções (aprendizado)
        # ------------------------------------------------------------------
        review_itens = [
            r for r in (lr.get("review") or [])
            if r.get("descricao_nova") and r.get("casou_com")
            and str(r.get("tipo", "")).startswith(("casamento", "casado", "IA juiz", "número igual"))
        ]
        if review_itens:
            st.markdown("### Revisão de casamentos")
            if len(review_itens) > 50:
                st.caption(f"Mostrando os 50 primeiros de {len(review_itens)} casamentos para revisão.")
                review_itens = review_itens[:50]
            st.caption(
                "Confirme ou rejeite os casamentos abaixo. Cada decisão salva vira regra "
                "permanente: nos próximos lotes ela é aplicada antes do fuzzy matching."
            )
            with st.form("form_correcoes_matching"):
                decisoes_ui = []
                for idx, r in enumerate(review_itens):
                    col_info, col_dec = st.columns([4, 1])
                    with col_info:
                        st.markdown(
                            f"**{r.get('tipo')}** (score {r.get('score')})\n\n"
                            f"• {r.get('descricao_nova')}\n\n"
                            f"• {r.get('casou_com')}"
                        )
                    with col_dec:
                        decisao = st.selectbox(
                            "Decisão",
                            ["(manter)", "Confirmar", "Rejeitar"],
                            key=f"decisao_match_{idx}",
                            label_visibility="collapsed",
                        )
                    decisoes_ui.append((r, decisao))
                salvar_decisoes = st.form_submit_button("Salvar decisões no aprendizado")

            if salvar_decisoes:
                try:
                    conn_learn_ui = learning_db.get_connection()
                    n_salvas = 0
                    for r, decisao in decisoes_ui:
                        if decisao == "Confirmar":
                            ok = learning_db.registrar_correcao(
                                conn_learn_ui, r["descricao_nova"], r["casou_com"], "casar"
                            )
                        elif decisao == "Rejeitar":
                            ok = learning_db.registrar_correcao(
                                conn_learn_ui, r["descricao_nova"], r["casou_com"], "nao_casar"
                            )
                        else:
                            continue
                        n_salvas += 1 if ok else 0
                    total_mem = learning_db.contar_correcoes(conn_learn_ui)
                    st.success(
                        f"{n_salvas} decisão(ões) salva(s). A memória de correções tem agora "
                        f"{total_mem} regra(s), aplicadas automaticamente nos próximos processamentos."
                    )
                except Exception as exc:
                    st.error(f"Falha ao salvar decisões: {exc}")


with aba_configuracoes:
    st.subheader("Configurações")

    # ------------------------------------------------------------------
    # 🔐 Administração: chave OpenRouter e modelos (protegido por senha)
    # ------------------------------------------------------------------
    with st.expander("🔐 Administração (chave OpenRouter e modelos)", expanded=False):
        _cfg = app_config.carregar_config()
        _MODELOS_PRINCIPAIS = ["google/gemini-2.5-flash", "deepseek/deepseek-chat-v3.2", "openai/gpt-5-mini"]
        _MODELOS_FORTES = ["google/gemini-2.5-pro", "anthropic/claude-sonnet-4.5", "openai/gpt-5"]

        if not app_config.tem_senha(_cfg):
            st.info(
                "Primeiro acesso: defina a senha de administrador. Depois disso, "
                "só quem tiver a senha altera chave e modelos — os demais usuários "
                "apenas fazem upload dos arquivos."
            )
            with st.form("form_admin_criar_senha"):
                s1 = st.text_input("Nova senha de administrador", type="password")
                s2 = st.text_input("Confirme a senha", type="password")
                criar = st.form_submit_button("Definir senha")
            if criar:
                if not s1 or len(s1) < 6:
                    st.error("Use uma senha com pelo menos 6 caracteres.")
                elif s1 != s2:
                    st.error("As senhas não conferem.")
                else:
                    _cfg = app_config.definir_senha(_cfg, s1)
                    if app_config.salvar_config(_cfg):
                        st.success("Senha definida. Recarregando…")
                        st.rerun()
                    else:
                        st.error("Não foi possível gravar config_app.json.")
        elif not st.session_state.get("admin_unlocked"):
            with st.form("form_admin_login"):
                senha_login = st.text_input("Senha de administrador", type="password")
                entrar = st.form_submit_button("Desbloquear")
            if entrar:
                if app_config.verificar_senha(_cfg, senha_login):
                    st.session_state.admin_unlocked = True
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
        else:
            st.success("Área administrativa desbloqueada.")
            with st.form("form_admin_config"):
                chave_atual = _cfg.get("openrouter_api_key") or ""
                nova_chave = st.text_input(
                    "Chave da API OpenRouter",
                    value=chave_atual,
                    type="password",
                    help="Salva localmente em config_app.json (fora do git). "
                         "Usuários finais nunca veem este campo.",
                )
                _mod = _cfg.get("modelo") or _MODELOS_PRINCIPAIS[0]
                idx_m = _MODELOS_PRINCIPAIS.index(_mod) if _mod in _MODELOS_PRINCIPAIS else 0
                modelo_admin = st.selectbox("Modelo principal (extração/classificação/juiz)", _MODELOS_PRINCIPAIS, index=idx_m)
                _esc = _cfg.get("modelo_escalonamento") or _MODELOS_FORTES[0]
                idx_e = _MODELOS_FORTES.index(_esc) if _esc in _MODELOS_FORTES else 0
                modelo_esc_admin = st.selectbox("Modelo forte (escalonamento automático)", _MODELOS_FORTES, index=idx_e)
                juiz_admin = st.checkbox("IA juiz habilitado por padrão", value=bool(_cfg.get("usar_ia_juiz", True)))
                clf_admin = st.checkbox("Classificação de e-mails com IA por padrão", value=bool(_cfg.get("classificar_emails_com_ia", True)))
                salvar_admin = st.form_submit_button("Salvar configuração")
            if salvar_admin:
                _cfg.update({
                    "openrouter_api_key": (nova_chave or "").strip(),
                    "modelo": modelo_admin,
                    "modelo_escalonamento": modelo_esc_admin,
                    "usar_ia_juiz": bool(juiz_admin),
                    "classificar_emails_com_ia": bool(clf_admin),
                })
                if app_config.salvar_config(_cfg):
                    # aplica imediatamente na sessão atual
                    st.session_state.cfg_model = modelo_admin
                    st.session_state.cfg_usar_ia_juiz = bool(juiz_admin)
                    st.session_state.cfg_classificar_emails_com_ia = bool(clf_admin)
                    st.success("Configuração salva. Usuários finais já não precisam configurar nada.")
                    st.rerun()
                else:
                    st.error("Não foi possível gravar config_app.json.")

            col_lock, col_senha = st.columns(2)
            with col_lock:
                if st.button("Bloquear área administrativa"):
                    st.session_state.admin_unlocked = False
                    st.rerun()
            with col_senha:
                with st.popover("Trocar senha"):
                    with st.form("form_admin_trocar_senha"):
                        ns1 = st.text_input("Nova senha", type="password", key="adm_ns1")
                        ns2 = st.text_input("Confirmar nova senha", type="password", key="adm_ns2")
                        trocar = st.form_submit_button("Trocar")
                    if trocar:
                        if not ns1 or len(ns1) < 6 or ns1 != ns2:
                            st.error("Senha inválida ou não confere (mínimo 6 caracteres).")
                        else:
                            _cfg = app_config.definir_senha(_cfg, ns1)
                            app_config.salvar_config(_cfg)
                            st.success("Senha alterada.")

    st.write("Cache de orçamentos e manutenção do processo selecionado")

    st.session_state.budget_db_path = st.text_input(
        "Banco/cache de orçamentos",
        value=st.session_state.budget_db_path,
        key="cfg_budget_db_path",
    )
    st.session_state.forcar_reprocessamento = st.checkbox(
        "Reprocessar tudo (ignorar cache)",
        value=bool(st.session_state.forcar_reprocessamento),
        key="cfg_forcar_reprocessamento",
    )

    conn_preview = db_utils.get_connection(st.session_state.budget_db_path)
    removidos_antigos = db_utils.purge_old_versions(conn_preview)
    n_cached = db_utils.count_files(conn_preview)
    if removidos_antigos:
        st.caption(f"{removidos_antigos} entrada(s) antigas removidas do cache de orçamentos.")
    st.caption(f"{n_cached} arquivo(s) já salvos no cache de orçamentos.")

    st.divider()
    st.write("Exclusão do processo selecionado")

    processo_ativo = None
    if st.session_state.selected_processo_view_id is not None:
        processo_ativo = next(
            (p for p in process_db.listar_processos(conn_proc)
             if p["id"] == st.session_state.selected_processo_view_id),
            None,
        )

    if processo_ativo is None:
        st.info("Nenhum processo selecionado.")
    else:
        st.warning(
            f"Processo selecionado: {processo_ativo['numero']} - "
            f"{processo_ativo.get('titulo') or 'sem titulo'}"
        )
        st.caption(
            "A exclusão remove apenas este processo e seus dados relacionados, "
            "preservando todos os demais processos."
        )
        if st.button("Solicitar exclusão do processo selecionado"):
            st.session_state.confirmar_deletar_processo = True

        if st.session_state.confirmar_deletar_processo:
            col_yes, col_no = st.columns(2)
            with col_yes:
                if st.button("Sim, excluir processo"):
                    process_db.deletar_processo(conn_proc, processo_ativo["id"])
                    st.session_state.confirmar_deletar_processo = False
                    st.session_state.selected_processo_view_id = None
                    st.session_state.force_modo_novo = True
                    st.success("Processo excluído com sucesso.")
                    st.rerun()
            with col_no:
                if st.button("Não, cancelar exclusão"):
                    st.session_state.confirmar_deletar_processo = False
                    st.info("Exclusão cancelada.")

