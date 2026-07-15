"""
Extração de texto de PDF (com OCR para escaneados), Word e Excel,
e interpretação estruturada via LLM (OpenRouter).
"""
from __future__ import annotations

import io
import json
import os
import re
import shutil
import time
import csv
from collections import Counter

import requests

from confidence import calcular_confianca_estrutural
from normalize_utils import (
    extrair_cnpj,
    extrair_telefone,
    extrair_razao_social,
    limpar_quebras_e_caracteres,
    normalizar_item,
)
from structured_extract import (
    detectar_tipo_e_rotear,
    extrair_docx_estruturado,
    extrair_xlsx_estruturado,
    tentar_extracao_estrutural_pdf,
)
from text_similarity import token_set_ratio

BOILERPLATE_MAX_LEN_SEM_DIGITO = 150
UNIT_TOKENS = (
    "UN", "UND", "UNID", "UNIDADE", "PCT", "PC", "PÇ", "EMB", "KG", "G", "MG",
    "L", "LT", "ML", "CX", "FR", "FD", "RL", "M", "M2", "M3", "PAR", "CJ",
)

# Precos aproximados por 1M tokens (USD) — usados apenas como fallback de estimativa.
# O custo real vem da propria API via payload {"usage": {"include": true}}.
MODEL_PRICING_PER_MILLION = {
    "google/gemini-2.5-flash": {"input": 0.30, "output": 2.50},
    "openai/gpt-5-mini": {"input": 0.25, "output": 2.00},
    "deepseek/deepseek-chat-v3.2": {"input": 0.25, "output": 0.40},
    # legados (mantidos para compatibilidade com cache antigo)
    "openai/gpt-4o-mini": {"input": 0.15, "output": 0.60},
    "anthropic/claude-3-haiku": {"input": 0.25, "output": 1.25},
}

# Modelo padrao de extracao: barato, rapido, com suporte a structured outputs.
DEFAULT_EXTRACTION_MODEL = "google/gemini-2.5-flash"

# Escalonamento: quando o modelo barato falha ou extrai zero itens, uma unica
# retentativa com modelo forte (pago so nos ~5% de documentos dificeis).
ESCALATION_MODEL = "google/gemini-2.5-pro"

_OCR_MIN_ALNUM = 25
_OCR_MIN_TOKENS = 6


def _to_int(value) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _extract_usage_info(response_json: dict, model: str) -> dict:
    usage = response_json.get("usage") or {}
    prompt_tokens = _to_int(usage.get("prompt_tokens") or usage.get("input_tokens"))
    completion_tokens = _to_int(usage.get("completion_tokens") or usage.get("output_tokens"))
    total_tokens = _to_int(usage.get("total_tokens"))
    if total_tokens <= 0:
        total_tokens = prompt_tokens + completion_tokens

    cost_usd = usage.get("cost")
    if cost_usd is None:
        cost_usd = usage.get("total_cost")

    estimated = False
    if cost_usd is None:
        pricing = MODEL_PRICING_PER_MILLION.get(model)
        if pricing:
            cost_usd = (
                (prompt_tokens / 1_000_000.0) * pricing["input"]
                + (completion_tokens / 1_000_000.0) * pricing["output"]
            )
            estimated = True
        else:
            cost_usd = 0.0
            estimated = True

    try:
        cost_usd = float(cost_usd)
    except (TypeError, ValueError):
        cost_usd = 0.0

    return {
        "prompt_tokens": prompt_tokens,
        "completion_tokens": completion_tokens,
        "total_tokens": total_tokens,
        "cost_usd": cost_usd,
        "estimated": estimated,
    }


def _import_pytesseract():
    import pytesseract
    return pytesseract


def _import_image():
    from PIL import Image
    return Image


def _import_document():
    from docx import Document
    return Document


def _import_openpyxl():
    import openpyxl
    return openpyxl


def _import_pdfplumber():
    import pdfplumber
    return pdfplumber


def _ocr_lang_candidates(ocr_lang: str) -> list[str]:
    base = (ocr_lang or "por").strip()
    candidates = [base]
    if "+" not in base:
        candidates.append(f"{base}+eng")
    if "eng" not in candidates:
        candidates.append("eng")
    vistos = set()
    unicos = []
    for cand in candidates:
        if cand and cand not in vistos:
            vistos.add(cand)
            unicos.append(cand)
    return unicos


def ocr_runtime_status() -> tuple[bool, str | None]:
    """Valida se o OCR esta operacional no ambiente atual."""
    if shutil.which("tesseract") is None:
        return False, "Binario 'tesseract' nao encontrado no PATH."

    try:
        pytesseract = _import_pytesseract()
        _ = pytesseract.get_tesseract_version()
    except Exception as exc:
        return False, f"Falha ao inicializar pytesseract/tesseract: {exc}"

    return True, None


def _texto_pobre_para_ocr(texto: str) -> bool:
    """Indica se o texto extraido da pagina esta fraco e deve tentar OCR."""
    base = re.sub(r"\s+", " ", str(texto or "")).strip()
    if not base:
        return True

    alnum = sum(1 for ch in base if ch.isalnum())
    tokens = re.findall(r"\w+", base)
    if alnum < _OCR_MIN_ALNUM:
        return True
    if len(tokens) < _OCR_MIN_TOKENS:
        return True
    return False


def _preferir_texto_ocr(texto_pdf: str, texto_ocr: str) -> bool:
    """Troca para OCR quando ele traz ganho claro de conteudo util."""
    base_pdf = re.sub(r"\s+", " ", str(texto_pdf or "")).strip()
    base_ocr = re.sub(r"\s+", " ", str(texto_ocr or "")).strip()
    if not base_ocr:
        return False
    if not base_pdf:
        return True
    return len(base_ocr) >= int(len(base_pdf) * 1.2)


def _ocr_image(image_obj, ocr_lang: str = "por") -> tuple[str, str | None]:
    ok_ocr, erro_ocr = ocr_runtime_status()
    if not ok_ocr:
        return "", erro_ocr

    pytesseract = _import_pytesseract()
    last_error = None
    for lang in _ocr_lang_candidates(ocr_lang):
        try:
            texto = pytesseract.image_to_string(image_obj, lang=lang)
            if texto and texto.strip():
                return texto.strip(), None
        except Exception as exc:
            last_error = str(exc)
    return "", last_error


def extract_text_from_image(path: str, ocr_lang: str = 'por') -> tuple[str, str | None]:
    Image = _import_image()
    with Image.open(path) as img:
        return _ocr_image(img, ocr_lang=ocr_lang)


# ---------- Extração de texto bruto ----------

def extract_text_from_pdf(path: str, ocr_lang: str = 'por') -> str:
    pdfplumber = _import_pdfplumber()
    text_parts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_text = (page.extract_text() or "").strip()
            if _texto_pobre_para_ocr(page_text):
                try:
                    page_img = page.to_image(resolution=250).original
                    ocr_text, _ = _ocr_image(page_img, ocr_lang=ocr_lang)
                    if _preferir_texto_ocr(page_text, ocr_text):
                        page_text = ocr_text
                except Exception:
                    pass
            text_parts.append(page_text)
    return "\n".join(text_parts)


def extract_text_from_docx(path: str) -> str:
    Document = _import_document()
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(parts)


def extract_text_from_xlsx(path: str) -> str:
    openpyxl = _import_openpyxl()
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    for sheet in wb.worksheets:
        parts.append(f"[Aba: {sheet.title}]")
        for row in sheet.iter_rows(values_only=True):
            row_vals = [str(c) for c in row if c is not None]
            if row_vals:
                parts.append(" | ".join(row_vals))
    return "\n".join(parts)


def extract_text_from_csv(path: str) -> str:
    parts = []
    with open(path, "r", encoding="utf-8", errors="ignore", newline="") as fh:
        sample = fh.read(4096)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            reader = csv.reader(fh, dialect)
        except Exception:
            reader = csv.reader(fh)
        for row in reader:
            vals = [str(c).strip() for c in row if str(c).strip()]
            if vals:
                parts.append(" | ".join(vals))
    return "\n".join(parts)


def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == '.pdf':
        return extract_text_from_pdf(path)
    elif ext in ('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp'):
        txt, _ = extract_text_from_image(path)
        return txt
    elif ext in ('.docx', '.doc'):
        return extract_text_from_docx(path)
    elif ext in ('.xlsx', '.xls'):
        return extract_text_from_xlsx(path)
    elif ext == '.csv':
        return extract_text_from_csv(path)
    raise ValueError(f"Formato não suportado: {ext}")


def pre_filtrar_texto(text: str) -> str:
    """Remove ruído (parágrafos legais longos sem números, cabeçalhos/rodapés repetidos)
    antes de enviar o texto para a IA — reduz tokens sem descartar itens de verdade.
    Nunca remove uma linha só por não ter preço, para não perder itens não cotados."""
    linhas = [l.strip() for l in text.split("\n")]
    contagem = Counter(l for l in linhas if len(l) > 10)

    linhas_filtradas = []
    repetidas_vistas = Counter()

    for linha in linhas:
        if not linha:
            continue

        tem_digito = any(c.isdigit() for c in linha)
        if len(linha) > BOILERPLATE_MAX_LEN_SEM_DIGITO and not tem_digito:
            continue  # provável parágrafo jurídico/termos, sem nenhum número

        if len(linha) > 10 and contagem[linha] >= 3:
            repetidas_vistas[linha] += 1
            if repetidas_vistas[linha] > 1:
                continue  # cabeçalho/rodapé repetido em várias páginas: mantém só a 1ª ocorrência

        linhas_filtradas.append(linha)

    return "\n".join(linhas_filtradas)


def _parse_numero_br(valor: str):
    if valor is None:
        return None
    texto = str(valor).strip()
    if not texto:
        return None
    texto = re.sub(r"[^\d,.-]", "", texto)
    if not texto:
        return None
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def _clean_unidade(valor: str):
    if valor is None:
        return None
    texto = re.sub(r"\s+", " ", str(valor).strip().upper())
    return texto or None


def _extract_unit_qty_from_segment(segmento: str):
    if not segmento:
        return None, None

    partes = [p.strip() for p in re.split(r"\s*\|\s*", segmento) if p.strip()]
    if len(partes) >= 2:
        for idx, parte in enumerate(partes[:-1]):
            unidade = _clean_unidade(parte)
            if unidade in UNIT_TOKENS:
                quantidade = _parse_numero_br(partes[idx + 1])
                if quantidade is not None:
                    return unidade, quantidade

        for idx, parte in enumerate(partes[1:], start=1):
            unidade = _clean_unidade(parte)
            if unidade in UNIT_TOKENS:
                quantidade = _parse_numero_br(partes[idx - 1])
                if quantidade is not None:
                    return unidade, quantidade

    padroes = [
        (rf"\b({'|'.join(UNIT_TOKENS)})\b\s+(\d+(?:[\.,]\d+)?)\b", "unit_first"),
        (rf"\b(\d+(?:[\.,]\d+)?)\b\s+({'|'.join(UNIT_TOKENS)})\b", "qty_first"),
    ]
    for padrao, ordem in padroes:
        match = re.search(padrao, segmento, flags=re.IGNORECASE)
        if not match:
            continue
        if ordem == "unit_first":
            unidade, quantidade = match.group(1), match.group(2)
        else:
            quantidade, unidade = match.group(1), match.group(2)
        quantidade = _parse_numero_br(quantidade)
        unidade = _clean_unidade(unidade)
        if quantidade is not None and unidade:
            return unidade, quantidade

    return None, None


def _candidate_item_lines(text: str, numero_item, descricao: str):
    linhas = [l.strip() for l in text.splitlines() if l.strip()]
    candidatos = []
    numero = str(numero_item).strip() if numero_item is not None else ""
    descricao_norm = re.sub(r"\s+", " ", (descricao or "").strip().lower())
    palavras_desc = [p for p in re.findall(r"\w+", descricao_norm) if len(p) >= 4][:6]

    for idx, linha in enumerate(linhas):
        linha_norm = linha.lower()
        score = 0
        if numero and re.search(rf"\b{re.escape(numero)}\b", linha):
            score += 5
        score += sum(1 for palavra in palavras_desc if palavra in linha_norm)
        if score > 0:
            candidatos.append((score, idx, linha))

    candidatos.sort(reverse=True)
    resultado = []
    vistos = set()
    for _, idx, linha in candidatos[:5]:
        for near_idx in (idx, idx + 1, idx - 1):
            if 0 <= near_idx < len(linhas) and near_idx not in vistos:
                vistos.add(near_idx)
                resultado.append(linhas[near_idx])
    return resultado


def _enrich_item_fields(result: dict, text: str) -> dict:
    itens = result.get("itens")
    if not isinstance(itens, list):
        return result

    for item in itens:
        if not isinstance(item, dict):
            continue
        precisa_unidade = not item.get("unidade")
        precisa_quantidade = item.get("quantidade") is None
        if not (precisa_unidade or precisa_quantidade):
            continue

        for linha in _candidate_item_lines(text, item.get("numero_item"), item.get("descricao", "")):
            unidade, quantidade = _extract_unit_qty_from_segment(linha)
            if precisa_unidade and unidade:
                item["unidade"] = unidade
                precisa_unidade = False
            if precisa_quantidade and quantidade is not None:
                item["quantidade"] = quantidade
                precisa_quantidade = False
            if not (precisa_unidade or precisa_quantidade):
                break

    return result


# ---------- Interpretação estruturada via LLM (OpenRouter) ----------

EXTRACTION_SYSTEM_PROMPT = """Você é um assistente especializado em extrair itens de orçamentos/cotações de compras.
Dado o texto extraído de um documento de orçamento, retorne APENAS um JSON válido (sem markdown, sem texto adicional) no formato:

{
  "empresa": "nome da empresa/fornecedor, se identificável, senão null",
  "itens": [
    {
      "numero_item": "número do item conforme edital/TR, se houver, senão null",
      "codigo": "código único do item (PI, NSN, Part Number, Nº de Estoque, Código do Item), se houver, senão null",
      "descricao": "descrição do item/produto",
      "unidade": "unidade de medida, se houver, senão null",
      "quantidade": numero ou null,
      "preco_unitario": numero (float, use ponto decimal) ou null,
      "preco_total": numero ou null
    }
  ]
}

Regras:
- Extraia TODOS os itens de orçamento encontrados no texto, inclusive os que aparecem sem preço definido.
- "codigo" é o identificador único do MATERIAL (colunas tipo PI, NSN, P/N, Part Number, Nº Estoque, Cód. Item, Referência) — não confundir com "numero_item", que é a posição sequencial do item no edital (1, 2, 3...). Copie o código exatamente como está, com traços e pontos.
- Copie a "descricao" exatamente como está escrita no documento, por extenso — não abrevie, não resuma e não corrija siglas técnicas. Se o mesmo item aparecer mais de uma vez no documento com descrições diferentes (uma mais completa que a outra), use a versão mais completa.
- Quando houver coluna de unidade de fornecimento (UF, UND, UN, PCT, EMB, KG, LT etc.), preencha obrigatoriamente o campo "unidade".
- Quando houver coluna de quantidade (QTD, QUANT, QUANTIDADE), preencha obrigatoriamente o campo "quantidade".
- Em tabelas, preserve a associação da mesma linha do item entre descricao, unidade, quantidade e preco.
- Se preco_unitario nao estiver explicito mas preco_total e quantidade estiverem, calcule preco_unitario = preco_total / quantidade.
- Numeros devem ser float puro, sem simbolo de moeda ou separador de milhar (ex: 1234.56, nunca "R$ 1.234,56").
- Nao invente dados que nao estao no texto.
- Responda em português.
"""


# JSON Schema para structured outputs (OpenRouter response_format).
# Garante JSON valido e tipado direto do modelo, sem parsing fragil de markdown.
EXTRACTION_JSON_SCHEMA = {
    "name": "extracao_orcamento",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "empresa": {"type": ["string", "null"]},
            "itens": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "numero_item": {"type": ["string", "null"]},
                        "codigo": {"type": ["string", "null"]},
                        "descricao": {"type": "string"},
                        "unidade": {"type": ["string", "null"]},
                        "quantidade": {"type": ["number", "null"]},
                        "preco_unitario": {"type": ["number", "null"]},
                        "preco_total": {"type": ["number", "null"]},
                    },
                    "required": ["numero_item", "codigo", "descricao", "unidade",
                                 "quantidade", "preco_unitario", "preco_total"],
                    "additionalProperties": False,
                },
            },
        },
        "required": ["empresa", "itens"],
        "additionalProperties": False,
    },
}


def _call_openrouter_extract_once(text: str, api_key: str, model: str = None,
                                   filename_hint: str = None, max_retries: int = 3,
                                   max_chars: int = 50000, pre_filtrar: bool = True) -> dict:
    model = model or DEFAULT_EXTRACTION_MODEL
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    texto_base = pre_filtrar_texto(text) if pre_filtrar else text
    texto_truncado = len(texto_base) > max_chars
    user_content = texto_base[:max_chars]
    if filename_hint:
        user_content = f"Nome do arquivo: {filename_hint}\n\n{user_content}"

    def _montar_payload(com_schema: bool) -> dict:
        p = {
            "model": model,
            "messages": [
                {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
            "temperature": 0,
            # OpenRouter devolve o custo real da chamada em usage.cost
            "usage": {"include": True},
        }
        if com_schema:
            p["response_format"] = {"type": "json_schema", "json_schema": EXTRACTION_JSON_SCHEMA}
        return p

    usar_schema = True
    last_error = "erro desconhecido"
    for attempt in range(max_retries):
        try:
            resp = requests.post(url, headers=headers, json=_montar_payload(usar_schema), timeout=120)
            if resp.status_code in (400, 404) and usar_schema:
                # Provedor/modelo sem suporte a structured outputs: refaz sem schema
                usar_schema = False
                resp = requests.post(url, headers=headers, json=_montar_payload(False), timeout=120)
            resp.raise_for_status()
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            content_clean = re.sub(r"^```json\s*|\s*```$", "", content.strip(), flags=re.MULTILINE)
            content_clean = content_clean.strip().strip("`")
            parsed = json.loads(content_clean)
            parsed = _enrich_item_fields(parsed, texto_base)
            if isinstance(parsed.get("itens"), list):
                parsed["itens"] = [normalizar_item(i) for i in parsed["itens"] if isinstance(i, dict)]
            parsed["texto_truncado"] = texto_truncado
            parsed["usage"] = _extract_usage_info(data, model)
            return parsed
        except (requests.RequestException, json.JSONDecodeError, KeyError, IndexError) as exc:
            last_error = str(exc)
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)

    return {
        "empresa": filename_hint,
        "itens": [],
        "erro": last_error,
        "texto_truncado": texto_truncado,
        "usage": {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
            "cost_usd": 0.0,
            "estimated": True,
        },
    }


def call_openrouter_extract(text: str, api_key: str, model: str = None,
                             filename_hint: str = None, max_retries: int = 3,
                             max_chars: int = 50000, pre_filtrar: bool = True,
                             escalate: bool = True) -> dict:
    """
    Extração via LLM com escalonamento automático: tenta o modelo barato;
    se a extração falhar ou vier vazia, reextrai UMA vez com o modelo forte
    (ESCALATION_MODEL). O custo das duas chamadas é somado no usage.
    """
    model = model or DEFAULT_EXTRACTION_MODEL
    parsed = _call_openrouter_extract_once(
        text, api_key, model=model, filename_hint=filename_hint,
        max_retries=max_retries, max_chars=max_chars, pre_filtrar=pre_filtrar,
    )

    extracao_ruim = bool(parsed.get("erro")) or not parsed.get("itens")
    if escalate and extracao_ruim and model != ESCALATION_MODEL:
        usage_barato = parsed.get("usage") or {}
        parsed_forte = _call_openrouter_extract_once(
            text, api_key, model=ESCALATION_MODEL, filename_hint=filename_hint,
            max_retries=max_retries, max_chars=max_chars, pre_filtrar=pre_filtrar,
        )
        if parsed_forte.get("itens") or not parsed.get("itens"):
            # usa o resultado do modelo forte; soma o gasto da tentativa barata
            usage_forte = parsed_forte.get("usage") or {}
            for campo in ("prompt_tokens", "completion_tokens", "total_tokens"):
                usage_forte[campo] = _to_int(usage_forte.get(campo)) + _to_int(usage_barato.get(campo))
            try:
                usage_forte["cost_usd"] = float(usage_forte.get("cost_usd") or 0) + float(usage_barato.get("cost_usd") or 0)
            except (TypeError, ValueError):
                pass
            parsed_forte["usage"] = usage_forte
            parsed_forte["escalado_para"] = ESCALATION_MODEL
            return parsed_forte

    return parsed


def _render_pdf_paginas_png(path: str, max_paginas: int = 8, resolution: int = 200) -> list[bytes]:
    """Renderiza as primeiras paginas do PDF como PNG (para extracao por visao)."""
    import io as _io
    pdfplumber = _import_pdfplumber()
    imagens = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages[:max_paginas]:
            try:
                img = page.to_image(resolution=resolution).original
                buf = _io.BytesIO()
                img.save(buf, format="PNG")
                imagens.append(buf.getvalue())
            except Exception:
                continue
    return imagens


def call_openrouter_vision_extract(imagens: list[bytes], api_key: str, model: str = None,
                                    filename_hint: str = None, max_retries: int = 2,
                                    mime: str = "image/png") -> dict:
    """
    Extracao de itens diretamente da IMAGEM do documento, via modelo multimodal.

    Ultimo recurso da cascata de OCR: usado quando o Tesseract nao esta
    disponivel ou devolveu texto inutilizavel. O modelo de visao le layout,
    tabelas, carimbos e manuscrito melhor que OCR classico em documento ruim,
    ao custo de alguns centavos por documento.
    """
    import base64

    model = model or DEFAULT_EXTRACTION_MODEL
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    conteudo = [{
        "type": "text",
        "text": (
            (f"Nome do arquivo: {filename_hint}\n" if filename_hint else "")
            + "Extraia os itens de orcamento das imagens deste documento, "
              "seguindo exatamente as regras do sistema."
        ),
    }]
    for img_bytes in imagens[:8]:
        b64 = base64.b64encode(img_bytes).decode("ascii")
        conteudo.append({
            "type": "image_url",
            "image_url": {"url": f"data:{mime};base64,{b64}"},
        })

    def _payload(com_schema: bool) -> dict:
        p = {
            "model": model,
            "messages": [
                {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                {"role": "user", "content": conteudo},
            ],
            "temperature": 0,
            "usage": {"include": True},
        }
        if com_schema:
            p["response_format"] = {"type": "json_schema", "json_schema": EXTRACTION_JSON_SCHEMA}
        return p

    usar_schema = True
    last_error = "erro desconhecido"
    for attempt in range(max_retries + 1):
        try:
            resp = requests.post(url, headers=headers, json=_payload(usar_schema), timeout=180)
            if resp.status_code in (400, 404) and usar_schema:
                usar_schema = False
                resp = requests.post(url, headers=headers, json=_payload(False), timeout=180)
            resp.raise_for_status()
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            content_clean = re.sub(r"^```json\s*|\s*```$", "", content.strip(), flags=re.MULTILINE)
            parsed = json.loads(content_clean.strip().strip("`"))
            if isinstance(parsed.get("itens"), list):
                parsed["itens"] = [normalizar_item(i) for i in parsed["itens"] if isinstance(i, dict)]
            parsed["texto_truncado"] = False
            parsed["usage"] = _extract_usage_info(data, model)
            parsed["fonte_extracao"] = "visao"
            return parsed
        except (requests.RequestException, json.JSONDecodeError, KeyError, IndexError) as exc:
            last_error = str(exc)
            if attempt < max_retries:
                time.sleep(2 ** attempt)

    return {
        "empresa": filename_hint,
        "itens": [],
        "erro": f"Extracao por visao falhou: {last_error}",
        "texto_truncado": False,
        "usage": {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0,
                   "cost_usd": 0.0, "estimated": True},
    }


def _extract_text_pages_pdf(path: str, ocr_lang: str = "por"):
    pdfplumber = _import_pdfplumber()
    pages = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_text = (page.extract_text() or "").strip()
            veio_ocr = False
            erro_ocr = None
            tentou_ocr = _texto_pobre_para_ocr(page_text)
            if tentou_ocr:
                try:
                    page_img = page.to_image(resolution=250).original
                    ocr_text, erro_ocr = _ocr_image(page_img, ocr_lang=ocr_lang)
                    if _preferir_texto_ocr(page_text, ocr_text):
                        page_text = ocr_text
                        veio_ocr = True
                except Exception as exc:
                    erro_ocr = str(exc)
            pages.append({
                "texto": page_text or "",
                "veio_ocr": veio_ocr,
                "tentou_ocr": tentou_ocr,
                "erro_ocr": erro_ocr,
            })
    return pages


def _inferir_origem_pagina(item: dict, pages: list[dict]) -> str:
    descricao = limpar_quebras_e_caracteres(str(item.get("descricao") or "")).lower()
    numero = str(item.get("numero_item") or "").strip()
    if not pages:
        return "documento"

    for idx, p in enumerate(pages, start=1):
        texto = p.get("texto", "").lower()
        if numero and re.search(rf"\b{re.escape(numero)}\b", texto):
            return f"Pagina {idx}"
        if descricao and descricao[:40] and descricao[:40] in texto:
            return f"Pagina {idx}"

    return "Pagina 1"


def _anotar_fonte_origem(itens: list[dict], fonte: str, pages: list[dict] | None = None) -> list[dict]:
    saida = []
    for item in itens or []:
        i2 = dict(item)
        i2["fonte_extracao"] = fonte
        if not i2.get("origem"):
            i2["origem"] = _inferir_origem_pagina(i2, pages or [])
        saida.append(normalizar_item(i2))
    return saida


def calcular_similaridade_item(item_parser: dict, item_ia: dict) -> float:
    num_parser = str(item_parser.get("numero_item") or "").strip()
    num_ia = str(item_ia.get("numero_item") or "").strip()
    if num_parser and num_ia and num_parser == num_ia:
        return 100.0

    desc_parser = limpar_quebras_e_caracteres(str(item_parser.get("descricao") or "")).lower()
    desc_ia = limpar_quebras_e_caracteres(str(item_ia.get("descricao") or "")).lower()
    sim_desc = token_set_ratio(desc_parser, desc_ia)

    unidade_parser = str(item_parser.get("unidade") or "").upper()
    unidade_ia = str(item_ia.get("unidade") or "").upper()
    sim_unidade = 100.0 if unidade_parser and unidade_parser == unidade_ia else 0.0

    qtd_p = item_parser.get("quantidade")
    qtd_i = item_ia.get("quantidade")
    sim_qtd = 0.0
    if qtd_p is not None and qtd_i is not None:
        if qtd_p == qtd_i:
            sim_qtd = 100.0
        elif max(abs(qtd_p), abs(qtd_i)) > 0:
            diff = abs(qtd_p - qtd_i) / max(abs(qtd_p), abs(qtd_i))
            if diff < 0.05:
                sim_qtd = 100.0

    return (0.60 * sim_desc) + (0.25 * sim_unidade) + (0.15 * sim_qtd)


def comparar_extracoes(itens_parser: list[dict], itens_ia: list[dict]):
    def _mesclar_item(parser_item: dict, ia_item: dict, score: float, numerico_bate: bool) -> dict:
        item_final = dict(parser_item)
        item_final["fonte_extracao"] = "dupla_checagem"
        item_final["confianca"] = "alta" if score >= 70 and numerico_bate else "baixa"

        if ia_item.get("descricao") and len(str(ia_item.get("descricao"))) > len(str(item_final.get("descricao") or "")):
            item_final["descricao"] = ia_item.get("descricao")

        for campo in ("unidade", "quantidade", "preco_total", "origem"):
            if item_final.get(campo) in (None, "") and ia_item.get(campo) not in (None, ""):
                item_final[campo] = ia_item.get(campo)

        if item_final.get("preco_unitario") is None and ia_item.get("preco_unitario") is not None:
            item_final["preco_unitario"] = ia_item.get("preco_unitario")

        return normalizar_item(item_final)

    pares = []
    for i, item_p in enumerate(itens_parser):
        for j, item_i in enumerate(itens_ia):
            score = calcular_similaridade_item(item_p, item_i)
            pares.append((score, i, j))

    pares.sort(reverse=True, key=lambda x: x[0])
    usados_parser = set()
    usados_ia = set()
    casamentos = []
    for score, i, j in pares:
        if i in usados_parser or j in usados_ia:
            continue
        usados_parser.add(i)
        usados_ia.add(j)
        casamentos.append((score, itens_parser[i], itens_ia[j]))

    itens_finais = []
    conflitos = []
    for score, parser_item, ia_item in casamentos:
        p_unit = parser_item.get("preco_unitario")
        i_unit = ia_item.get("preco_unitario")
        numerico_bate = False
        if p_unit is not None and i_unit is not None and max(abs(p_unit), abs(i_unit), 1e-9) > 0:
            numerico_bate = abs(p_unit - i_unit) / max(abs(p_unit), abs(i_unit)) < 0.01

        itens_finais.append(_mesclar_item(parser_item, ia_item, score, numerico_bate))

        if score < 70 or not numerico_bate:
            conflitos.append({
                "tipo": "divergência parser vs. IA",
                "numero_item": parser_item.get("numero_item") or ia_item.get("numero_item"),
                "descricao_nova": parser_item.get("descricao") or "",
                "casou_com": ia_item.get("descricao") or "",
                "score": round(score, 2),
            })

    for idx, parser_item in enumerate(itens_parser):
        if idx not in usados_parser:
            item_final = dict(parser_item)
            item_final["fonte_extracao"] = "dupla_checagem"
            item_final["confianca"] = "baixa"
            itens_finais.append(normalizar_item(item_final))
            conflitos.append({
                "tipo": "divergência parser vs. IA",
                "numero_item": parser_item.get("numero_item"),
                "descricao_nova": parser_item.get("descricao") or "",
                "casou_com": "sem par na IA",
                "score": 0,
            })

    for idx, ia_item in enumerate(itens_ia):
        if idx not in usados_ia:
            item_final = dict(ia_item)
            item_final["fonte_extracao"] = "dupla_checagem"
            item_final["confianca"] = "média"
            itens_finais.append(normalizar_item(item_final))
            conflitos.append({
                "tipo": "divergência parser vs. IA",
                "numero_item": ia_item.get("numero_item"),
                "descricao_nova": ia_item.get("descricao") or "",
                "casou_com": "sem par no parser",
                "score": 0,
            })

    return itens_finais, conflitos


def extrair_orcamento_em_camadas(path: str, api_key: str, model: str,
                                 pre_filtrar: bool = True,
                                 limiar_alto: int = 85,
                                 limiar_baixo: int = 40) -> dict:
    """Pipeline em camadas: estrutural quando possivel, IA por excecao."""
    tipo = detectar_tipo_e_rotear(path)
    nome_arquivo = os.path.basename(path)
    review = []
    confianca = None
    texto_base = ""
    texto_truncado = False
    fonte_global = "ia"
    debug_events = []
    debug_highlights = []

    debug_events.append(f"Arquivo recebido: {nome_arquivo}")
    debug_events.append(f"Percepcao inicial: tipo detectado = {tipo}")

    if tipo == "xlsx":
        debug_events.append("Biblioteca acionada: openpyxl (extracao estrutural de planilha)")
    elif tipo == "csv":
        debug_events.append("Biblioteca acionada: csv (leitura textual estruturada)")
    elif tipo == "docx":
        debug_events.append("Biblioteca acionada: python-docx (extracao estrutural de documento Word)")
    elif tipo in ("pdf_texto", "pdf_escaneado"):
        debug_events.append("Biblioteca acionada: pdfplumber (leitura de paginas PDF)")
        debug_events.append("Biblioteca potencial: requests (chamada IA quando necessario)")
    elif tipo == "imagem":
        debug_events.append("Biblioteca acionada: Pillow + pytesseract (OCR de imagem)")
        debug_events.append("Biblioteca potencial: requests (chamada IA quando necessario)")

    if tipo == "xlsx":
        itens = extrair_xlsx_estruturado(path)
        if itens:
            fonte_global = "estrutural"
            texto_base = extract_text(path)
            empresa_det = extrair_razao_social(texto_base)
            debug_events.append(f"Achado: {len(itens)} item(ns) extraido(s) por modo estrutural")
            debug_highlights.append("RELEVANTE: processamento 100% estrutural (sem consumo de IA)")
            return {
                "empresa": empresa_det or os.path.splitext(nome_arquivo)[0],
                "itens": itens,
                "review": review,
                "confianca_estrutural": 100,
                "fonte_processamento": fonte_global,
                "texto_truncado": False,
                "debug_events": debug_events,
                "debug_highlights": debug_highlights,
                "cnpj": extrair_cnpj(texto_base),
                "telefone": extrair_telefone(texto_base),
                "usage": {
                    "prompt_tokens": 0,
                    "completion_tokens": 0,
                    "total_tokens": 0,
                    "cost_usd": 0.0,
                    "estimated": False,
                },
            }

    if tipo == "docx":
        itens = extrair_docx_estruturado(path)
        if itens:
            fonte_global = "estrutural"
            texto_base = extract_text(path)
            empresa_det = extrair_razao_social(texto_base)
            debug_events.append(f"Achado: {len(itens)} item(ns) extraido(s) por modo estrutural")
            debug_highlights.append("RELEVANTE: processamento 100% estrutural (sem consumo de IA)")
            return {
                "empresa": empresa_det or os.path.splitext(nome_arquivo)[0],
                "itens": itens,
                "review": review,
                "confianca_estrutural": 100,
                "fonte_processamento": fonte_global,
                "texto_truncado": False,
                "debug_events": debug_events,
                "debug_highlights": debug_highlights,
                "cnpj": extrair_cnpj(texto_base),
                "telefone": extrair_telefone(texto_base),
                "usage": {
                    "prompt_tokens": 0,
                    "completion_tokens": 0,
                    "total_tokens": 0,
                    "cost_usd": 0.0,
                    "estimated": False,
                },
            }

        if tipo == "eml":
            # processa .eml: se houver anexos, delega ao anexo; se nao, usa corpo do email
            try:
                import email
                with open(path, 'rb') as fh:
                    raw = fh.read()
                msg = email.message_from_bytes(raw)
            except Exception:
                # nao conseguiu abrir como eml: cair no fluxo normal
                msg = None
            if msg is not None:
                # coleta anexos
                attach_paths = []
                for idx, part in enumerate(msg.walk()):
                    fn = part.get_filename()
                    if fn:
                        payload = part.get_payload(decode=True)
                        if payload:
                            tmp = f"/tmp/{nome_arquivo}.attach.{idx}.{fn}"
                            try:
                                with open(tmp, 'wb') as out:
                                    out.write(payload)
                                attach_paths.append(tmp)
                            except Exception:
                                continue
                if attach_paths:
                    # procura remetente para possivel override do nome da empresa
                    try:
                        from_hdr = msg.get('From') or ''
                        m = re.match(r"(?P<name>.+?)\s*<.+?>", from_hdr)
                        if m:
                            empresa_from = m.group('name').strip()
                        else:
                            empresa_from = (from_hdr.split('@')[0] if '@' in from_hdr else from_hdr).strip()
                    except Exception:
                        empresa_from = None

                    # processa o primeiro anexo detectado
                    res_attach = extrair_orcamento_em_camadas(attach_paths[0], api_key, model, pre_filtrar, limiar_alto, limiar_baixo)
                    if empresa_from and (not res_attach.get('empresa') or any(tok in str(res_attach.get('empresa')).lower() for tok in ('image', 'attachment', 'unnamed', 'proposta'))):
                        res_attach['empresa'] = empresa_from
                    return res_attach

                # sem anexos: monta texto a partir das partes textuais
                parts = []
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain' and not part.get_filename():
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                parts.append(payload.decode('utf-8', errors='ignore'))
                        except Exception:
                            pass
                texto_base = "\n".join(parts) if parts else ''
                pages = [{"texto": texto_base, "veio_ocr": False, "tentou_ocr": False, "erro_ocr": None}]

                llm = call_openrouter_extract(
                    texto_base,
                    api_key,
                    model=model,
                    filename_hint=nome_arquivo,
                    pre_filtrar=pre_filtrar,
                )
                itens_ia = _anotar_fonte_origem(llm.get("itens", []), "ia", pages)
                itens_finais = itens_ia

                empresa_det = extrair_razao_social(texto_base)
                empresa_llm = llm.get("empresa")

                # Heuristica: se empresa detectada aparenta ser o comprador (ex: 'marinha', 'solicita'),
                # prefira o remetente (From) do email como empresa
                buyer_indicators = ('marinha', 'solicita', 'solicitação', 'solicitacao', 'pedido', 'res_')
                nome_empresa_detectada = (empresa_llm or empresa_det or '').lower()
                prefer_from = any(tok in nome_empresa_detectada for tok in buyer_indicators) or not nome_empresa_detectada
                if prefer_from:
                    try:
                        from_hdr = msg.get('From') or ''
                        m = re.match(r"(?P<name>.+?)\s*<.+?>", from_hdr)
                        if m:
                            empresa_from = m.group('name').strip()
                        else:
                            empresa_from = (from_hdr.split('@')[0] if '@' in from_hdr else from_hdr).strip()
                        if empresa_from:
                            empresa_llm = empresa_from
                            empresa_det = empresa_from
                    except Exception:
                        pass

                # Corrige preco_unitario quando necessario
                itens_corrigidos = []
                for i in itens_finais:
                    p_unit = i.get('preco_unitario')
                    p_total = i.get('preco_total')
                    qtd = i.get('quantidade')
                    try:
                        if (p_unit is None or (p_total is not None and qtd)) and p_total is not None and qtd:
                            i['preco_unitario'] = float(p_total) / float(qtd)
                        elif p_unit is not None and p_total is not None and qtd:
                            if float(p_unit) > float(p_total):
                                i['preco_unitario'] = float(p_total) / float(qtd)
                    except Exception:
                        pass
                    itens_corrigidos.append(i)

                return {
                    'empresa': empresa_llm or empresa_det or os.path.splitext(nome_arquivo)[0],
                    'itens': [normalizar_item(i) for i in itens_corrigidos],
                    'review': review,
                    'confianca_estrutural': None,
                    'fonte_processamento': 'ia',
                    'texto_truncado': bool(llm.get('texto_truncado')),
                    'cnpj': extrair_cnpj(texto_base),
                    'telefone': extrair_telefone(texto_base),
                    'debug_events': debug_events,
                    'debug_highlights': debug_highlights,
                    'usage': llm.get('usage', {}),
                }
    if tipo in ("pdf_texto", "pdf_escaneado"):
        pages = _extract_text_pages_pdf(path)
        texto_base = "\n".join(p["texto"] for p in pages)
        debug_events.append(f"Achado: {len(pages)} pagina(s) lida(s) no PDF")
        debug_events.append(f"Achado: {len(texto_base)} caractere(s) extraido(s) do documento")
        paginas_ocr = sum(1 for p in pages if p.get("veio_ocr"))
        paginas_tentou_ocr = sum(1 for p in pages if p.get("tentou_ocr"))
        if paginas_tentou_ocr:
            debug_events.append(f"Achado: OCR tentado em {paginas_tentou_ocr} pagina(s)")
        if paginas_ocr:
            debug_events.append(f"Achado: OCR aplicado em {paginas_ocr} pagina(s)")

        if not texto_base.strip():
            # --- Fallback de visao: OCR falhou/indisponivel -> manda a IMAGEM
            # das paginas para um modelo multimodal (ultimo recurso da cascata)
            if api_key:
                try:
                    imagens = _render_pdf_paginas_png(path)
                except Exception:
                    imagens = []
                if imagens:
                    debug_events.append(
                        f"Fallback de visao: {len(imagens)} pagina(s) enviadas como imagem ao modelo multimodal"
                    )
                    llm_visao = call_openrouter_vision_extract(
                        imagens, api_key, model=model, filename_hint=nome_arquivo
                    )
                    if llm_visao.get("itens"):
                        debug_highlights.append(
                            "RELEVANTE: itens extraidos por VISAO (OCR classico falhou neste documento)"
                        )
                        return {
                            "empresa": llm_visao.get("empresa") or os.path.splitext(nome_arquivo)[0],
                            "itens": [normalizar_item(i) for i in llm_visao["itens"]],
                            "review": review,
                            "confianca_estrutural": None,
                            "fonte_processamento": "visao",
                            "texto_truncado": False,
                            "cnpj": None,
                            "telefone": None,
                            "debug_events": debug_events,
                            "debug_highlights": debug_highlights,
                            "usage": llm_visao.get("usage", {}),
                        }

            erros_ocr = [p.get("erro_ocr") for p in pages if p.get("erro_ocr")]
            mensagem_erro = "Texto vazio apos leitura do PDF; nao foi possivel extrair conteudo util."
            if erros_ocr:
                mensagem_erro = (
                    "Texto vazio apos leitura do PDF e OCR indisponivel/falhou. "
                    f"Detalhe: {erros_ocr[0]}"
                )
                debug_highlights.append("RELEVANTE: OCR nao disponivel no ambiente")
            return {
                "empresa": os.path.splitext(nome_arquivo)[0],
                "itens": [],
                "erro": mensagem_erro,
                "review": review,
                "confianca_estrutural": 0,
                "fonte_processamento": "ia",
                "texto_truncado": False,
                "cnpj": None,
                "telefone": None,
                "debug_events": debug_events,
                "debug_highlights": debug_highlights,
                "usage": {
                    "prompt_tokens": 0,
                    "completion_tokens": 0,
                    "total_tokens": 0,
                    "cost_usd": 0.0,
                    "estimated": False,
                },
            }

        resultado_estrutural = tentar_extracao_estrutural_pdf(path)
        confianca = calcular_confianca_estrutural(resultado_estrutural)
        if tipo == "pdf_escaneado":
            confianca = max(0, confianca - 10)
            debug_events.append("Ajuste aplicado: PDF escaneado detectado, penalidade de confianca estrutural")

        debug_events.append(f"Percepcao do sistema: confianca estrutural = {confianca}")

        if confianca >= limiar_alto and resultado_estrutural.get("itens"):
            itens_finais = _anotar_fonte_origem(resultado_estrutural["itens"], "estrutural", pages)
            fonte_global = "estrutural"
            empresa_det = extrair_razao_social(texto_base)
            debug_events.append(f"Achado: {len(itens_finais)} item(ns) por parser estrutural")
            debug_highlights.append(
                f"RELEVANTE: confianca estrutural alta ({confianca}) -> IA nao foi necessaria"
            )
            return {
                "empresa": empresa_det or os.path.splitext(nome_arquivo)[0],
                "itens": itens_finais,
                "review": review,
                "confianca_estrutural": confianca,
                "fonte_processamento": fonte_global,
                "texto_truncado": False,
                "cnpj": extrair_cnpj(texto_base),
                "telefone": extrair_telefone(texto_base),
                "debug_events": debug_events,
                "debug_highlights": debug_highlights,
                "usage": {
                    "prompt_tokens": 0,
                    "completion_tokens": 0,
                    "total_tokens": 0,
                    "cost_usd": 0.0,
                    "estimated": False,
                },
            }

        llm = call_openrouter_extract(
            texto_base,
            api_key,
            model=model,
            filename_hint=nome_arquivo,
            pre_filtrar=pre_filtrar,
        )
        debug_events.append(f"Biblioteca acionada: requests (modelo IA = {model})")
        texto_truncado = bool(llm.get("texto_truncado"))
        if texto_truncado:
            debug_highlights.append("RELEVANTE: texto truncado para caber no limite enviado a IA")
        itens_ia = _anotar_fonte_origem(llm.get("itens", []), "ia", pages)
        debug_events.append(f"Achado IA: {len(itens_ia)} item(ns) retornado(s)")

        if confianca is not None and confianca < limiar_baixo:
            fonte_global = "ia"
            itens_finais = itens_ia
            debug_highlights.append(
                f"RELEVANTE: confianca estrutural baixa ({confianca}) -> decisao final por IA"
            )
        elif resultado_estrutural.get("itens"):
            fonte_global = "dupla_checagem"
            itens_finais, conflitos = comparar_extracoes(resultado_estrutural["itens"], itens_ia)
            review.extend(conflitos)
            debug_events.append(
                f"Percepcao: dupla checagem executada (parser vs IA), conflitos={len(conflitos)}"
            )
            if conflitos:
                debug_highlights.append(
                    f"RELEVANTE: divergencias detectadas entre parser e IA ({len(conflitos)})"
                )
        else:
            fonte_global = "ia"
            itens_finais = itens_ia
            debug_events.append("Percepcao: sem base estrutural valida, decisao final por IA")

        empresa_det = extrair_razao_social(texto_base)
        empresa_llm = llm.get("empresa")
        if empresa_llm:
            debug_events.append(f"Achado: empresa identificada pela IA = {empresa_llm}")
        elif empresa_det:
            debug_events.append(f"Achado: empresa inferida por heuristica = {empresa_det}")

        if extrair_cnpj(texto_base):
            debug_events.append("Achado: CNPJ identificado no documento")

        return {
            "empresa": empresa_llm or empresa_det or os.path.splitext(nome_arquivo)[0],
            "itens": [normalizar_item(i) for i in itens_finais],
            "review": review,
            "confianca_estrutural": confianca,
            "fonte_processamento": fonte_global,
            "texto_truncado": texto_truncado,
            "cnpj": extrair_cnpj(texto_base),
            "telefone": extrair_telefone(texto_base),
            "debug_events": debug_events,
            "debug_highlights": debug_highlights,
            "usage": llm.get("usage", {
                "prompt_tokens": 0,
                "completion_tokens": 0,
                "total_tokens": 0,
                "cost_usd": 0.0,
                "estimated": True,
            }),
        }

    if tipo == "imagem":
        texto_base, erro_ocr = extract_text_from_image(path)
        debug_events.append(f"Achado: {len(texto_base)} caractere(s) extraido(s) da imagem")

        # --- Fallback de visao para imagem sem texto OCR utilizavel ----------
        if not texto_base.strip() and api_key:
            try:
                with open(path, "rb") as _fh:
                    img_bytes = _fh.read()
                ext_img = os.path.splitext(path)[1].lower().lstrip(".")
                mime_img = f"image/{'jpeg' if ext_img in ('jpg', 'jpeg') else ext_img or 'png'}"
                llm_visao = call_openrouter_vision_extract(
                    [img_bytes], api_key, model=model,
                    filename_hint=nome_arquivo, mime=mime_img,
                )
            except Exception:
                llm_visao = {}
            if llm_visao.get("itens"):
                debug_events.append("Fallback de visao: imagem enviada ao modelo multimodal")
                debug_highlights.append(
                    "RELEVANTE: itens extraidos por VISAO (OCR classico falhou nesta imagem)"
                )
                return {
                    "empresa": llm_visao.get("empresa") or os.path.splitext(nome_arquivo)[0],
                    "itens": [normalizar_item(i) for i in llm_visao["itens"]],
                    "review": review,
                    "confianca_estrutural": None,
                    "fonte_processamento": "visao",
                    "texto_truncado": False,
                    "cnpj": None,
                    "telefone": None,
                    "debug_events": debug_events,
                    "debug_highlights": debug_highlights,
                    "usage": llm_visao.get("usage", {}),
                }

        if erro_ocr and not texto_base.strip():
            debug_highlights.append("RELEVANTE: OCR nao disponivel no ambiente")
            return {
                "empresa": os.path.splitext(nome_arquivo)[0],
                "itens": [],
                "erro": f"Falha no OCR da imagem: {erro_ocr}",
                "review": review,
                "confianca_estrutural": 0,
                "fonte_processamento": "ia",
                "texto_truncado": False,
                "cnpj": None,
                "telefone": None,
                "debug_events": debug_events,
                "debug_highlights": debug_highlights,
                "usage": {
                    "prompt_tokens": 0,
                    "completion_tokens": 0,
                    "total_tokens": 0,
                    "cost_usd": 0.0,
                    "estimated": False,
                },
            }

        if not texto_base.strip():
            return {
                "empresa": os.path.splitext(nome_arquivo)[0],
                "itens": [],
                "erro": "A imagem nao possui texto legivel para extracao.",
                "review": review,
                "confianca_estrutural": 0,
                "fonte_processamento": "ia",
                "texto_truncado": False,
                "cnpj": None,
                "telefone": None,
                "debug_events": debug_events,
                "debug_highlights": debug_highlights,
                "usage": {
                    "prompt_tokens": 0,
                    "completion_tokens": 0,
                    "total_tokens": 0,
                    "cost_usd": 0.0,
                    "estimated": False,
                },
            }

        llm = call_openrouter_extract(
            texto_base,
            api_key,
            model=model,
            filename_hint=nome_arquivo,
            pre_filtrar=pre_filtrar,
        )
        debug_events.append(f"Biblioteca acionada: requests (modelo IA = {model})")
        debug_events.append(f"Achado IA: {len(llm.get('itens', []))} item(ns) retornado(s)")
        itens = _anotar_fonte_origem(llm.get("itens", []), "ia")
        empresa_det = extrair_razao_social(texto_base)
        debug_highlights.append("RELEVANTE: imagem processada por OCR + IA")

        return {
            "empresa": llm.get("empresa") or empresa_det or os.path.splitext(nome_arquivo)[0],
            "itens": itens,
            "review": review,
            "confianca_estrutural": confianca,
            "fonte_processamento": "ia",
            "texto_truncado": bool(llm.get("texto_truncado")),
            "cnpj": extrair_cnpj(texto_base),
            "telefone": extrair_telefone(texto_base),
            "debug_events": debug_events,
            "debug_highlights": debug_highlights,
            "usage": llm.get("usage", {
                "prompt_tokens": 0,
                "completion_tokens": 0,
                "total_tokens": 0,
                "cost_usd": 0.0,
                "estimated": True,
            }),
        }

    texto_base = extract_text(path)
    llm = call_openrouter_extract(
        texto_base,
        api_key,
        model=model,
        filename_hint=nome_arquivo,
        pre_filtrar=pre_filtrar,
    )
    debug_events.append(f"Biblioteca acionada: requests (modelo IA = {model})")
    debug_events.append(f"Achado IA: {len(llm.get('itens', []))} item(ns) retornado(s)")
    itens = _anotar_fonte_origem(llm.get("itens", []), "ia")
    empresa_det = extrair_razao_social(texto_base)
    debug_highlights.append("RELEVANTE: arquivo roteado diretamente para IA")

    return {
        "empresa": llm.get("empresa") or empresa_det or os.path.splitext(nome_arquivo)[0],
        "itens": itens,
        "review": review,
        "confianca_estrutural": confianca,
        "fonte_processamento": fonte_global,
        "texto_truncado": bool(llm.get("texto_truncado")),
        "cnpj": extrair_cnpj(texto_base),
        "telefone": extrair_telefone(texto_base),
        "debug_events": debug_events,
        "debug_highlights": debug_highlights,
        "usage": llm.get("usage", {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
            "cost_usd": 0.0,
            "estimated": True,
        }),
    }
