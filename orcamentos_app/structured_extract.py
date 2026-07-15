"""Extracao estrutural deterministica e roteamento por tipo de arquivo."""
from __future__ import annotations

import os
import re
import unicodedata
from typing import Any

from normalize_utils import limpar_quebras_e_caracteres, normalizar_item, parse_valor_brl

_HEADER_KEYWORDS = {
    "item", "descricao", "qtd", "quantidade", "valor", "preco", "unitario"
}

_COL_SYNONYMS = {
    "numero_item": ["item", "num", "numero", "n", "cod", "codigo"],
    "descricao": ["descricao", "especificacao", "produto", "objeto", "material"],
    "unidade": ["uf", "und", "unidade", "emb", "cx", "kg", "lt"],
    "quantidade": ["qtd", "quantidade", "quant", "qtde"],
    "preco_unitario": ["valor unit", "vl unit", "unitario", "preco unit", "valor"],
    "preco_total": ["valor total", "vl total", "total"],
}

_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
_ITEM_LINE_RE = re.compile(r"^\d{1,4}$")
_PI_LINE_RE = re.compile(r"^\d{6,12}$")
_PRICE_LINE_RE = re.compile(r"^R\$\s*")
_PRAZO_RE = re.compile(r"\bsemana", re.IGNORECASE)


def _import_fitz():
    raise RuntimeError("PyMuPDF desativado neste build")


def _import_pandas():
    raise RuntimeError("pandas desativado neste build")


def _import_document():
    from docx import Document
    return Document


def _import_openpyxl():
    import openpyxl
    return openpyxl


def _import_pdfplumber():
    try:
        import pdfplumber
    except Exception:  # pragma: no cover - dependencia opcional
        return None
    return pdfplumber


def _sem_acento(texto: str) -> str:
    base = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in base if not unicodedata.combining(ch))


def _norm(texto: Any) -> str:
    if texto is None:
        return ""
    txt = _sem_acento(str(texto)).lower()
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def detectar_tipo_e_rotear(path: str) -> str:
    """Retorna: xlsx, docx, pdf_texto, pdf_escaneado, imagem."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return "xlsx"
    if ext in (".docx", ".doc"):
        return "docx"
    if ext in _IMAGE_EXT:
        return "imagem"
    if ext != ".pdf":
        return "pdf_texto"

    pdfplumber = _import_pdfplumber()
    if pdfplumber is None:
        return "pdf_escaneado"

    with pdfplumber.open(path) as pdf:
        paginas = len(pdf.pages)
        if paginas == 0:
            return "pdf_escaneado"
        com_texto = 0
        for page in pdf.pages:
            if (page.extract_text() or "").strip():
                com_texto += 1
        return "pdf_texto" if com_texto >= max(1, paginas // 2) else "pdf_escaneado"


def _detectar_linha_cabecalho(rows: list[list[Any]]) -> int | None:
    for i in range(min(len(rows), 40)):
        linha = rows[i]
        col_hits = 0
        for val in linha:
            norm = _norm(val)
            if not norm:
                continue
            tokens = set(re.findall(r"\w+", norm))
            if tokens.intersection(_HEADER_KEYWORDS):
                col_hits += 1
        if col_hits >= 2:
            return i
    return None


def _mapear_colunas(headers: list[str]) -> dict[str, int]:
    def _match_term(header_norm: str, term: str) -> bool:
        term = _norm(term)
        if not term:
            return False

        if " " in term:
            return term in header_norm

        tokens = set(re.findall(r"\w+", header_norm))
        return term in tokens

    mapeamento = {}
    for idx, cab in enumerate(headers):
        norm = _norm(cab)
        if not norm:
            continue
        for campo, termos in _COL_SYNONYMS.items():
            if campo in mapeamento:
                continue
            if any(_match_term(norm, t) for t in termos):
                mapeamento[campo] = idx
    return mapeamento


def _extrair_itens_pdf_por_linhas(path: str) -> dict:
    """Fallback para PDFs com tabela textual multi-pagina sem cabecalho repetido."""
    pdfplumber = _import_pdfplumber()
    if pdfplumber is None:
        return {"itens": [], "header_reconhecido": False}

    with pdfplumber.open(path) as pdf:
        linhas_tabela = []
        header_reconhecido = False
        em_secao_preco = False

        for p_idx, page in enumerate(pdf.pages, start=1):
            linhas = [l.strip() for l in (page.extract_text() or "").splitlines() if l.strip()]
            for linha in linhas:
                if "4.1 Preço" in linha or "4.1 Preco" in linha:
                    em_secao_preco = True
                if not em_secao_preco:
                    continue
                if linha == "TOTAL" or linha.startswith("4.2 ") or "4.2 Impostos" in linha:
                    em_secao_preco = False
                    break
                if linha == "Item":
                    header_reconhecido = True
                    continue
                if not header_reconhecido:
                    continue
                if linha in {
                    "PI", "Referência cadastrada", "Referencia cadastrada", "Descrição do item",
                    "Descricao do item", "STF", "Prazo de", "entrega", "Qnt",
                    "Valor Unitário", "Valor Unitario", "Valor Total"
                }:
                    continue
                if linha in {"PROPOSTA", "REV 00"} or linha.startswith("PTC ") or linha.startswith("Página "):
                    continue
                linhas_tabela.append((p_idx, linha))

        itens = []
        idx = 0
        while idx < len(linhas_tabela):
            pagina, linha = linhas_tabela[idx]
            proxima_linha = linhas_tabela[idx + 1][1] if idx + 1 < len(linhas_tabela) else ""
            if not (_ITEM_LINE_RE.fullmatch(linha) and _PI_LINE_RE.fullmatch(proxima_linha)):
                idx += 1
                continue

            numero_item = linha
            idx += 2  # pula numero do item e PI
            descricao_partes = []
            quantidade = None
            preco_unitario = None
            preco_total = None

            while idx < len(linhas_tabela):
                pagina_atual, linha_atual = linhas_tabela[idx]
                prox = linhas_tabela[idx + 1][1] if idx + 1 < len(linhas_tabela) else ""

                if _ITEM_LINE_RE.fullmatch(linha_atual) and _PI_LINE_RE.fullmatch(prox):
                    break

                if linha_atual == "TOTAL" or linha_atual.startswith("4.2 "):
                    break

                if quantidade is None and re.fullmatch(r"\d+(?:[\.,]\d+)?", linha_atual) and _PRICE_LINE_RE.match(prox):
                    quantidade = parse_valor_brl(linha_atual)
                    idx += 1
                    continue

                if preco_unitario is None and _PRICE_LINE_RE.match(linha_atual):
                    preco_unitario = parse_valor_brl(linha_atual)
                    idx += 1
                    continue

                if preco_unitario is not None and preco_total is None and _PRICE_LINE_RE.match(linha_atual):
                    preco_total = parse_valor_brl(linha_atual)
                    idx += 1
                    continue

                if _PRAZO_RE.search(linha_atual):
                    idx += 1
                    continue

                if linha_atual.startswith("STF"):
                    idx += 1
                    while idx < len(linhas_tabela):
                        _, norma = linhas_tabela[idx]
                        prox_norma = linhas_tabela[idx + 1][1] if idx + 1 < len(linhas_tabela) else ""
                        if _PRAZO_RE.search(norma) or (
                            re.fullmatch(r"\d+(?:[\.,]\d+)?", norma) and _PRICE_LINE_RE.match(prox_norma)
                        ):
                            break
                        idx += 1
                    continue

                descricao_partes.append(linha_atual)
                pagina = pagina_atual
                idx += 1

            descricao = limpar_quebras_e_caracteres(" ".join(descricao_partes))
            if descricao and preco_unitario is not None:
                itens.append(normalizar_item({
                    "numero_item": numero_item,
                    "descricao": descricao,
                    "unidade": None,
                    "quantidade": quantidade,
                    "preco_unitario": preco_unitario,
                    "preco_total": preco_total,
                    "fonte_extracao": "estrutural",
                    "origem": f"Pagina {pagina}",
                }))

        return {
            "itens": itens,
            "header_reconhecido": header_reconhecido,
        }


def extrair_xlsx_estruturado(path: str) -> list[dict] | None:
    """Extrai itens de XLSX/XLS por leitura tabular deterministica."""
    openpyxl = _import_openpyxl()
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return None

    itens = []
    for sheet in wb.worksheets:
        nome_aba = sheet.title
        rows = [list(r) for r in sheet.iter_rows(values_only=True)]
        if not rows:
            continue

        header_idx = _detectar_linha_cabecalho(rows)
        if header_idx is None:
            continue

        headers = [str(x) if x is not None else "" for x in rows[header_idx]]
        colunas = _mapear_colunas(headers)
        if len(colunas) < 2 or "descricao" not in colunas:
            continue

        for ridx in range(header_idx + 1, len(rows)):
            row = rows[ridx]
            descricao = row[colunas["descricao"]] if "descricao" in colunas and colunas["descricao"] < len(row) else None
            descricao = limpar_quebras_e_caracteres(str(descricao or ""))
            if not descricao:
                continue

            item = {
                "numero_item": (
                    str(row[colunas["numero_item"]]).strip()
                    if "numero_item" in colunas and colunas["numero_item"] < len(row) and row[colunas["numero_item"]] is not None
                    else None
                ),
                "descricao": descricao,
                "unidade": (
                    str(row[colunas["unidade"]]).strip()
                    if "unidade" in colunas and colunas["unidade"] < len(row) and row[colunas["unidade"]] is not None
                    else None
                ),
                "quantidade": parse_valor_brl(row[colunas["quantidade"]]) if "quantidade" in colunas and colunas["quantidade"] < len(row) else None,
                "preco_unitario": parse_valor_brl(row[colunas["preco_unitario"]]) if "preco_unitario" in colunas and colunas["preco_unitario"] < len(row) else None,
                "preco_total": parse_valor_brl(row[colunas["preco_total"]]) if "preco_total" in colunas and colunas["preco_total"] < len(row) else None,
                "fonte_extracao": "estrutural",
                "origem": f"{nome_aba}!{ridx + 1}",
            }
            itens.append(normalizar_item(item))

    return itens or None


def extrair_docx_estruturado(path: str) -> list[dict] | None:
    """Extrai itens de DOCX por tabelas reconhecidas por cabecalho."""
    Document = _import_document()
    try:
        doc = Document(path)
    except Exception:
        return None

    itens = []
    for t_idx, tabela in enumerate(doc.tables, start=1):
        linhas = [[cell.text.strip() for cell in row.cells] for row in tabela.rows]
        if not linhas:
            continue

        header_idx = _detectar_linha_cabecalho(linhas)
        if header_idx is None:
            continue

        headers = [str(x) if x is not None else "" for x in linhas[header_idx]]
        colunas = _mapear_colunas(headers)
        if len(colunas) < 2 or "descricao" not in colunas:
            continue

        for l_idx in range(header_idx + 1, len(linhas)):
            row = linhas[l_idx]
            descricao = row[colunas["descricao"]] if "descricao" in colunas and colunas["descricao"] < len(row) else None
            descricao = limpar_quebras_e_caracteres(str(descricao or ""))
            if not descricao:
                continue

            item = {
                "numero_item": (
                    str(row[colunas["numero_item"]]).strip()
                    if "numero_item" in colunas and colunas["numero_item"] < len(row) and str(row[colunas["numero_item"]]).strip()
                    else None
                ),
                "descricao": descricao,
                "unidade": (
                    str(row[colunas["unidade"]]).strip()
                    if "unidade" in colunas and colunas["unidade"] < len(row) and str(row[colunas["unidade"]]).strip()
                    else None
                ),
                "quantidade": parse_valor_brl(row[colunas["quantidade"]]) if "quantidade" in colunas and colunas["quantidade"] < len(row) else None,
                "preco_unitario": parse_valor_brl(row[colunas["preco_unitario"]]) if "preco_unitario" in colunas and colunas["preco_unitario"] < len(row) else None,
                "preco_total": parse_valor_brl(row[colunas["preco_total"]]) if "preco_total" in colunas and colunas["preco_total"] < len(row) else None,
                "fonte_extracao": "estrutural",
                "origem": f"Tabela {t_idx}, linha {l_idx + 1}",
            }
            itens.append(normalizar_item(item))

    return itens or None


def tentar_extracao_estrutural_pdf(path: str) -> dict:
    """Tenta extrair tabelas de PDF com pdfplumber e retorna metadados para score."""
    pdfplumber = _import_pdfplumber()
    itens = []
    encontrou_tabela = False
    encontrou_colunas = False
    blocos_item = 0

    if pdfplumber is None:
        return {
            "itens": itens,
            "encontrou_tabela": False,
            "encontrou_colunas": False,
            "linhas_extraidas": 0,
            "blocos_item": 0,
        }

    with pdfplumber.open(path) as pdf:
        for p_idx, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            blocos_item += len(re.findall(r"\bitem\s*\d+\b", page_text, flags=re.IGNORECASE))

            tabelas = page.extract_tables() or []
            if not tabelas:
                continue
            encontrou_tabela = True

            for tabela in tabelas:
                if not tabela:
                    continue
                header_idx = _detectar_linha_cabecalho(tabela)
                if header_idx is None:
                    continue

                headers = [str(x) if x is not None else "" for x in tabela[header_idx]]
                colunas = _mapear_colunas(headers)
                if "descricao" in colunas and "quantidade" in colunas and ("preco_unitario" in colunas or "preco_total" in colunas):
                    encontrou_colunas = True

                for ridx in range(header_idx + 1, len(tabela)):
                    row = tabela[ridx] or []
                    descricao = row[colunas["descricao"]] if "descricao" in colunas and colunas["descricao"] < len(row) else None
                    descricao = limpar_quebras_e_caracteres(str(descricao or ""))
                    if not descricao:
                        continue

                    item = {
                        "numero_item": (
                            str(row[colunas["numero_item"]]).strip()
                            if "numero_item" in colunas and colunas["numero_item"] < len(row) and str(row[colunas["numero_item"]]).strip()
                            else None
                        ),
                        "descricao": descricao,
                        "unidade": (
                            str(row[colunas["unidade"]]).strip()
                            if "unidade" in colunas and colunas["unidade"] < len(row) and str(row[colunas["unidade"]]).strip()
                            else None
                        ),
                        "quantidade": parse_valor_brl(row[colunas["quantidade"]]) if "quantidade" in colunas and colunas["quantidade"] < len(row) else None,
                        "preco_unitario": parse_valor_brl(row[colunas["preco_unitario"]]) if "preco_unitario" in colunas and colunas["preco_unitario"] < len(row) else None,
                        "preco_total": parse_valor_brl(row[colunas["preco_total"]]) if "preco_total" in colunas and colunas["preco_total"] < len(row) else None,
                        "fonte_extracao": "estrutural",
                        "origem": f"Pagina {p_idx}",
                    }
                    itens.append(normalizar_item(item))

    resultado_linhas = _extrair_itens_pdf_por_linhas(path)
    if len(resultado_linhas.get("itens", [])) > len(itens):
        itens = resultado_linhas["itens"]
        if resultado_linhas.get("header_reconhecido"):
            encontrou_colunas = True

    return {
        "itens": itens,
        "encontrou_tabela": encontrou_tabela,
        "encontrou_colunas": encontrou_colunas,
        "linhas_extraidas": len(itens),
        "blocos_item": blocos_item,
    }
