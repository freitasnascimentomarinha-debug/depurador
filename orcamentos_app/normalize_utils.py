"""Utilitarios deterministicos de normalizacao e validacao de campos extraidos."""
from __future__ import annotations

import re
import unicodedata
from typing import Any

SINONIMOS_UNIDADE = {
    "UN": ["UN", "UND", "UNID", "UNIDADE", "UNIDADES", "PC", "PCA", "PÇ", "PECA", "PEÇA"],
    "CX": ["CX", "CAIXA", "CAIXAS"],
    "KG": ["KG", "QUILO", "QUILOGRAMA", "QUILOGRAMAS"],
    "L": ["L", "LT", "LITRO", "LITROS"],
    "M": ["M", "MT", "METRO", "METROS"],
    "KIT": ["KIT", "KITS", "CJ", "CONJUNTO", "CONJUNTOS"],
    "EMB": ["EMB", "EMBALAGEM", "EMBALAGENS"],
    "FR": ["FR", "FRASCO", "FRASCOS", "FD"],
    "RL": ["RL", "ROLO", "ROLOS"],
    "G": ["G", "GR", "GRAMA", "GRAMAS"],
    "ML": ["ML", "MILILITRO", "MILILITROS"],
    "PAR": ["PAR", "PARES"],
}

REGEX_CNPJ = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
REGEX_VALOR_BR = re.compile(r"R?\$?\s?\d{1,3}(?:\.\d{3})*,\d{2}")


def _sem_acento(texto: str) -> str:
    base = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in base if not unicodedata.combining(ch))


def normalizar_unidade(valor: str | None) -> str | None:
    """Mapeia unidade em texto livre para forma canonica."""
    if valor is None:
        return None
    texto = _sem_acento(str(valor)).upper()
    texto = re.sub(r"\s+", " ", texto).strip()
    if not texto:
        return None

    # Regra de seguranca: UF nunca deve parecer valor monetario/numerico.
    if "R$" in texto or re.search(r"\d", texto):
        return None

    for canonica, sinonimos in SINONIMOS_UNIDADE.items():
        if texto == canonica or texto in sinonimos:
            return canonica

    # Valores desconhecidos ainda podem existir, mas mantemos apenas codigos curtos.
    if len(texto) <= 8 and re.fullmatch(r"[A-ZÇ]+", texto):
        return texto
    return None


def extrair_cnpj(texto: str) -> str | None:
    """Extrai o primeiro CNPJ encontrado no texto bruto."""
    if not texto:
        return None
    match = REGEX_CNPJ.search(texto)
    if not match:
        return None
    cnpj = re.sub(r"\D", "", match.group(0))
    if len(cnpj) != 14:
        return None
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"


def extrair_razao_social(texto: str) -> str | None:
    """Busca linha candidata a razao social perto de ocorrencia de CNPJ."""
    if not texto:
        return None

    linhas = [re.sub(r"\s+", " ", l).strip() for l in texto.splitlines() if l.strip()]
    if not linhas:
        return None

    idx_cnpj = None
    for idx, linha in enumerate(linhas):
        if REGEX_CNPJ.search(linha):
            idx_cnpj = idx
            break

    if idx_cnpj is None:
        return None

    termos_empresa = ("LTDA", "S/A", "S.A.", "EIRELI", "ME", "EPP", "SOCIEDADE", "EMPRESA")
    janela_ini = max(0, idx_cnpj - 4)
    janela_fim = min(len(linhas), idx_cnpj + 5)
    candidatas = linhas[janela_ini:janela_fim]

    for linha in candidatas:
        linha_up = _sem_acento(linha).upper()
        if any(t in linha_up for t in termos_empresa):
            return linha[:180]

    return None


def parse_valor_brl(texto: Any) -> float | None:
    """Converte valor BRL textual para float (1234.56)."""
    if texto is None:
        return None

    if isinstance(texto, (int, float)):
        return float(texto)

    bruto = str(texto).strip()
    if not bruto:
        return None

    achado = REGEX_VALOR_BR.search(bruto)
    if achado:
        bruto = achado.group(0)

    bruto = bruto.replace("R$", "").replace("$", "")
    bruto = re.sub(r"\s+", "", bruto)
    bruto = bruto.replace(".", "").replace(",", ".")
    bruto = re.sub(r"[^\d.-]", "", bruto)

    if not bruto:
        return None

    try:
        return float(bruto)
    except ValueError:
        return None


def limpar_quebras_e_caracteres(texto: str) -> str:
    """Normaliza espacos e remove caracteres de controle de um texto livre."""
    if not texto:
        return ""

    texto = texto.replace("\r", "\n")
    texto = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", " ", texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def normalizar_item(item: dict) -> dict:
    """Normaliza campos basicos de um item extraido sem alterar semantica."""
    saida = dict(item)
    saida["descricao"] = limpar_quebras_e_caracteres(str(saida.get("descricao") or ""))
    saida["unidade"] = normalizar_unidade(saida.get("unidade"))

    if saida.get("quantidade") is not None:
        saida["quantidade"] = parse_valor_brl(saida.get("quantidade"))
    if saida.get("preco_unitario") is not None:
        saida["preco_unitario"] = parse_valor_brl(saida.get("preco_unitario"))
    if saida.get("preco_total") is not None:
        saida["preco_total"] = parse_valor_brl(saida.get("preco_total"))

    numero_item = saida.get("numero_item")
    if numero_item is not None:
        saida["numero_item"] = str(numero_item).strip() or None

    if not saida.get("origem"):
        saida["origem"] = "desconhecida"
    if not saida.get("fonte_extracao"):
        saida["fonte_extracao"] = "ia"

    return saida
