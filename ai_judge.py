"""
IA juiz para a zona cinzenta do casamento de itens.

Pares de descrições com similaridade intermediária (entre zona_min e o limiar
fuzzy) não são nem casados nem descartados automaticamente — são enviados em
lote para um LLM barato decidir, com JSON Schema estrito. Uma chamada julga
até BATCH_SIZE pares, mantendo o custo em centavos.
"""
from __future__ import annotations

import json
import re
import time

import requests

DEFAULT_JUDGE_MODEL = "google/gemini-2.5-flash"
BATCH_SIZE = 20

_SYSTEM_PROMPT = """Você é um especialista em compras públicas e catalogação de materiais.
Receberá pares de descrições de itens vindos de orçamentos de fornecedores diferentes.
Para cada par, decida se as duas descrições se referem ao MESMO item/produto do processo
de compra (mesmo que uma seja mais detalhada, abreviada ou use sinônimos técnicos).

Regras:
- Unidades de fornecimento incompatíveis (ex.: PCT vs KG) indicam itens DIFERENTES.
- Quantidades muito divergentes sugerem itens diferentes, mas não são decisivas sozinhas.
- Marca/modelo diferente do mesmo tipo de produto ainda pode ser o mesmo item do edital.
- Na dúvida real, responda mesmo_item=false com confianca baixa (falso positivo em mapa
  de preços oficial é pior que falso negativo).
"""

JUDGE_JSON_SCHEMA = {
    "name": "julgamento_pares",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "decisoes": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "id": {"type": "integer"},
                        "mesmo_item": {"type": "boolean"},
                        "confianca": {"type": "number"},
                    },
                    "required": ["id", "mesmo_item", "confianca"],
                    "additionalProperties": False,
                },
            }
        },
        "required": ["decisoes"],
        "additionalProperties": False,
    },
}


def _formatar_par(par: dict) -> str:
    linhas = [f"PAR {par['id']}:"]
    linhas.append(f"  A: {par.get('descricao_a', '')}")
    if par.get("unidade_a") or par.get("quantidade_a") is not None:
        linhas.append(f"     UF: {par.get('unidade_a') or '?'} | QTD: {par.get('quantidade_a') if par.get('quantidade_a') is not None else '?'}")
    linhas.append(f"  B: {par.get('descricao_b', '')}")
    if par.get("unidade_b") or par.get("quantidade_b") is not None:
        linhas.append(f"     UF: {par.get('unidade_b') or '?'} | QTD: {par.get('quantidade_b') if par.get('quantidade_b') is not None else '?'}")
    return "\n".join(linhas)


def julgar_pares(
    pares: list[dict],
    api_key: str,
    model: str | None = None,
    timeout: int = 60,
    max_retries: int = 2,
) -> tuple[dict[int, dict], dict]:
    """
    Julga pares de descrições em lotes.

    pares: [{"id", "descricao_a", "descricao_b", "unidade_a", "unidade_b",
             "quantidade_a", "quantidade_b"}, ...]

    Retorna:
      decisoes: {id: {"mesmo_item": bool, "confianca": float}}
      usage_total: {"total_tokens": int, "cost_usd": float, "chamadas": int}
    """
    decisoes: dict[int, dict] = {}
    usage_total = {"total_tokens": 0, "cost_usd": 0.0, "chamadas": 0}
    if not pares or not api_key:
        return decisoes, usage_total

    model = model or DEFAULT_JUDGE_MODEL
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    for inicio in range(0, len(pares), BATCH_SIZE):
        lote = pares[inicio:inicio + BATCH_SIZE]
        user_msg = (
            "Julgue os pares abaixo e responda no formato JSON especificado.\n\n"
            + "\n\n".join(_formatar_par(p) for p in lote)
        )
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": user_msg},
            ],
            "temperature": 0,
            "usage": {"include": True},
            "response_format": {"type": "json_schema", "json_schema": JUDGE_JSON_SCHEMA},
        }

        for attempt in range(max_retries + 1):
            try:
                resp = requests.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers=headers, json=payload, timeout=timeout,
                )
                if resp.status_code in (400, 404) and "response_format" in payload:
                    payload = {k: v for k, v in payload.items() if k != "response_format"}
                    resp = requests.post(
                        "https://openrouter.ai/api/v1/chat/completions",
                        headers=headers, json=payload, timeout=timeout,
                    )
                resp.raise_for_status()
                data = resp.json()
                content = data["choices"][0]["message"]["content"].strip()
                content = re.sub(r"^```json\s*|```\s*$", "", content, flags=re.MULTILINE).strip()
                parsed = json.loads(content)
                for d in parsed.get("decisoes", []):
                    try:
                        decisoes[int(d["id"])] = {
                            "mesmo_item": bool(d["mesmo_item"]),
                            "confianca": float(d.get("confianca", 0)),
                        }
                    except (KeyError, TypeError, ValueError):
                        continue
                usage = data.get("usage") or {}
                usage_total["total_tokens"] += int(usage.get("total_tokens") or 0)
                try:
                    usage_total["cost_usd"] += float(usage.get("cost") or 0)
                except (TypeError, ValueError):
                    pass
                usage_total["chamadas"] += 1
                break
            except (requests.RequestException, json.JSONDecodeError, KeyError, IndexError):
                if attempt < max_retries:
                    time.sleep(2 ** attempt)
                # se esgotar as tentativas, os pares deste lote ficam sem decisão
                # (comportamento seguro: não casar)

    return decisoes, usage_total
