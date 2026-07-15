"""
Classificação de emails via LLM (OpenRouter).

Categorias:
  pedido_orcamento   – Marinha/empresa enviando pedido de orçamento a fornecedores
  orcamento_recebido – Fornecedor respondendo com orçamento (com ou sem anexo)
  confirmacao_leitura– Confirmação de leitura / aviso de entrega (MDN)
  declinio           – Fornecedor recusando participar
  duvida             – Fornecedor fazendo perguntas sobre o processo ou itens
  tramite_interno    – Tráfego interno (Marinha na cópia, não é o destinatário principal)
  outro              – Qualquer outra coisa
"""
from __future__ import annotations

import json
import re
import time
from typing import Optional

import requests

MODEL_PRICING_PER_MILLION = {
    "google/gemini-2.5-flash": {"input": 0.30, "output": 2.50},
    "openai/gpt-5-mini": {"input": 0.25, "output": 2.00},
    "deepseek/deepseek-chat-v3.2": {"input": 0.25, "output": 0.40},
    # legados
    "openai/gpt-4o-mini": {"input": 0.15, "output": 0.60},
    "anthropic/claude-3-haiku": {"input": 0.25, "output": 1.25},
}

# Classificacao e tarefa simples: use o modelo mais barato disponivel.
DEFAULT_CLASSIFIER_MODEL = "google/gemini-2.5-flash"

# JSON Schema para structured outputs na classificacao
CLASSIFICATION_JSON_SCHEMA = {
    "name": "classificacao_email",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "tipo": {
                "type": "string",
                "enum": ["pedido_orcamento", "orcamento_recebido", "confirmacao_leitura",
                          "declinio", "duvida", "tramite_interno", "outro"],
            },
            "confianca": {"type": "number"},
            "resumo": {"type": "string"},
            "numero_processo": {"type": ["string", "null"]},
        },
        "required": ["tipo", "confianca", "resumo", "numero_processo"],
        "additionalProperties": False,
    },
}

TIPOS_VALIDOS = {
    "pedido_orcamento",
    "orcamento_recebido",
    "confirmacao_leitura",
    "declinio",
    "duvida",
    "tramite_interno",
    "outro",
}

_ROTULOS_PT = {
    "pedido_orcamento":    "📤 Pedido de Orçamento",
    "orcamento_recebido":  "💰 Orçamento Recebido",
    "confirmacao_leitura": "👁️ Confirmação de Leitura",
    "declinio":            "🚫 Declínio",
    "duvida":              "❓ Dúvida",
    "tramite_interno":     "🏢 Trâmite Interno",
    "outro":               "📄 Outro",
}

_CORES = {
    "pedido_orcamento":    "#1565C0",
    "orcamento_recebido":  "#2E7D32",
    "confirmacao_leitura": "#6A1B9A",
    "declinio":            "#BF360C",
    "duvida":              "#E65100",
    "tramite_interno":     "#37474F",
    "outro":               "#757575",
}


def rotulo_tipo(tipo: str) -> str:
    return _ROTULOS_PT.get(tipo, tipo)


def cor_tipo(tipo: str) -> str:
    return _CORES.get(tipo, "#757575")


# ---------------------------------------------------------------------------
# Classificação heurística (sem LLM, rápida)
# ---------------------------------------------------------------------------

def _heuristica(parsed: dict) -> Optional[str]:
    """
    Tenta classificar sem chamar a IA.
    Retorna tipo ou None se não tiver certeza.
    """
    if parsed.get("is_mdn"):
        return "confirmacao_leitura"

    assunto = (parsed.get("assunto") or "").lower()
    corpo = (parsed.get("corpo") or "").lower()[:1500]
    remetente = (parsed.get("remetente_email") or "").lower()

    # Confirmação de leitura por assunto
    if any(k in assunto for k in ("lida:", "lido:", "confirmação de leitura", "read receipt",
                                   "delivery notification", "entregue:", "recebida:")):
        return "confirmacao_leitura"

    # Pedido de orçamento (remetente governa: domínio .mil.br ou endereços institucionais conhecidos)
    _REMETENTES_INSTITUCIONAIS = {"sobressalentes.comrj@gmail.com"}
    if re.search(r"\.mil\.br$|\.mar\.mil\.br$|\.marinha\.mil\.br$", remetente) or remetente in _REMETENTES_INSTITUCIONAIS:
        if any(k in assunto + corpo for k in ("orçamento", "cotação", "cotacao", "proposta", "solicita")):
            return "pedido_orcamento"

    # Declínio
    decl_words = ("não podemos", "nao podemos", "impossibilitados", "sem condições",
                  "sem condicoes", "declino", "declínio", "não participar", "nao participar",
                  "não enviaremos", "nao enviaremos", "desculpe", "lamentamos")
    if any(w in corpo for w in decl_words):
        return "declinio"

    # Orçamento recebido: resposta com conteúdo de proposta/preço
    orcamento_words = ("proposta de preços", "proposta de precos", "nossa proposta",
                       "segue proposta", "segue cotação", "segue cotacao",
                       "conforme solicitado", "valor unitário", "valor unitario",
                       "preço unitário", "preco unitario", "r$/un", "r$/pct",
                       "nossa cotação", "nossa cotacao")
    if any(w in corpo for w in orcamento_words):
        return "orcamento_recebido"

    # Dúvida do fornecedor: pergunta sobre itens
    duvida_words = ("gostaria de esclarecimento", "poderia esclarecer", "qual a especificação",
                    "qual a especificacao", "o que é o item", "o que e o item",
                    "pode detalhar", "poderia detalhar", "prazo de entrega",
                    "em qual cidade", "local de entrega", "pergunta sobre")
    if any(w in corpo for w in duvida_words):
        return "duvida"

    return None


# ---------------------------------------------------------------------------
# Classificação via LLM
# ---------------------------------------------------------------------------

_PROMPT_SISTEMA = """
Você é um assistente especializado em classificar emails de processos licitatórios
da Marinha do Brasil. Classifique o email fornecido em UMA das categorias:

pedido_orcamento   – Email enviado pela Marinha/empresa solicitando orçamento ou proposta de preços a fornecedores
orcamento_recebido – Fornecedor responde com orçamento, proposta ou preços (com ou sem arquivo anexo)
confirmacao_leitura– Aviso automático de leitura, confirmação de entrega ou acuse de recebimento
declinio           – Fornecedor informa que não participará / não enviará orçamento
duvida             – Fornecedor faz perguntas sobre itens, especificações ou o processo
tramite_interno    – Email interno da empresa/órgão, Marinha apenas em cópia, não relacionado ao orçamento
outro              – Qualquer outra situação

Responda APENAS com JSON (sem markdown) no formato:
{
  "tipo": "<categoria>",
  "confianca": <0 a 100>,
  "resumo": "<1 a 2 frases resumindo o email>",
  "numero_processo": "<número extraído do assunto ou null>"
}
""".strip()


def classificar_email(
    parsed: dict,
    api_key: str,
    model: str = None,
    timeout: int = 30,
    max_retries: int = 2,
) -> dict:
    """
    Classifica um email. Tenta heurística primeiro; se não concluir chama LLM.
    Retorna {'tipo', 'confianca', 'resumo', 'numero_processo', 'uso_ia', 'usage'}.
    """
    model = model or DEFAULT_CLASSIFIER_MODEL

    # Tenta heurística rápida
    tipo_heur = _heuristica(parsed)
    if tipo_heur:
        return {
            "tipo": tipo_heur,
            "confianca": 90,
            "resumo": "",
            "numero_processo": None,
            "uso_ia": False,
            "usage": {},
        }

    # Monta contexto para o LLM (truncado para economizar tokens)
    corpo_curto = (parsed.get("corpo") or "")[:2500]
    destinatarios_str = ", ".join(
        f'{d.get("nome") or ""} <{d.get("email") or ""}>'.strip()
        for d in (parsed.get("destinatarios") or [])[:6]
    )
    user_msg = f"""Assunto: {parsed.get("assunto", "")}
Remetente: {parsed.get("remetente_nome", "")} <{parsed.get("remetente_email", "")}>
Destinatários: {destinatarios_str}
Tem anexo: {"Sim" if parsed.get("tem_anexo") else "Não"}
{"Nomes dos anexos: " + ", ".join(parsed.get("nomes_anexos", [])) if parsed.get("nomes_anexos") else ""}

Corpo do email:
{corpo_curto}"""

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _PROMPT_SISTEMA},
            {"role": "user", "content": user_msg},
        ],
        "max_tokens": 300,
        "temperature": 0,
        "usage": {"include": True},
        "response_format": {"type": "json_schema", "json_schema": CLASSIFICATION_JSON_SCHEMA},
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/freitasnascimentomarinha-debug/depurador",
    }

    for attempt in range(max_retries + 1):
        try:
            resp = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=timeout,
            )
            if resp.status_code in (400, 404) and "response_format" in payload:
                # Provedor/modelo sem suporte a structured outputs: refaz sem schema
                payload = {k: v for k, v in payload.items() if k != "response_format"}
                resp = requests.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=timeout,
                )
            resp.raise_for_status()
            data = resp.json()
            content = data["choices"][0]["message"]["content"].strip()
            # limpa markdown se o modelo insistir
            content = re.sub(r"^```json\s*|```\s*$", "", content, flags=re.MULTILINE).strip()
            result = json.loads(content)
            tipo = result.get("tipo", "outro")
            if tipo not in TIPOS_VALIDOS:
                tipo = "outro"

            usage_raw = data.get("usage") or {}
            usage = _calcular_custo(usage_raw, model)

            return {
                "tipo": tipo,
                "confianca": int(result.get("confianca") or 0),
                "resumo": str(result.get("resumo") or ""),
                "numero_processo": result.get("numero_processo"),
                "uso_ia": True,
                "usage": usage,
            }
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt < max_retries:
                time.sleep(1)
                continue
            # fallback sem IA
            return {
                "tipo": "outro",
                "confianca": 0,
                "resumo": "Classificação indisponível (timeout).",
                "numero_processo": None,
                "uso_ia": False,
                "usage": {},
            }
        except json.JSONDecodeError:
            # LLM retornou texto não-JSON
            return {
                "tipo": "outro",
                "confianca": 0,
                "resumo": "Não foi possível classificar (resposta inválida da IA).",
                "numero_processo": None,
                "uso_ia": False,
                "usage": {},
            }
        except Exception as exc:
            return {
                "tipo": "outro",
                "confianca": 0,
                "resumo": f"Erro na classificação: {exc}",
                "numero_processo": None,
                "uso_ia": False,
                "usage": {},
            }


def _calcular_custo(usage_raw: dict, model: str) -> dict:
    def _i(k):
        try:
            return int(usage_raw.get(k) or 0)
        except (TypeError, ValueError):
            return 0

    prompt_tokens = _i("prompt_tokens") or _i("input_tokens")
    completion_tokens = _i("completion_tokens") or _i("output_tokens")
    total_tokens = _i("total_tokens") or (prompt_tokens + completion_tokens)

    cost_usd = usage_raw.get("cost") or usage_raw.get("total_cost")
    estimated = False
    if cost_usd is None:
        pricing = MODEL_PRICING_PER_MILLION.get(model)
        if pricing:
            cost_usd = (
                (prompt_tokens / 1_000_000.0) * pricing["input"]
                + (completion_tokens / 1_000_000.0) * pricing["output"]
            )
        else:
            cost_usd = 0.0
        estimated = True

    return {
        "prompt_tokens": prompt_tokens,
        "completion_tokens": completion_tokens,
        "total_tokens": total_tokens,
        "cost_usd": float(cost_usd or 0),
        "estimated": estimated,
    }
