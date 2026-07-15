"""
Lógica de casamento (matching) de itens entre orçamentos de empresas diferentes.

Prioridade:
1. Número do item (do edital/TR) -> casamento exato, alta confiança.
2. Descrição -> correspondência aproximada (fuzzy), confiança média,
   e registrada na lista de revisão manual.
"""
from collections import defaultdict
from statistics import median

from normalize_utils import normalizar_unidade
from text_similarity import token_sort_ratio


def normalize_desc(desc: str) -> str:
    if not desc:
        return ""
    return " ".join(desc.strip().lower().split())


def build_master_from_consensus(all_extractions, min_agree: int = 3, consensus_threshold: int = 80):
    """
    Gera uma tabela mestre automaticamente, sem precisar do usuário subir uma.

    Para cada número de item que aparece em `min_agree` orçamentos ou mais, verifica se as
    descrições concordam entre si (por similaridade). Se sim, escolhe a descrição mais
    "consensual" (a que mais bate com as outras do grupo) e a retorna como referência fixa —
    exatamente como se fosse uma lista mestra fornecida manualmente.

    Números sem consenso suficiente não entram na tabela mestre (o casamento continua
    seguindo a lógica normal, sem trava).
    """
    grupos = defaultdict(list)
    for extraction in all_extractions:
        for item in extraction.get("itens", []):
            numero = item.get("numero_item")
            descricao = item.get("descricao")
            if numero and descricao:
                grupos[str(numero).strip()].append(descricao)

    master_items = []
    for numero, descricoes in grupos.items():
        if len(descricoes) < min_agree:
            continue

        melhor_descricao = None
        melhor_contagem = 0
        for candidata in descricoes:
            contagem = sum(
                1 for outra in descricoes
                if token_sort_ratio(normalize_desc(candidata), normalize_desc(outra)) >= consensus_threshold
            )
            e_melhor = contagem > melhor_contagem or (
                contagem == melhor_contagem
                and melhor_descricao is not None
                and len(candidata) > len(melhor_descricao)
            )
            if e_melhor:
                melhor_contagem = contagem
                melhor_descricao = candidata

        if melhor_contagem >= min_agree:
            master_items.append({"numero_item": numero, "descricao": melhor_descricao})

    return master_items


# Equivalências comuns de abreviação de unidade (normaliza antes de comparar)
def normalize_uf(uf: str) -> str:
    """Normaliza abreviações de unidade para comparação."""
    return normalizar_unidade(uf) or ""


def detectar_outliers_preco(matrix, rows) -> list[dict]:
    """Sinaliza preços fora da curva por IQR em cada linha com 3+ preços válidos."""
    alertas = []

    for row in rows:
        key = row.get("_key")
        if key is None:
            key = ("num", str(row["numero_item"]).strip()) if row.get("numero_item") else ("desc", normalize_desc(row.get("descricao", "")))
        dados = matrix.get(key, {})
        precos = []
        for empresa, info in dados.items():
            preco = info.get("preco_unitario")
            if preco is not None:
                precos.append((empresa, float(preco)))

        if len(precos) < 3:
            continue

        vals = sorted(v for _, v in precos)
        q1 = median(vals[: len(vals) // 2])
        metade_superior = vals[(len(vals) + 1) // 2 :]
        q3 = median(metade_superior) if metade_superior else vals[-1]
        iqr = q3 - q1
        lim_inf = q1 - 1.5 * iqr
        lim_sup = q3 + 1.5 * iqr

        for empresa, valor in precos:
            if valor < lim_inf or valor > lim_sup:
                alertas.append({
                    "tipo": "preço fora da curva",
                    "numero_item": row.get("numero_item"),
                    "descricao_nova": f"{row.get('descricao', '')} [{empresa}: R$ {valor:.2f}]",
                    "casou_com": f"Faixa esperada: R$ {lim_inf:.2f} a R$ {lim_sup:.2f}",
                    "score": 0,
                })

    return alertas


def _qtd_diverge(qtd_a, qtd_b, fator: float = 2.0) -> bool:
    """Retorna True se as quantidades diferem por um fator maior que `fator`."""
    if qtd_a is None or qtd_b is None:
        return False
    mn, mx = min(qtd_a, qtd_b), max(qtd_a, qtd_b)
    if mn <= 0:
        return False
    return (mx / mn) > fator


def build_comparison_table(all_extractions, master_items=None, fuzzy_threshold: int = 85,
                            sanity_threshold: int = 50,
                            usar_uf: bool = True, usar_qtd: bool = False,
                            bloquear_numero_incoerente: bool = True):
    """
    all_extractions: lista de dicts {"empresa": str, "arquivo": str, "itens": [...]}
    master_items: lista opcional [{"numero_item": ..., "descricao": ...}] vinda do edital/TR
    fuzzy_threshold: pontuação mínima (0-100) para casar dois itens por descrição.
    sanity_threshold: sinaliza revisão quando número bate mas descrição é muito diferente.
    usar_uf: quando True —
        • Casamento por descrição: UF igual reduz o threshold em 8 pts (mais fácil casar);
          UF diferente e definida bloqueia o casamento (PCT ≠ KG → nunca o mesmo produto).
        • Casamento por número: UF divergente é sinalizada para revisão.
    usar_qtd: quando True, sinaliza para revisão quando quantidade difere >2× entre
        o item já registrado e o novo (sugere cotação parcial ou item incorreto).
    bloquear_numero_incoerente: quando True, não permite fundir itens só por número
        quando a descrição diverge abaixo de `sanity_threshold`.

    Retorna:
        rows: lista ordenada de linhas canônicas {"numero_item", "descricao", "unidade", "quantidade"}
        row_index: dict {chave -> índice em rows}
        matrix: dict {chave: {empresa: {"preco_unitario", "quantidade", "unidade", "confianca", "arquivo"}}}
        review: lista de casamentos que merecem revisão manual
    """
    rows = []
    row_index = {}
    matrix = {}
    review = []

    def get_or_create_row(numero_item, descricao, quantidade=None, unidade=None, tem_preco=False, fixa=False):
        if numero_item:
            key = ("num", str(numero_item).strip())
        else:
            key = None
            desc_norm = normalize_desc(descricao)
            uf_nova = normalize_uf(unidade or "")

            for existing_key in row_index:
                if existing_key[0] != "desc":
                    continue

                score = token_sort_ratio(desc_norm, existing_key[1])

                # Calcula o threshold efetivo levando em conta a UF quando ativado
                if usar_uf and uf_nova:
                    uf_existente = normalize_uf(rows[row_index[existing_key]].get("unidade") or "")
                    if uf_existente:
                        if uf_existente == uf_nova:
                            # UF bate → mais fácil confirmar o casamento
                            threshold_efetivo = max(fuzzy_threshold - 8, 60)
                        else:
                            # UF diferente → jamais é o mesmo produto (PCT ≠ KG)
                            threshold_efetivo = 101  # impossível de atingir
                    else:
                        threshold_efetivo = fuzzy_threshold
                else:
                    threshold_efetivo = fuzzy_threshold

                if score >= threshold_efetivo:
                    key = existing_key
                    tipo = "casamento por descrição"
                    if usar_uf and uf_nova:
                        uf_existente = normalize_uf(rows[row_index[existing_key]].get("unidade") or "")
                        if uf_existente and uf_existente == uf_nova:
                            tipo = "casamento por descrição + UF confirmada"
                    review.append({
                        "tipo": tipo,
                        "numero_item": None,
                        "descricao_nova": descricao,
                        "casou_com": rows[row_index[existing_key]]["descricao"],
                        "score": score,
                    })
                    break

            if key is None:
                key = ("desc", desc_norm)

        if key not in row_index:
            row_index[key] = len(rows)
            rows.append({
                "numero_item": numero_item,
                "descricao": descricao,
                "quantidade": quantidade,
                "unidade": unidade,
                "_tem_preco": tem_preco,
                "_fixa": fixa,
            })
            matrix[key] = {}
        else:
            existing = rows[row_index[key]]

            if key[0] == "num" and not fixa:
                # Sanity: descrição muito diferente apesar do mesmo número
                score_desc = None
                if descricao and existing.get("descricao"):
                    score_desc = token_sort_ratio(
                        normalize_desc(descricao), normalize_desc(existing["descricao"])
                    )
                    if score_desc < sanity_threshold:
                        review.append({
                            "tipo": "número igual, descrição muito diferente",
                            "numero_item": numero_item,
                            "descricao_nova": descricao,
                            "casou_com": existing["descricao"],
                            "score": score_desc,
                        })

                # Se descricao diverge demais, nao funde so pelo numero: cria linha separada.
                if (
                    bloquear_numero_incoerente
                    and score_desc is not None
                    and score_desc < sanity_threshold
                ):
                    chave_alt = ("num_desc", str(numero_item).strip(), normalize_desc(descricao)[:180])
                    if chave_alt not in row_index:
                        row_index[chave_alt] = len(rows)
                        rows.append({
                            "numero_item": numero_item,
                            "descricao": descricao,
                            "quantidade": quantidade,
                            "unidade": unidade,
                            "_tem_preco": tem_preco,
                            "_fixa": False,
                        })
                        matrix[chave_alt] = {}
                    review.append({
                        "tipo": "número igual, bloqueado por descrição incompatível",
                        "numero_item": numero_item,
                        "descricao_nova": descricao,
                        "casou_com": existing.get("descricao"),
                        "score": score_desc,
                    })
                    return chave_alt

                # Sanity UF: mesmo número mas unidade incompatível
                if usar_uf and unidade and existing.get("unidade"):
                    uf_nova = normalize_uf(unidade)
                    uf_existente = normalize_uf(existing["unidade"])
                    if uf_nova and uf_existente and uf_nova != uf_existente:
                        review.append({
                            "tipo": "número igual, UF diferente",
                            "numero_item": numero_item,
                            "descricao_nova": f"{descricao} [UF: {unidade}]",
                            "casou_com": f"{existing['descricao']} [UF: {existing['unidade']}]",
                            "score": 0,
                        })

                # Sanity QTD: mesmo número mas quantidade muito divergente
                if usar_qtd and _qtd_diverge(quantidade, existing.get("quantidade")):
                    review.append({
                        "tipo": "número igual, quantidade muito diferente",
                        "numero_item": numero_item,
                        "descricao_nova": f"{descricao} [QTD: {quantidade}]",
                        "casou_com": f"{existing['descricao']} [QTD: {existing.get('quantidade')}]",
                        "score": 0,
                    })

            if not existing.get("_fixa"):
                descricao_atual = existing.get("descricao") or ""
                veio_de_item_sem_preco = not existing.get("_tem_preco")
                nova_e_melhor = tem_preco and veio_de_item_sem_preco
                nova_e_mais_completa = (
                    bool(descricao) and len(descricao) > len(descricao_atual)
                    and (tem_preco or veio_de_item_sem_preco)
                )
                if nova_e_melhor or nova_e_mais_completa:
                    existing["descricao"] = descricao
                    existing["_tem_preco"] = existing.get("_tem_preco") or tem_preco
                if numero_item and not existing.get("numero_item"):
                    existing["numero_item"] = numero_item
                if existing.get("quantidade") is None and quantidade is not None:
                    existing["quantidade"] = quantidade
                if not existing.get("unidade") and unidade:
                    existing["unidade"] = unidade
        return key

    # Pré-popula com a lista mestra do edital/TR (autoridade máxima — nunca sobrescrita)
    if master_items:
        for mi in master_items:
            numero = mi.get("numero_item")
            descricao = mi.get("descricao") or ""
            get_or_create_row(numero, descricao, tem_preco=True, fixa=True)

    for extraction in all_extractions:
        empresa = extraction.get("empresa") or extraction.get("arquivo", "Empresa desconhecida")
        arquivo = extraction.get("arquivo")
        for item in extraction.get("itens", []):
            preco = item.get("preco_unitario")
            tem_preco = preco is not None
            quantidade = item.get("quantidade")
            unidade = item.get("unidade")
            key = get_or_create_row(
                item.get("numero_item"),
                item.get("descricao", ""),
                quantidade=quantidade,
                unidade=unidade,
                tem_preco=tem_preco,
            )
            matrix[key][empresa] = {
                "preco_unitario": preco,
                "quantidade": quantidade,
                "unidade": unidade,
                "confianca": "alta" if item.get("numero_item") else "média",
                "arquivo": arquivo,
            }

    # remove os campos internos de controle antes de devolver
    rows_limpas = [
        {
            "_key": k,
            "numero_item": r["numero_item"],
            "descricao": r["descricao"],
            "quantidade": r.get("quantidade"),
            "unidade": r.get("unidade"),
        }
        for k, r in zip(row_index.keys(), rows)
    ]
    review.extend(detectar_outliers_preco(matrix, rows_limpas))
    return rows_limpas, row_index, matrix, review
