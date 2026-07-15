"""
Autoverificação do mapa comparativo antes da entrega.

Replica o passo de "conferir antes de entregar" da metodologia: checagens
determinísticas sobre o resultado final que pegam erros silenciosos comuns.
Nenhuma checagem corrige nada — apenas sinaliza (erros do fornecedor devem
aparecer no relatório, não ser corrigidos em silêncio).
"""
from __future__ import annotations

from text_similarity import token_sort_ratio


def verificar_mapa(
    rows: list[dict],
    matrix: dict,
    empresas: list[str],
    all_extractions: list[dict] | None = None,
) -> list[str]:
    """Retorna lista de avisos (strings prontas para log/interface)."""
    avisos: list[str] = []

    # 1. Empresa presente nas extrações mas sem nenhum preço no mapa
    #    (sinal de falha de extração ou de casamento, não de recusa)
    empresas_com_preco: set[str] = set()
    for dados in matrix.values():
        for empresa, info in dados.items():
            if info.get("preco_unitario") is not None:
                empresas_com_preco.add(empresa)
    for empresa in empresas:
        if empresa not in empresas_com_preco:
            avisos.append(
                f"Coluna vazia: '{empresa}' não tem nenhum preço no mapa — "
                "verifique se a extração do arquivo desse fornecedor falhou."
            )

    # 2. unit × qtde ≠ total informado pelo fornecedor (>1% de diferença)
    #    Erro do próprio fornecedor no preenchimento: sinalizar, nunca corrigir.
    if all_extractions:
        for ext in all_extractions:
            empresa = ext.get("empresa") or ext.get("arquivo", "?")
            n_diverg = 0
            exemplo = ""
            for item in ext.get("itens", []):
                pu, qt, pt = item.get("preco_unitario"), item.get("quantidade"), item.get("preco_total")
                if pu is None or not qt or pt is None:
                    continue
                try:
                    esperado = float(pu) * float(qt)
                    informado = float(pt)
                except (TypeError, ValueError):
                    continue
                if informado > 0 and abs(esperado - informado) / informado > 0.01:
                    n_diverg += 1
                    if not exemplo:
                        exemplo = (
                            f"ex.: '{(item.get('descricao') or '')[:60]}' "
                            f"({pu} × {qt} = {esperado:.2f}, total informado {informado:.2f})"
                        )
            if n_diverg:
                avisos.append(
                    f"Aritmética divergente em {n_diverg} linha(s) de '{empresa}' "
                    f"(unitário × qtde ≠ total impresso; {exemplo}). "
                    "Possível erro de preenchimento do fornecedor — citar no relatório."
                )

    # 3. Fornecedores com nomes quase idênticos (possível duplicidade de coluna)
    for i, e1 in enumerate(empresas):
        for e2 in empresas[i + 1:]:
            score = token_sort_ratio(e1.lower().strip(), e2.lower().strip())
            if score >= 90:
                avisos.append(
                    f"Fornecedores possivelmente duplicados no mapa: '{e1}' e '{e2}' "
                    f"(similaridade {score}). Se for a mesma empresa, os preços estão divididos em duas colunas."
                )

    # 4. Linhas com preço mas sem descrição utilizável
    n_sem_desc = 0
    for row in rows:
        key = row.get("_key")
        if key is None:
            continue
        tem_preco = any(
            info.get("preco_unitario") is not None
            for info in (matrix.get(key) or {}).values()
        )
        if tem_preco and not (row.get("descricao") or "").strip():
            n_sem_desc += 1
    if n_sem_desc:
        avisos.append(
            f"{n_sem_desc} linha(s) do mapa têm preço mas descrição vazia — "
            "provável falha de extração; conferir a aba Fontes."
        )

    # 5. Volume: extração que trouxe zero itens
    if all_extractions:
        for ext in all_extractions:
            if not ext.get("itens"):
                avisos.append(
                    f"Extração vazia: '{ext.get('empresa') or ext.get('arquivo', '?')}' "
                    "não produziu nenhum item — arquivo pode ser catálogo, template ou ter falhado."
                )

    return avisos
