"""Calculo de confianca da extracao estrutural de PDF."""
from normalize_utils import REGEX_VALOR_BR


def calcular_confianca_estrutural(resultado_extracao: dict) -> int:
    """Calcula score (0-100) para confianca na estrutura extraida."""
    score = 0
    itens = resultado_extracao.get("itens", []) or []

    if resultado_extracao.get("encontrou_tabela"):
        score += 40

    if resultado_extracao.get("encontrou_colunas"):
        score += 25

    linhas_extraidas = resultado_extracao.get("linhas_extraidas", len(itens))
    blocos_item = resultado_extracao.get("blocos_item", 0)
    if linhas_extraidas > 0:
        if blocos_item == 0:
            score += 15
        else:
            ratio = min(linhas_extraidas, blocos_item) / max(linhas_extraidas, blocos_item)
            if ratio >= 0.6:
                score += 15

    valores_ok = True
    for item in itens:
        valor = item.get("preco_unitario")
        if valor is None:
            continue
        texto_valor = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if not REGEX_VALOR_BR.search(texto_valor):
            valores_ok = False
            break
    if valores_ok:
        score += 10

    if itens:
        validos = 0
        consistentes = 0
        for item in itens:
            qtd = item.get("quantidade")
            pu = item.get("preco_unitario")
            pt = item.get("preco_total")
            if qtd is None or pu is None or pt is None:
                continue
            validos += 1
            esperado = pu * qtd
            if esperado == 0:
                continue
            if abs(pt - esperado) / max(abs(esperado), 1e-9) <= 0.05:
                consistentes += 1
        if validos == 0 or (consistentes / validos) >= 0.8:
            score += 10

    return max(0, min(100, int(score)))
