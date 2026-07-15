# Memória de correções de casamento (aprendizado incremental)

Este arquivo é a memória permanente das decisões de casamento validadas pelo
usuário. **Consulte-o ANTES de qualquer matching** (passo 5 do SKILL.md):
decisões registradas aqui têm prioridade absoluta sobre código, número e fuzzy.

## Como registrar

Uma linha por decisão, no formato de tabela abaixo. Normalize as descrições
(minúsculas, espaços simples) ao comparar, mas registre-as como aparecem nos
documentos para legibilidade.

| Decisão | Descrição A | Descrição B | Processo de origem | Data |
|---|---|---|---|---|
| <!-- CASAR ou NAO_CASAR --> | | | | |

## Decisões registradas

_(nenhuma ainda — registre a primeira correção do usuário aqui)_

## Regras derivadas de correções

Quando várias correções apontarem o mesmo padrão, generalize aqui como regra.
Exemplos do tipo de regra que nasce das correções:

- _(exemplo)_ "ARRUELA LISA" e "ARRUELA PLANA" são sinônimos neste domínio → tratar como equivalentes no fuzzy.
- _(exemplo)_ Fornecedor X usa numeração própria de catálogo no campo Item → ignorar o número dele e casar por código/descrição.
