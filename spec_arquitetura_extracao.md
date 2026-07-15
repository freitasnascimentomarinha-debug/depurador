# Especificação Técnica: Arquitetura em Camadas para Extração de Orçamentos

## Contexto e objetivo

Este documento especifica melhorias a serem implementadas num sistema Python/Streamlit já existente
que extrai itens de orçamentos (PDF, Word, Excel — inclusive PDF escaneado via OCR) enviados por
30-60 fornecedores diferentes, casa os itens entre eles e gera uma planilha comparativa de preços.

**Problema a resolver:** o sistema atual manda o texto de *todo* arquivo (inclusive Excel e Word já
estruturados) para um LLM interpretar do zero. Isso é ineficiente (gasta token à toa em documentos
que já vêm estruturados), pouco auditável (não dá pra apontar de onde exatamente veio um valor) e
sujeito a erros de interpretação evitáveis.

**Objetivo:** reestruturar o pipeline de extração em camadas, usando ferramentas determinísticas
sempre que a estrutura do documento permitir, reservando o LLM para os casos em que a extração
estrutural falha ou tem baixa confiança — sem nunca sacrificar a robustez a formatos heterogêneos
que o LLM oferece hoje.

Módulos existentes do projeto que serão modificados: `extract_utils.py`, `match_utils.py`,
`export_utils.py`, `app.py`. Novo módulo a ser criado: `structured_extract.py` e `confidence.py`.

---

## Arquitetura proposta (visão geral)

```
Arquivo
  │
  ├── XLSX/XLS  ──────────────────► Extração determinística direta (pandas/openpyxl)
  │                                  Nunca passa pelo LLM.
  │
  ├── DOCX com tabela ────────────► Extração determinística direta (python-docx)
  │                                  Nunca passa pelo LLM.
  │
  ├── PDF com texto selecionável ─► 1. Tenta extração estrutural (pdfplumber)
  │                                 2. Calcula score de confiança
  │                                 3. Roteia conforme confiança (ver seção 4)
  │
  └── PDF escaneado / imagem ─────► 1. OCR (Tesseract)
                                     2. Tenta extração estrutural sobre o texto OCR
                                     3. Calcula score de confiança (penalizado por vir de OCR)
                                     4. Roteia conforme confiança (ver seção 4)

  Em todos os casos → Normalização determinística → Matching → Validação estatística → Planilha final
```

A ideia central: **o LLM deixa de ser a primeira opção e passa a ser uma camada de exceção e de
arbitragem**, não o interpretador universal do documento.

---

## 1. Detecção de tipo de arquivo e roteamento inicial

Criar `detectar_tipo_e_rotear(path: str) -> str` em `structured_extract.py`, retornando um dos
rótulos: `"xlsx"`, `"docx"`, `"pdf_texto"`, `"pdf_escaneado"`, `"imagem"`.

Regra para diferenciar `pdf_texto` de `pdf_escaneado`: já existe no código atual (checagem de
`page.get_text()` vazio no `fitz`) — reaproveitar essa lógica, só que decidindo o roteamento antes
de extrair, não durante.

---

## 2. Extração determinística para XLSX/DOCX (bypass total do LLM)

### 2.1. Excel

Implementar `extrair_xlsx_estruturado(path) -> list[dict] | None`:
- Ler com `pandas.read_excel(path, sheet_name=None)` (todas as abas)
- Para cada aba, tentar identificar a linha de cabeçalho procurando por palavras-chave em
  qualquer célula da linha: `{"item", "descrição", "descricao", "qtd", "quantidade", "valor",
  "preço", "preco", "unitário", "unitario"}` (case-insensitive, sem acento)
- Se uma linha de cabeçalho for encontrada com pelo menos 2 dessas palavras-chave em colunas
  diferentes, mapear as colunas por similaridade de nome (ex: "Vl. Unit." → `preco_unitario`,
  usando um dicionário de sinônimos, não fuzzy matching genérico) e extrair as linhas abaixo como
  itens estruturados diretamente
- Se nenhuma linha de cabeçalho reconhecível for encontrada, retornar `None` (sinaliza para cair
  no fallback de IA, tratando a aba como texto solto)

### 2.2. Word

Implementar `extrair_docx_estruturado(path) -> list[dict] | None`:
- Iterar `doc.tables` (já usado no código atual só como texto solto)
- Para cada tabela, aplicar a mesma lógica de reconhecimento de cabeçalho da seção 2.1
- Se a tabela tiver cabeçalho reconhecível, extrair direto célula por célula
- Parágrafos fora de tabela (texto corrido) continuam sendo tratados como candidatos a fallback de
  IA, já que orçamentos em Word às vezes descrevem itens em prosa, não em tabela

### 2.3. Resultado do bypass

Se `extrair_xlsx_estruturado`/`extrair_docx_estruturado` retornar uma lista não vazia, os itens já
vêm com um campo adicional `"fonte_extracao": "estrutural"` e `"origem": "B14"` (referência de
célula/linha, para auditabilidade — ver seção 6). Esses itens **pulam completamente** a chamada
`call_openrouter_extract`.

Se retornar `None` ou lista vazia, o arquivo cai no fluxo atual (texto solto → LLM), sem alteração
de comportamento.

---

## 3. Camada de normalização determinística

Criar `normalize_utils.py` com funções que rodam **depois** de qualquer extração (estrutural ou via
LLM) — ou seja, mesmo os itens que passarem pelo LLM se beneficiam dessa camada, e o prompt da IA
pode ser simplificado para não precisar mais fazer essas conversões sozinho.

### 3.1. Unidades
```python
SINONIMOS_UNIDADE = {
    "UN": ["UN", "UND", "UNID", "UNIDADE", "UNIDADES", "PC", "PÇ", "PECA", "PEÇA"],
    "CX": ["CX", "CAIXA", "CAIXAS"],
    "KG": ["KG", "QUILO", "QUILOGRAMA"],
    "L": ["L", "LT", "LITRO", "LITROS"],
    "M": ["M", "MT", "METRO", "METROS"],
    "KIT": ["KIT", "KITS", "CJ", "CONJUNTO"],
    # completar com os termos observados nos orçamentos reais do usuário
}

def normalizar_unidade(valor: str) -> str:
    """Mapeia uma unidade em texto livre para sua forma canônica, usando o dicionário acima.
    Retorna a forma canônica ou o valor original em maiúsculo se não houver correspondência."""
```
Isso substitui a normalização simplista (`strip().upper()`) já usada em `match_utils.py`.

### 3.2. CNPJ e razão social
```python
import re
REGEX_CNPJ = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")

def extrair_cnpj(texto: str) -> str | None:
    """Busca um padrão de CNPJ no texto bruto do documento (independente do LLM)."""

def extrair_razao_social(texto: str) -> str | None:
    """Heurística: procura por linhas próximas ao CNPJ contendo 'LTDA', 'S/A', 'S.A.', 'EIRELI',
    'ME', 'EPP' — candidatas fortes a razão social. Usado como reforço/validação do campo
    'empresa' que hoje só vem do LLM."""
```
Objetivo: ter uma segunda fonte (determinística) para o nome da empresa, permitindo comparar com o
que o LLM extraiu e sinalizar divergência (mesmo princípio da checagem de UF/quantidade já existente
em `match_utils.py`).

### 3.3. Valores monetários
```python
REGEX_VALOR_BR = re.compile(r"R?\$?\s?\d{1,3}(?:\.\d{3})*,\d{2}")

def parse_valor_brl(texto: str) -> float | None:
    """Converte '1.234,56' ou 'R$ 1.234,56' para 1234.56 (float). Determinístico, sem LLM."""
```
Usado tanto na extração estrutural (seção 2) quanto para validar/corrigir valores vindos do LLM.

### 3.4. Limpeza de texto
```python
def limpar_quebras_e_caracteres(texto: str) -> str:
    """Remove quebras de linha no meio de uma célula/descrição, normaliza espaços múltiplos,
    remove caracteres de controle. Roda antes de qualquer comparação de similaridade."""
```

---

## 4. Score de confiança para extração estrutural de PDF

Criar `confidence.py` com:

```python
def calcular_confianca_estrutural(resultado_extracao: dict) -> int:
    """
    Soma pontos (0-100) com base em sinais estruturais de que a extração via
    pdfplumber/tabela funcionou corretamente:

    +40  se uma tabela foi identificada no documento (linhas/colunas consistentes)
    +25  se as colunas esperadas foram encontradas (descrição, quantidade, valor)
    +15  se o número de linhas extraídas é compatível com o número de "blocos" de texto
         identificados de forma independente (ex: contagem de padrões tipo 'Item \\d+')
    +10  se todos os valores extraídos batem com REGEX_VALOR_BR (nenhum valor "estranho")
    +10  se a soma dos preco_total bate com preco_unitario * quantidade em pelo menos
         80% das linhas (checagem aritmética simples)

    IMPORTANTE: este score mede se a ESTRUTURA foi reconhecida corretamente, não se o
    CONTEÚDO está correto (ex: '25,90' extraído como '2590' passaria nos primeiros 4
    critérios). Por isso o critério de +10 acima (checagem aritmética) é essencial:
    é o único sinal que pega esse tipo de erro sem precisar de uma segunda extração.
    """
```

### 4.1. Roteamento por faixa de confiança

Em `extract_utils.py`, modificar o fluxo de PDF para:

```python
resultado_estrutural = tentar_extracao_estrutural_pdf(path)  # pdfplumber, tabelas
confianca = calcular_confianca_estrutural(resultado_estrutural)

if confianca >= 85:
    # Alta confiança: usa direto, não chama a IA. Economia máxima.
    itens = resultado_estrutural["itens"]
    fonte = "estrutural"

elif confianca < 40:
    # Baixa confiança: estrutura não foi reconhecida de forma confiável,
    # nem vale a pena comparar. Vai direto para a IA, como hoje.
    itens = call_openrouter_extract(texto_ocr_ou_pdf, ...)["itens"]
    fonte = "ia"

else:
    # Confiança média (a faixa de maior risco de erro silencioso):
    # roda os dois métodos e compara. Só aqui se paga o custo extra da IA.
    itens_ia = call_openrouter_extract(texto_ocr_ou_pdf, ...)["itens"]
    itens, conflitos = comparar_extracoes(resultado_estrutural["itens"], itens_ia)
    fonte = "dupla_checagem"
    # conflitos vão direto para a aba de revisão (ver seção 5)
```

Essa é a decisão arquitetural central desta especificação: **não escolher entre parser OU IA**,
e sim usar os dois de forma seletiva, pagando o custo da IA só na faixa onde o risco de erro
silencioso é maior. Isso preserva a maior parte do ganho de custo da extração estrutural (a maioria
dos PDFs bem formatados cairá na faixa ≥85) sem abrir mão da robustez da IA para os casos realmente
ambíguos.

Os limiares (85 / 40) devem ser expostos como configuráveis na barra lateral do Streamlit, com os
valores acima como padrão, para calibração posterior com dados reais.

---

## 5. Extração dupla e comparação (faixa de confiança média)

O casamento entre os itens extraídos pelo parser e pela IA **não deve depender só do número do
item ou da posição/ordem no documento** — essa condição é frágil (um item extra ou faltante
desalinha a ordem inteira) e perde casos legítimos em que os dois métodos descrevem o mesmo
produto com nível de detalhe diferente (ex: parser lê `"Mouse USB"`, IA lê `"Mouse USB Logitech
M90"` — mesmo item, palavras diferentes).

```python
def calcular_similaridade_item(item_parser: dict, item_ia: dict) -> float:
    """
    Combina múltiplos sinais para decidir se item_parser e item_ia representam o mesmo item,
    em vez de depender só do número ou da posição no documento.

    Usa fuzz.token_set_ratio() (não token_sort_ratio) para a descrição: token_set_ratio ignora
    palavras extras quando um texto é subconjunto do outro, então 'Mouse USB' vs 'Mouse USB
    Logitech M90' já pontua alto nesse critério sozinho, sem precisar de nenhum peso adicional
    para cobrir esse caso.

    score = (0.60 * similaridade_descricao_token_set)
          + (0.25 * 100 se unidade_normalizada bate, senão 0)
          + (0.15 * 100 se quantidade bate exatamente ou com diferença < 5%, senão 0)

    Número do item, quando presente nos dois, não entra na média — funciona como atalho: se
    bater, considera casamento direto (score = 100) sem precisar computar o resto.
    Ordem de aparecimento no documento só é usada como critério de desempate de último recurso,
    quando dois ou mais itens do outro método empatam no score acima — nunca como critério
    principal, por ser frágil a itens extras/faltantes.
    """

def comparar_extracoes(itens_parser: list[dict], itens_ia: list[dict]) -> tuple[list[dict], list[dict]]:
    """
    Casa cada item de itens_parser com o item de itens_ia de maior calcular_similaridade_item()
    (casamento guloso: maior score primeiro, cada item usado uma única vez).

    Para cada par casado, compara campo a campo: descricao, quantidade, preco_unitario.
    - Se o score de casamento for >= 70 e os valores numéricos batem (preco_unitario com
      diferença < 1%): usa o valor do parser (mais auditável) e marca confiança "alta".
    - Se o score de casamento for >= 70 mas os valores numéricos divergem, ou o score < 70
      (nenhum item do outro método pareceu corresponder): registra um conflito com os dois
      valores lado a lado, usa o valor do parser como principal, e marca confiança "baixa" —
      vai para a aba de revisão manual.

    Retorna (itens_finais, conflitos).
    """
```

Os conflitos devem ser adicionados à mesma lista `review` já existente em `match_utils.py`, com
`"tipo": "divergência parser vs. IA"`, reaproveitando a estrutura de revisão que já existe hoje
(coluna "Tipo" na aba "Revisar Casamentos").

---

## 6. Rastreabilidade / auditabilidade

Cada item extraído (por qualquer método) deve carregar, além dos campos já existentes
(`numero_item`, `descricao`, `unidade`, `quantidade`, `preco_unitario`), dois novos campos:

- `"fonte_extracao"`: um de `"estrutural"`, `"ia"`, `"dupla_checagem"`
- `"origem"`: referência à localização exata do dado — número da célula (XLSX), número da
  tabela/linha (DOCX), ou número da página (PDF). Para itens vindos de IA sem estrutura clara,
  usar apenas o número da página.

Na planilha final (`export_utils.py`), adicionar esses dois campos como colunas extras na aba
**"Fontes"** já existente (hoje ela só tem Empresa/Arquivo — passa a ter também Fonte de Extração
e Localização), permitindo responder com precisão, numa eventual auditoria: "de onde veio esse
valor, e como ele foi obtido".

---

## 7. Validação estatística de outliers (nova etapa, independente do resto)

Adicionar em `match_utils.py`, depois que a tabela comparativa estiver montada (uma checagem por
linha, entre os preços de diferentes empresas para o mesmo item — não depende de nenhuma das
mudanças acima, pode ser implementada isoladamente):

```python
def detectar_outliers_preco(matrix, rows) -> list[dict]:
    """
    Para cada linha (item) com 3+ preços válidos de empresas diferentes, calcula a mediana e o
    desvio interquartil (IQR). Preços fora de [Q1 - 1.5*IQR, Q3 + 1.5*IQR] são sinalizados como
    outlier — tipicamente indicam erro de digitação (vírgula decimal, dígito a mais/a menos) ou
    item cotado errado, mais do que uma diferença legítima de preço de mercado.

    Retorna uma lista de alertas no mesmo formato usado em 'review', com
    "tipo": "preço fora da curva", para reaproveitar a aba de revisão existente.
    """
```

Este item é o de implementação mais simples e mais barata de todas as propostas deste documento
(não precisa de LLM, não precisa reestruturar o pipeline de extração) e deve ser priorizado
primeiro na ordem de implementação.

---

## 8. Ordem de implementação recomendada

Por relação esforço/benefício, do maior ganho com menor risco para o de maior complexidade:

1. **Detecção de outliers de preço** (seção 7) — isolado, sem dependência de nada mais
2. **Camada de normalização determinística** (seção 3) — reduz erro mesmo nos itens que ainda
   passam pela IA, sem exigir mudança de roteamento
3. **Extração estrutural para XLSX/DOCX** (seção 2) — maior ganho de custo/velocidade, risco baixo
   porque esses formatos têm estrutura genuína (diferente do PDF)
4. **Score de confiança + roteamento de PDF em 3 faixas** (seção 4) — a mudança mais delicada,
   deve ser testada com um lote real de PDFs antes de virar padrão, comparando os limiares
   85/40 sugeridos contra os resultados observados
5. **Extração dupla + comparação** (seção 5) — depende da etapa 4 já estar validada

## 9. O que NÃO muda

- A extração via LLM continua existindo e sendo o caminho principal para PDF escaneado com baixa
  confiança estrutural e qualquer formato fora do padrão — não é uma proposta de eliminar a IA do
  sistema, é uma proposta de usá-la apenas onde ela agrega valor real sobre um método determinístico.
- Toda a lógica de casamento por número + descrição + UF + quantidade, os alertas de divergência,
  a tabela mestre por consenso e o cache SQLite já implementados permanecem exatamente como estão —
  as mudanças deste documento afetam a etapa de **extração**, anterior a essa lógica, não a
  substituem.

---

## 10. Decisão registrada: múltiplos parsers de PDF (Camelot) — não adotar agora

Foi avaliado adicionar o Camelot como parser estrutural alternativo ao pdfplumber (rodar os dois
e usar o resultado de maior `calcular_confianca_estrutural`). **Decisão: não incorporar nesta fase.**

Motivos:
- Camelot depende de Ghostscript (binário de sistema) e OpenCV — dependências pesadas para rodar
  no Streamlit Community Cloud (1GB de RAM compartilhada) e mais um ponto de falha de instalação
  num projeto mantido por uma pessoa só, sem time de infraestrutura dedicado.
- O modo *stream* do Camelot (o mais aplicável a PDFs sem linha de grade visível, que é o caso
  mais comum aqui) usa uma heurística de alinhamento de texto semelhante à do pdfplumber — as duas
  ferramentas tendem a falhar nos mesmos documentos, reduzindo o ganho real de redundância. O modo
  *lattice* (onde o Camelot é claramente superior) só se aplica a PDFs com linhas de grade reais.
- A comparação parser-vs-IA (seção 5) já cumpre o papel de "segunda opinião" que o Camelot
  tentaria cumprir, sem dependência de sistema adicional.

**Critério objetivo para reconsiderar esta decisão:** depois que as seções 1-8 estiverem em
produção com dados reais, medir a proporção de PDFs que caem na faixa de confiança < 40
(seção 4.1) por falha genuína de reconhecimento de tabela (não por serem PDF escaneado, que é
tratado à parte). Se essa proporção for alta **e** uma amostra desses PDFs tiver linhas de grade
visíveis (tabela real, não apenas texto alinhado), reavaliar o Camelot como parser adicional
específico para esse subconjunto — decisão orientada por dado observado, não por precaução
antecipada.
