---
name: pesquisa-precos-navais
description: "Processa pacotes .tgz de e-mails (.eml) de fornecedores recebidos em respostas a pedidos de cotação/pesquisa de preços da Marinha do Brasil (COMRJ/COPAB/CMM), e gera automaticamente o Mapa Comparativo de Preços (.xlsx) e o Relatório Gerencial da Pesquisa de Preços (.docx). Use esta skill sempre que o usuário enviar um .tgz com e-mails de cotação de sobressalentes/materiais, mencionar 'pesquisa de preços', 'mapa comparativo', 'cotação de fornecedores', número de processo tipo 'PE-xxxxx/AAAA', ou pedir para classificar e-mails de orçamento/pedido de orçamento/dúvida/recusa. Também use para atualizar ou reprocessar um mapa/relatório já existente com novos e-mails."
---

# Pesquisa de Preços Navais — Processamento de E-mails de Fornecedores

Pipeline completo para transformar um pacote `.tgz` de e-mails `.eml` (respostas de
fornecedores a pedidos de cotação) em dois entregáveis profissionais:

1. **Mapa Comparativo de Preços** (`.xlsx`) — itens x fornecedores, coluna de código
   PI/NSN, menor preço e outliers destacados, aba por fornecedor, aba de revisão de
   casamentos.
2. **Relatório Gerencial da Pesquisa de Preços** (`.docx`) — cronologia das rodadas
   de cotação, classificação dos e-mails, tempo de resposta por fornecedor, cobertura
   de itens, gráficos de pizza, indicadores de gestão, limitações e recomendações.

Consulte `references/metodologia.md` para o detalhamento completo da metodologia
(classificação, deduplicação, extração em camadas, matching, métricas). Consulte
`references/correcoes_matching.md` ANTES do matching — é a memória de correções
já validadas pelo usuário. Este SKILL.md cobre o fluxo operacional passo a passo.

## Entradas esperadas

| Entrada | Obrigatória | Descrição |
|---|---|---|
| `.tgz` | Sim | Pacote com `.eml` de e-mails trocados com fornecedores |
| Tabela mestre `.xlsx` | Recomendada | Colunas `ITEM` e `NOMENCLATURA` (numeração oficial do processo) e, se houver, coluna de **código PI/NSN** — torna o casamento praticamente infalível. Se ausente, construa a lista de itens dinamicamente a partir dos próprios orçamentos. |

Se apenas o `.tgz` for enviado sem tabela mestre, prossiga mesmo assim e avise que a
numeração dos itens será inferida dos próprios anexos de cotação (menos confiável).

## Princípio de eficiência (leia antes de começar)

**Script determinístico primeiro, julgamento seu depois.** Rode os scripts de
`scripts/` para tudo que é mecânico (parsing, hash, extração estrutural, mapa,
métricas) e reserve sua capacidade de julgamento para os pontos ambíguos:
classificação difícil, extração de layout inédito, casamento na zona cinzenta.
Não reescreva do zero o que os scripts já fazem — é lento, caro e reintroduz
erros já resolvidos.

## Passo a passo

### 1. Extrair e inspecionar

```bash
mkdir -p work/extracted && tar -xzf pacote.tgz -C work/extracted
find work/extracted -name "*.eml" | wc -l
```

### 2. Parsear os e-mails

Rode `scripts/parse_emails.py <diretorio_com_eml> <saida.json>`. Isso extrai de cada
`.eml`: remetente, assunto, data, corpo (texto), anexos (nome, hash, tamanho) e
histórico de thread (`in-reply-to`/corpo citado). Os anexos são salvos em
`work/attachments/`.

**Pré-passagem de templates (fazer ANTES de classificar):** colete os hashes de
todos os anexos enviados por remetentes institucionais (domínio `@marinha.mil.br`
ou conta institucional, ex. `sobressalentes.comrj@gmail.com`). Esses hashes são os
"templates em branco do órgão". Qualquer anexo de fornecedor com hash idêntico é o
modelo devolvido sem preencher — não é orçamento, descarte para fins de preço.
Fazer isso antes garante que funcione independentemente da ordem dos e-mails no lote.

### 3. Classificar cada e-mail

Classifique em uma das 5 categorias, lendo o corpo e observando o remetente:

- **Pedido de orçamento** — remetente é o órgão; corpo solicita proposta comercial.
- **Orçamento** — fornecedor envia preços (anexo com valores preenchidos).
- **Dúvida** — pergunta sobre prazo/especificação/status, sem apresentar preço novo.
- **Recusa** — fornecedor declina cotar.
- **Outros** — notificações internas (Asana, etc.), catálogos soltos, spam.

**Regra de ouro:** a classificação do e-mail e o aproveitamento do anexo são
decisões **independentes**. E-mails de dúvida/recusa às vezes trazem anexo com
cotação de preços válida (caso Nexbolt, PE-90043/2026) — sempre abra os anexos
antes de descartar o e-mail como "sem dado de preço".

Recupere as datas de rodadas anteriores (1ª, 2ª cotação) a partir do texto citado
("De: ... Enviada em: ...") quando a mensagem original não estiver isolada no lote.

### 4. Extrair os itens e preços de cada anexo priorizado por Orçamento

Ordem de confiabilidade (cascata — use a camada mais barata que funcionar):

1. Excel/CSV estruturado (`openpyxl`, ver `/mnt/skills/public/xlsx/SKILL.md`)
2. DOCX com tabela
3. PDF nativo: `pdftotext -layout` (ver `/mnt/skills/public/pdf-reading/SKILL.md`);
   para PDFs longos com um item por bloco, regex — ver `scripts/extract_pdf_items_regex.py`
4. PDF escaneado: OCR (tesseract)
5. **Visão direta**: se o OCR falhar ou devolver texto inútil (foto torta, carimbo,
   manuscrito), converta a página em imagem (`pdftoppm -png -r 150`) e **leia a
   imagem diretamente com sua capacidade multimodal** — você enxerga layout de
   tabela melhor que o OCR clássico. Registre `fonte_extracao: "visao"`.

**Campos a extrair por item:** `item` (nº no edital), **`codigo` (PI/NSN/Part
Number/Nº Estoque — o identificador único do material, NÃO confundir com o nº do
item)**, `descricao`, `unidade`, `qtde`, `valor_unitario`, `valor_total`, `prazo`.
Copie o código exatamente como está (com traços/pontos).

**Regras de qualidade de dado (aprendidas em campo):**
- Hash do **conteúdo** do anexo (não o nome) define duplicidade — a chave é sempre
  `remetente + hash`; nomes idênticos entre remetentes diferentes não são duplicata.
- Anexo com o mesmo hash do template do órgão = template em branco, não cotação.
- Item sem valor preenchido = "não orçado por aquele fornecedor", não erro de extração.
- Sempre confira `valor_unitário × quantidade ≈ valor_total`; divergência é erro do
  fornecedor — **sinalize no relatório, não corrija silenciosamente**, e prefira o
  menor unitário por item a somar totais brutos.
- Planilhas "Aba1 RESUMO DA COTAÇÃO" (1 linha/item) e "Aba2 MODELO DE PROPOSTA"
  (bloco multi-linha por item — escaneie linha a linha procurando preço numérico).

Consolide tudo em `quotes.json`, lista de objetos:
```json
{"item": 44, "codigo": "123569517", "fornecedor": "NOME LTDA", "qtde": 884,
 "valor_unitario": 6.26, "valor_total": 5533.84, "prazo": "IMEDIATO",
 "fonte_extracao": "pdf_nativo", "origem_arquivo": "cotacao_12750.pdf"}
```

### 5. Matching de itens (hierarquia estrita)

Consulte PRIMEIRO `references/correcoes_matching.md` — decisões já validadas pelo
usuário têm prioridade absoluta sobre qualquer critério abaixo.

1. **Código PI/NSN/Part Number** (critério mais forte): normalize antes de comparar
   — maiúsculas, sem traços/pontos/espaços (`5305-01-234-5678` = `5305.01.234.5678`).
   Ignore códigos com menos de 4 caracteres úteis ou placeholders ("0", "N/A").
   Código igual = mesmo item, mesmo com descrições diferentes (mas se a descrição
   for RADICALMENTE diferente, sinalize na aba de revisão). **Herança de código:**
   quando um fornecedor informa o código e outro não, o item casado por número/
   descrição herda o código — use-o para os fornecedores seguintes. **Dois códigos
   definidos e diferentes NUNCA são o mesmo item.**
2. **Número do item do edital**: casamento direto; se o número bate mas a descrição
   diverge muito, bloqueie e sinalize (fornecedores erram número).
3. **Descrição (fuzzy)**: similaridade alta casa; registre na aba de revisão.
4. **Zona cinzenta (similaridade intermediária, ~60–84):** julgue você mesmo, par a
   par, com estas travas: unidades de fornecimento incompatíveis (PCT ≠ KG) nunca
   casam; o mesmo fornecedor tendo cotado os dois = itens diferentes; **na dúvida
   real, NÃO case** — falso positivo em mapa oficial é pior que falso negativo.
   Todo casamento seu de zona cinzenta vai para a aba de revisão com justificativa.

### 6. Gerar o Mapa Comparativo

```bash
python3 scripts/build_mapa_comparativo.py \
  --quotes quotes.json --tabela-mestre tabela_mestre.json \
  --out Mapa_Comparativo_<PROCESSO>.xlsx
python3 /mnt/skills/public/xlsx/scripts/recalc.py Mapa_Comparativo_<PROCESSO>.xlsx
```
Confirme `status: success` e `total_errors: 0` no recalc. Inclua a coluna
**Código (PI/NSN)** entre o Nº do item e a Nomenclatura.

### 7. Autoverificação do mapa (obrigatória, antes do relatório)

Checklist numérico — rode sobre o mapa pronto e corrija/sinalize antes de seguir:

- [ ] Nenhum fornecedor com coluna 100% vazia (se houver: a extração daquele arquivo
      falhou silenciosamente — volte ao passo 4, tente a camada seguinte da cascata)
- [ ] `unitário × qtde ≈ total` em todas as linhas (divergências listadas para o relatório)
- [ ] Sem colunas de fornecedor quase duplicadas (mesmo CNPJ/nome com grafia diferente)
- [ ] Nenhuma linha com preço e descrição vazia
- [ ] Total de itens do mapa compatível com a tabela mestre (se fornecida)
- [ ] Nenhuma extração retornou zero itens sem justificativa (template/catálogo/recusa)

### 8. Calcular métricas do relatório

Use `scripts/compute_metrics.py` (ou calcule manualmente): cobertura de itens
(0/1/2/3+ orçamentos), tempo de resposta por fornecedor, totais por fornecedor,
outliers ≥3x. Gráficos de pizza com matplotlib (`Agg`) — ver `scripts/build_charts.py`.

### 9. Gerar o Relatório Gerencial (.docx)

Use o módulo genérico `scripts/report_builder.js` (não reescreva do zero — ele já
resolve os problemas de formatação de `references/docx_table_gotchas.md`):

```js
const { buildReport } = require('./scripts/report_builder.js');
buildReport({
  outPath: 'Relatorio_Gerencial_<PROCESSO>.docx',
  headerText: 'Processo <PROCESSO> — Relatório Gerencial de Pesquisa de Preços',
  blocks: [ /* ver schema em report_builder.js e exemplo em references/exemplo_blocks.json */ ]
});
```

No relatório, inclua a limitação metodológica padrão: o pacote normalmente não traz
a lista completa de convidados (BCC) — a taxa de resposta é sobre quem respondeu.

Depois de gerar, **sempre renderize e confira visualmente**:
```bash
python /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf Relatorio_*.docx
pdftoppm -jpeg -r 100 Relatorio_*.pdf page
```
E use a ferramenta de visualização de imagem em cada `page-N.jpg`.

### 10. Entregar

Copie os dois arquivos finais para `/mnt/user-data/outputs/` e use `present_files`.
Nomeie como `Mapa_Comparativo_<PROCESSO>.xlsx` e `Relatorio_Gerencial_<PROCESSO>.docx`.

## Aprendizado incremental

Cada vez que o usuário corrigir um casamento de item, registre a decisão em
`references/correcoes_matching.md` (formato descrito lá) — ela vale para TODOS os
processamentos futuros e tem prioridade sobre qualquer critério automático.
Correções de classificação de e-mail e regras de negócio novas vão neste SKILL.md;
problemas de formatação docx vão em `references/docx_table_gotchas.md`. O objetivo:
o próximo processamento já nasce correto.
