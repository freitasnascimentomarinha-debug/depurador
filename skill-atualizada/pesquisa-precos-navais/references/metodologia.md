# Metodologia de Processamento — Pesquisa de Preços via E-mails de Fornecedores

**Escopo:** Este documento define a metodologia que orienta o processamento de um pacote `.tgz` contendo e-mails `.eml` de fornecedores (orçamentos, pedidos de cotação, dúvidas, recusas), com o objetivo de gerar automaticamente o **Mapa Comparativo de Preços** e o **Relatório Gerencial da Pesquisa de Preços**.

Este arquivo serve como referência de processo. Após seu envio, o usuário fará o upload do `.tgz` (e, opcionalmente, da **tabela mestre** em `.xlsx`) na conversa, e o processamento será executado diretamente sobre esses arquivos.

---

## 1. Entradas

| Entrada | Obrigatória | Descrição |
|---|---|---|
| `.tgz` | Sim | Pacote com múltiplos `.eml` de e-mails trocados com fornecedores |
| Tabela mestre `.xlsx` | Não | Numeração oficial dos itens do processo: nº item, **código (PI/NSN)**, nomenclatura/descrição |
| Nº do processo | Não | Se não informado, tenta-se extrair do assunto/corpo dos e-mails |

Sem tabela mestre, o sistema constrói a lista de itens dinamicamente a partir dos próprios orçamentos, agrupando por código e similaridade.

---

## 2. Etapa 1 — Ingestão e Parsing

1. Extrair todos os `.eml` do `.tgz`.
2. Para cada `.eml`, extrair: remetente, destinatário(s), assunto, data/hora, corpo (texto e HTML), anexos, e cabeçalhos de thread (`In-Reply-To`, `References`).
3. **Recuperar histórico embutido**: quando o corpo do e-mail contém a mensagem original citada (">>> De: órgão..."), extrair dela a data do pedido de cotação original — essencial para calcular tempo de resposta quando o pedido não está em `.eml` separado.
4. Identificar a empresa fornecedora por:
   - Domínio do remetente (`@nomedaempresa.com.br`)
   - Assinatura/rodapé do corpo do e-mail (CNPJ, razão social, nome comercial)
   - Em caso de conflito entre os dois, priorizar a assinatura/CNPJ (mais confiável que o domínio).
5. **Pré-passagem de templates**: antes de classificar, registrar os hashes de conteúdo de todos os anexos enviados por remetentes institucionais. Esses hashes identificam o "template em branco do órgão" e alimentam o filtro da Etapa 3 — independentemente da ordem dos e-mails no lote.

---

## 3. Etapa 2 — Classificação dos E-mails

Cada e-mail é classificado em uma destas categorias:

| Categoria | Critério típico |
|---|---|
| **Pedido de orçamento** | Remetente é o órgão; corpo solicita cotação; geralmente contém planilha vazia anexa |
| **Orçamento** | Remetente é fornecedor; contém valores, anexo com itens/preços |
| **Dúvida** | Pergunta sobre especificação, prazo, condição, sem apresentar preços |
| **Recusa** | Fornecedor declina cotar (explícita ou "não trabalhamos com este item") |
| **Outros** | Confirmação de recebimento, catálogo sem preço, spam, fora de escopo |

Classificação por regras (remetente, palavras-chave, presença de anexo com valores) com fallback para LLM/julgamento nos casos ambíguos.

**A classificação do e-mail e o aproveitamento do anexo são decisões independentes** — anexos de e-mails de dúvida/recusa devem ser abertos e, se contiverem cotação válida, aproveitados.

---

## 4. Etapa 3 — Deduplicação e filtro de templates

Um orçamento é considerado **duplicado** (e descartado, mantendo-se apenas a versão mais completa/recente) quando:
- Mesmo remetente (mesmo domínio/CNPJ) **e**
- Mesmo conteúdo de anexo (hash do conteúdo, não do nome do arquivo)

**Importante:** nomes de arquivo idênticos entre remetentes diferentes **não** configuram duplicidade — a chave de deduplicação é sempre `remetente + hash de conteúdo`, nunca o nome do arquivo.

**Filtro de template:** anexo de fornecedor cujo hash de conteúdo é idêntico ao de um anexo institucional (pré-passagem da Etapa 1) é o modelo em branco devolvido sem preenchimento — descartado para fins de preço, mesmo que o nome pareça uma proposta.

---

## 5. Etapa 4 — Extração de Itens dos Anexos

Pipeline em camadas, da extração mais confiável para a mais custosa:

1. **Excel/CSV** — parser estrutural direto (cabeçalho detectado automaticamente).
2. **DOCX** — parser estrutural de tabelas.
3. **PDF nativo (texto clicável)** — extração de texto/tabela com técnicas estruturais.
4. **PDF escaneado (imagem)** — OCR seguido de estruturação.
5. **LLM (fallback textual)** — quando a estrutura não é reconhecível (colunas fora de padrão, texto corrido, planilha sem cabeçalho claro).
6. **Visão direta (fallback final)** — quando o OCR falha ou devolve texto inútil (foto torta, carimbo sobre o texto, manuscrito), a página é convertida em imagem e lida diretamente por capacidade multimodal. Registrar `fonte_extracao: "visao"`.

**Campos por item:** nº do item, **código único do material (PI/NSN/Part Number/Nº Estoque — não confundir com o nº do item, que é a posição no edital)**, descrição, unidade, quantidade, valor unitário, valor total, prazo.

Cada item extraído recebe metadados de auditoria:
- `fonte_extracao` (excel/docx/pdf_nativo/ocr/llm/visao)
- `origem_arquivo` e `remetente`
- `confianca` (score de confiança da extração)

**Filtro de ruído** — são identificados e descartados/isolados automaticamente:
- Catálogos de produto sem relação com os itens do processo
- Logos e imagens decorativas
- A planilha vazia do pedido de cotação original (detectada por hash — Etapa 3)
- Itens listados sem valor cotado — tratados como "item não orçado por este fornecedor", não como erro de extração

---

## 6. Etapa 5 — Matching de Itens (hierarquia estrita)

Decisões da memória de correções (`references/correcoes_matching.md`) têm prioridade absoluta sobre todos os critérios abaixo.

1. **Código único (PI/NSN/Part Number)** — critério mais forte. Normalização antes de comparar: maiúsculas, sem traços/pontos/espaços (`5305-01-234-5678` ≡ `5305.01.234.5678`). Códigos com menos de 4 caracteres úteis ou placeholders ("0", "N/A") são ignorados. Código igual = mesmo item (descrição radicalmente diferente → alerta na revisão). **Herança:** item casado por número/descrição herda o código do primeiro fornecedor que o informou; fornecedores seguintes casam direto por ele. **Dois códigos definidos e diferentes nunca são o mesmo item.**
2. **Número do item do edital** — casamento direto; número igual com descrição incompatível é bloqueado e sinalizado.
3. **Fuzzy matching por nomenclatura/descrição** (RapidFuzz) — na ausência ou divergência dos critérios anteriores.
4. **Zona cinzenta (similaridade intermediária)** — julgamento caso a caso com travas determinísticas: UF incompatível nunca casa; mesmo fornecedor cotou ambos = itens distintos; na dúvida, não casar (falso positivo é pior que falso negativo em mapa oficial). Casamentos de zona cinzenta sempre vão à aba **"Revisar Casamentos"**.

- **Sem tabela mestre:** agrupamento por código e similaridade textual, formando a lista de itens dinamicamente.
- **Aprendizado incremental:** cada correção manual do usuário é registrada em `references/correcoes_matching.md` e aplicada como prioridade máxima nos processamentos futuros.

---

## 7. Etapa 6 — Autoverificação do mapa (antes do relatório)

Checklist obrigatório sobre o mapa pronto:

- Nenhum fornecedor com coluna 100% vazia (indicaria falha silenciosa de extração — voltar à Etapa 4 e tentar a camada seguinte)
- `unitário × quantidade ≈ total` em todas as linhas (divergências → sinalizar no relatório, nunca corrigir silenciosamente; usar o menor unitário por item em agregações)
- Sem colunas de fornecedor quase duplicadas (mesma empresa com grafias diferentes)
- Nenhuma linha com preço e descrição vazia
- Contagem de itens compatível com a tabela mestre (quando fornecida)
- Toda extração com zero itens tem justificativa (template, catálogo, recusa)

---

## 8. Etapa 7 — Métricas do Processo

Calculadas para compor o relatório gerencial:

- Número do processo
- Quantidade de pedidos de cotação enviados e respectivas datas
- Tempo de resposta por fornecedor (data do pedido → data da resposta com orçamento)
- Número de fornecedores consultados vs. número que respondeu
- Número de itens orçados por cada fornecedor
- Distribuição de cobertura dos itens: quantos itens têm 0, 1, 2, e 3+ orçamentos (com listagem de quais itens em cada faixa)
- Concentração de fornecedores (dependência de poucos fornecedores)
- Gráfico de pizza com a distribuição de cobertura de orçamentos
- Indicadores de gestão: taxa de resposta, tempo médio de resposta, amplitude de preços por item

---

## 9. Saídas

### A) Mapa Comparativo de Preços (planilha)
- **Aba Consolidado:** todos os itens × todos os fornecedores, com nº item, **código (PI/NSN)**, nomenclatura, quantidade e valor unitário; menor preço destacado; outliers destacados.
- **Uma aba por fornecedor:** preços organizados individualmente.
- **Aba Revisar Casamentos:** casamentos fuzzy, de zona cinzenta e alertas de sanidade.
- Cabeçalhos formatados para uso oficial.

### B) Relatório Gerencial da Pesquisa de Preços
- Todas as métricas da Seção 8, com gráfico de pizza e análise técnica/gerencial em linguagem profissional, adequada à documentação formal (referenciando IN SEGES/ME nº 65/2021 e Lei nº 14.133/2021 quando pertinente).
- Limitação metodológica padrão: sem a lista de convidados (BCC), a taxa de resposta é calculada sobre quem respondeu.

---

## 10. Casos-limite tratados

- Propostas sem numeração, código ou nomenclatura padronizada → matching por similaridade textual.
- Orçamentos parciais (só alguns itens da lista têm valor) → itens sem valor tratados como não orçados, não como erro.
- PDFs em imagem, PDFs clicáveis, planilhas, DOCX, foto/carimbo/manuscrito — todos suportados pela pipeline em camadas (a última camada é leitura visual direta).
- Anexos irrelevantes (catálogos, logos, planilha vazia/template por hash) → filtrados antes da extração de itens.
- Nome de arquivo idêntico entre empresas diferentes → não gera duplicidade indevida (chave por remetente + conteúdo).
- Histórico de e-mail embutido na resposta → usado para recuperar a data do pedido original quando não há `.eml` separado do pedido.
- Fornecedor com numeração própria de catálogo → ignorar o número dele; casar por código/descrição.

---

**Próximo passo:** envie o `.tgz` (e a tabela mestre `.xlsx`, se houver) nesta conversa para iniciar o processamento conforme esta metodologia.

---

## 11. Lições de campo (processo PE-90043/2026)

Casos reais encontrados ao processar o primeiro lote, que vale reaplicar em
processos futuros:

- **Dúvida com orçamento anexo:** um fornecedor (Nexbolt) escreveu um e-mail
  de puro acompanhamento de status ("você já analisou nossa proposta?"), mas
  trazia em anexo um PDF de cotação com preços válidos de uma proposta
  anterior. Classifique o e-mail como Dúvida pelo conteúdo textual, mas
  **sempre abra os anexos antes de descartar preço** — a classificação do
  e-mail e o aproveitamento do anexo são decisões independentes.

- **Template em branco reenviado como "resposta":** vários fornecedores que
  recusaram ou apenas comentaram devolveram o mesmo arquivo
  `.xlsx`/`.pdf` que o órgão enviou como modelo, sem preencher preço. O hash
  do conteúdo desse arquivo bate com o hash do anexo do e-mail de pedido de
  cotação original — é o sinal inequívoco de "não é orçamento", mesmo que o
  nome do arquivo pareça uma proposta.

- **Preço só preenchido em parte da planilha:** fornecedores frequentemente
  recebem uma planilha com centenas de itens e só preenchem preço em alguns
  poucos ("Aba1 RESUMO DA COTAÇÃO", coluna `VALOR UNITÁRIO`). Trate item sem
  preço preenchido como "não orçado por este fornecedor" — não como falha de
  extração.

- **Layout "Aba2 MODELO DE PROPOSTA" é multi-linha por item:** cada item
  ocupa um bloco de ~10-15 linhas (cabeçalho do item + características
  técnicas decodificadas). Extração tabular direta não funciona; escaneie
  linha a linha procurando onde a coluna de preço unitário é numérica (não
  `'*'` nem vazia).

- **PDFs longos (dezenas de páginas) com um item por bloco:** exportações de
  "Aba2 MODELO DE PROPOSTA" para PDF (ex.: 89 páginas para ~650 itens) exigem
  `pdftotext -layout` seguido de regex sobre padrões de linha
  `ITEM  Nº_ESTOQUE  NOMENCLATURA  UF  QT  PRAZO  R$ unit  R$ total`. Ver
  `scripts/extract_pdf_items_regex.py`. O `Nº_ESTOQUE` desses blocos é o
  **código PI** — capture-o no campo `codigo`, não o descarte.

- **Divergência unit×qtde ≠ total impresso:** aconteceu em pelo menos 1 linha
  de um fornecedor (erro do próprio fornecedor no preenchimento). Não
  corrija silenciosamente — sinalize no relatório e não deixe que esse valor
  distorça somas agregadas quando possível (prefira preço unitário mínimo por
  item a somar "valor total" bruto de todas as linhas).

- **Rodadas de cotação anteriores não isoladas no lote:** frequentemente só a
  solicitação da rodada mais recente aparece como e-mail próprio; as datas
  das rodadas anteriores (1ª, 2ª) só existem citadas dentro do corpo dos
  e-mails de resposta ("De: ... Enviada em: ..."). Sempre vasculhe o texto
  citado antes de concluir que uma data não está disponível.

- **Sem lista de destinatários (BCC):** o pacote de e-mails normalmente não
  contém a lista de fornecedores efetivamente convidados a cotar (envio em
  massa/BCC). O relatório deve deixar claro que a "taxa de resposta" é
  calculada sobre quem respondeu, não sobre o universo total de convidados —
  a menos que o usuário forneça essa lista separadamente.
