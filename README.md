# 🤖 Depurador de Orçamentos

Sistema que transforma **e-mails e arquivos de orçamento desestruturados** (PDF, Word, Excel, `.eml` — inclusive PDF escaneado via OCR) em dois entregáveis prontos para o processo administrativo:

1. **Mapa Comparativo de Preços** (`.xlsx`) — itens nas linhas, fornecedores nas colunas, menor preço destacado, aba de revisão e aba de fontes auditáveis.
2. **Relatório Gerencial da Pesquisa de Preços** (`.docx`) — indicadores, participação de fornecedores, cobertura de itens, análise técnica e limitações, com referência à Lei nº 14.133/2021 e à IN SEGES/ME nº 65/2021.

A filosofia do sistema: **código determinístico é o esqueleto; a IA é usada apenas onde há linguagem ambígua**. A IA nunca calcula, nunca deduplica e nunca escreve a planilha — ela lê texto confuso e devolve JSON estruturado, que o código valida.

---

## 🆕 Novidades desta versão (jul/2026)

- **Correção crítica**: o modelo padrão era `deepseek/deepseek-v4-flash`, um slug **inexistente** no OpenRouter — toda chamada de IA com o padrão falhava. Novo padrão: `google/gemini-2.5-flash` (alternativas no seletor: `deepseek/deepseek-chat-v3.2` e `openai/gpt-5-mini`).
- **Structured outputs**: extração e classificação agora enviam `response_format` com JSON Schema estrito (`temperature 0`), eliminando o parsing frágil de respostas em markdown. Se o provedor não suportar, o sistema refaz a chamada sem schema automaticamente.
- **Custo real por chamada**: o payload pede `usage: {include: true}` e o OpenRouter devolve o custo exato — a tabela de preços hardcoded virou apenas estimativa de reserva.
- **Template em branco barrado por hash**: anexos de fornecedor cujo hash de conteúdo é idêntico ao do arquivo-modelo enviado pelo órgão são descartados como "não é orçamento" (lição do processo PE-90043/2026).
- **Anexo de e-mail de dúvida/declínio aproveitado**: a classificação do e-mail e o aproveitamento do anexo são decisões independentes — um e-mail de "acompanhamento de status" pode trazer uma cotação válida anexa.
- **Relatório Gerencial `.docx`**: novo módulo `report_docx.py` + botão na aba de relatório.
- **Higiene**: removidos stubs mortos ("desativado neste build"), `.gitignore` reforçado, `cleanup.bat` para apagar arquivos acidentais, cópia duplicada `orcamentos_app/`, caches e `.venv` versionado.
- **Configuração administrada** (`app_config.py`): a chave OpenRouter e os modelos ficam salvos localmente e protegidos por senha de administrador (aba Configurações → Administração). Usuários finais **não configuram nada** — só fazem upload dos arquivos e clicam em Processar.
- **Fallback de visão no OCR**: quando o Tesseract falha ou devolve texto inútil, a imagem da página é enviada diretamente ao modelo multimodal, que lê layout, tabela, carimbo e até manuscrito. Documento que antes voltava como "erro de OCR" agora produz itens (ver seção "OCR: Tesseract, alternativas e modelos de visão").
- **Casamento por código único (PI/NSN/Part Number)**: novo critério de prioridade máxima no matching. O código do material é extraído (colunas PI/NSN/P/N/Cód. Item nos parsers estruturais e campo `codigo` no schema da IA), normalizado, indexado e usado antes do número do edital e do fuzzy. A tabela mestre aceita coluna de código, e o mapa ganhou a coluna "Código (PI/NSN/PN)".

---

## 📚 Aula: como o sistema funciona, etapa por etapa

Pense no sistema como uma linha de produção com 8 estações. Cada estação recebe um material bruto, faz uma única coisa bem feita e passa o resultado adiante. Entender **o que cada estação faz e por que ela existe** é entender o sistema inteiro.

```
.tgz/.zip com .eml  ou  arquivos avulsos (PDF/Word/Excel)
        │
   [1] INGESTÃO ─── parseia e-mails, extrai anexos, calcula hashes
        │
   [2] CLASSIFICAÇÃO ─── heurística primeiro, IA só nos ambíguos
        │
   [3] FILTRO DE ANEXOS ─── dedup por hash, barra templates em branco
        │
   [4] EXTRAÇÃO EM CAMADAS ─── parser estrutural → score → IA só se preciso
        │
   [5] NORMALIZAÇÃO ─── unidades, valores BRL, CNPJ, limpeza (100% código)
        │
   [6] MATCHING ─── número exato → fuzzy por descrição → aba de revisão
        │
   [7] VALIDAÇÃO ─── outliers por IQR, divergência parser×IA
        │
   [8] SAÍDAS ─── Mapa Comparativo .xlsx + Relatório Gerencial .docx
```

### Etapa 1 — Ingestão (`email_utils.py`, `file_utils.py`)

O pacote `.tgz`/`.zip` é aberto e cada `.eml` é parseado com a biblioteca padrão `email` do Python: remetente, destinatários, assunto, data, corpo (texto plano com fallback para HTML) e anexos em bytes. Cada anexo recebe um **hash SHA do conteúdo** — não do nome do arquivo. Isso importa porque todos os fornecedores recebem o mesmo template com o mesmo nome ("Pedido de Cotação e Modelo de Proposta.xlsx"); o nome não identifica nada, o conteúdo sim.

Aqui também acontece a **pré-passagem de templates**: antes de processar qualquer e-mail, o sistema varre os anexos enviados por remetentes institucionais (domínio `.mil.br` ou contas conhecidas da seção) e registra seus hashes. Esses são os "arquivos-modelo do órgão".

*Por que assim?* Porque a ordem dos e-mails no pacote é imprevisível — o pedido original pode aparecer depois das respostas. Coletar os hashes antes garante que o filtro da etapa 3 funcione independentemente da ordem.

### Etapa 2 — Classificação (`email_classifier.py`)

Cada e-mail vira uma de 7 categorias: `pedido_orcamento`, `orcamento_recebido`, `duvida`, `declinio`, `confirmacao_leitura`, `tramite_interno`, `outro`. O fluxo é **heurística primeiro, IA depois**: regras baratas (remetente institucional? palavras como "segue proposta"? assunto "Lida:"?) resolvem a maioria dos casos em microssegundos e custo zero. Só o que a heurística não decide vai para o LLM, que responde num JSON Schema estrito com `tipo`, `confianca`, `resumo` e `numero_processo`.

*Por que assim?* Um lote típico tem dezenas de e-mails e a maioria é óbvia. Pagar IA para classificar "Lida: Pedido de Cotação" seria desperdício. A IA entra exatamente onde ela é melhor que regex: texto livre e ambíguo.

**Regra de ouro aprendida em campo:** a classificação do e-mail e o aproveitamento do anexo são decisões **independentes**. Um fornecedor escreveu "você já analisou nossa proposta?" (classificação: dúvida) mas anexou um PDF de cotação com preços válidos. O sistema classifica o e-mail como dúvida *e* manda o anexo para a extração mesmo assim.

### Etapa 3 — Filtro e deduplicação de anexos (`app.py`)

Três barreiras, nesta ordem:

1. **Score do candidato** (`_score_email_attachment_candidate`): extensão suportada? nome tem sinais positivos ("cotação", "proposta") ou negativos ("catálogo", "logo")? O melhor anexo do e-mail é selecionado.
2. **Dedup por `(hash, remetente)`**: o mesmo arquivo do mesmo remetente não é processado duas vezes. A chave inclui o remetente de propósito — dois fornecedores diferentes *podem* mandar arquivos de bytes idênticos sem ser duplicidade.
3. **Barreira do template**: se o hash do anexo está no conjunto de templates institucionais da etapa 1, é o modelo em branco devolvido sem preencher — descartado com log explícito.

*Por que assim?* Cada documento que passa desta etapa custa dinheiro (possível chamada de IA) e risco (dado errado no mapa). Filtrar barato e cedo é a maior alavanca de custo do sistema.

### Etapa 4 — Extração em camadas (`structured_extract.py`, `extract_utils.py`, `confidence.py`)

A regra: **usar sempre a ferramenta mais confiável que a estrutura do documento permitir**, da mais barata para a mais cara:

| Camada | Documento | Ferramenta | Custo |
|---|---|---|---|
| 1 | Excel/CSV | openpyxl, detecção de cabeçalho por sinônimos | zero |
| 2 | Word com tabela | python-docx, mesma detecção | zero |
| 3 | PDF com texto | pdfplumber + score de confiança | zero |
| 4 | PDF escaneado | Tesseract OCR → tenta estrutural → IA | baixo |
| 5 | Qualquer coisa fora do padrão | LLM via OpenRouter (texto) | baixo |
| 6 | OCR falhou/inútil | **Visão multimodal**: a imagem da página vai direto ao modelo | baixo |

Para PDF, o **score de confiança** (`confidence.py`) decide a rota: acima do limiar alto (padrão 85), usa só o parser; abaixo do limiar baixo (padrão 40), vai direto para a IA; **no meio, roda os dois e compara** — é a faixa onde erros silenciosos são mais prováveis, e só nela se paga o custo da dupla checagem. O score inclui uma checagem aritmética (`unitário × quantidade ≈ total`) que pega erros que a análise estrutural sozinha não pegaria, como "25,90" lido como "2590".

A chamada de IA usa JSON Schema estrito: o modelo é obrigado a devolver `{empresa, itens[{numero_item, descricao, unidade, quantidade, preco_unitario, preco_total}]}` com tipos corretos. Cada item extraído carrega metadados de auditoria: `fonte_extracao` (estrutural/ia/dupla_checagem) e `origem` (célula, linha, página).

*Por que assim?* A IA é robusta a formatos caóticos mas não é auditável nem determinística; o parser é o contrário. Usar o parser onde a estrutura existe e a IA onde não existe dá o melhor dos dois. E a auditoria importa: numa pesquisa de preços oficial, você precisa responder "de onde veio esse valor?".

**Convenção importante:** item listado na planilha do fornecedor *sem preço preenchido* = "item não orçado por este fornecedor". Não é erro de extração — fornecedores recebem planilhas de centenas de itens e cotam poucos.

### Etapa 5 — Normalização (`normalize_utils.py`)

100% código, zero IA. Unidades são mapeadas para formas canônicas por dicionário de sinônimos (UN/UND/UNID/PEÇA → UN); valores em formato brasileiro ("R$ 1.234,56") viram float por regex; CNPJ e telefone são extraídos do texto bruto por padrão — servindo de **segunda fonte independente** para validar o que a IA extraiu.

*Por que assim?* Conversões mecânicas em LLM são custo e risco à toa. E ter duas fontes para o mesmo dado (regex no texto bruto × campo da IA) permite sinalizar divergências em vez de confiar cegamente.

### Etapa 6 — Matching de itens (`match_utils.py`, `text_similarity.py`)

O problema central do mapa: o fornecedor A escreve "PARAFUSO SEXT. M8x40 INOX" e o B escreve "Parafuso sextavado M8 40mm aço inox" — mesmo item. A solução em cascata, do critério mais forte para o mais fraco:

1. **Código único do material bate** (PI, NSN, Part Number, Nº de Estoque, Cód. do Item) → casamento direto, confiança máxima. O código identifica o material de forma inequívoca — é o critério que a metodologia naval prioriza. A comparação é normalizada (maiúsculas, sem traços/pontos/espaços: `5305-01-234-5678` = `5305.01.234.5678`), códigos curtos ou placeholders ("0", "N/A") são ignorados por segurança, e código igual com descrição totalmente diferente gera alerta na aba de revisão. Detalhe importante: quando um fornecedor informa o código e outro não, o primeiro orçamento com código "ensina" o índice — o item casado por número/descrição herda o código e os próximos fornecedores casam direto por ele.
2. **Número do item bate** (do edital/TR) → casamento direto, alta confiança. Com uma trava: se o número bate mas a descrição diverge demais, o casamento é bloqueado e sinalizado (fornecedores erram número).
3. **Fuzzy matching por descrição** (token sort ratio, limiar padrão 85) → casamento provável, vai para o mapa *e* para a aba "Revisar Casamentos".
4. **UF como bloqueio**: PCT ≠ KG — unidades diferentes e definidas nunca são o mesmo item, por mais parecida que a descrição seja.

Sem tabela mestre, o sistema pode construí-la **por consenso**: se 3+ fornecedores concordam no número e descrição de um item, isso vira referência fixa.

*Por que assim?* Nenhum critério isolado é confiável. Número erra, descrição varia, ordem no documento desalinha. A cascata usa o critério mais forte disponível e degrada com transparência — o que casou por similaridade fica marcado para conferência humana.

### Etapa 7 — Validação estatística (`match_utils.py`)

Para cada item com 3+ preços, o sistema calcula quartis e sinaliza preços fora de `[Q1 − 1.5×IQR, Q3 + 1.5×IQR]` como outliers. Tipicamente indicam vírgula decimal errada ou item cotado errado, não diferença legítima de mercado. Divergências `unitário × qtde ≠ total` também são sinalizadas — **nunca corrigidas silenciosamente**, porque podem ser erro do próprio fornecedor e isso precisa aparecer no relatório.

### Etapa 8 — Saídas (`export_utils.py`, `report_docx.py`)

O `.xlsx` sai com três abas: **Mapa Comparativo** (menor preço em verde, casamentos fuzzy em itálico), **Revisar Casamentos** (tudo que merece conferência humana, com score) e **Fontes** (empresa, arquivo, fonte de extração, localização — a trilha de auditoria completa). O `.docx` sai do `report_docx.py` com síntese de indicadores, classificação dos e-mails, participação por fornecedor com tempo de resposta, cobertura de itens por faixa, análise técnica e limitações metodológicas.

Cache em SQLite (`orcamentos.db`, `processos_emails.db`): arquivos já processados (mesmo hash) não gastam API de novo.

---

## 🧠 Workflow v2 — "confiança em cascata" (jul/2026)

Quatro mecanismos aproximam o sistema do comportamento de um analista com julgamento — a IA continua nunca calculando nem escrevendo a planilha, mas agora **decide, escala e confere**:

**1. Memória de correções (`learning_db.py`).** Cada casamento que você confirma ou rejeita na seção "Revisão de casamentos" (aba do mapa) vira uma regra permanente em `aprendizado.db`, aplicada *antes* do fuzzy matching nos próximos lotes com prioridade absoluta sobre qualquer score. É o equivalente das "lições de campo" da skill: o sistema melhora com o uso e o mesmo erro nunca precisa ser corrigido duas vezes.

**2. IA juiz na zona cinzenta (`ai_judge.py`).** Pares de descrições com similaridade 60–84 — parecidos demais para ignorar, diferentes demais para casar no automático — são enviados em lote (até 20 pares por chamada) para o LLM decidir, com salvaguardas determinísticas: só itens sem número, empresas disjuntas, UF compatível, e na dúvida o juiz é instruído a *não* casar (falso positivo em mapa oficial é pior que falso negativo). Fusões aprovadas ficam registradas na aba de revisão com a confiança declarada.

**3. Escalonamento de modelo (`extract_utils.py`).** Se a extração com o modelo barato falha ou traz zero itens, o sistema reextrai automaticamente uma única vez com o modelo forte (`google/gemini-2.5-pro`). Resultado: qualidade de modelo caro nos ~5% de documentos difíceis, custo de modelo barato nos 95% restantes.

**4. Autoverificação pré-entrega (`sanity_check.py`).** Antes de entregar o mapa, checagens determinísticas replicam o "olhar o resultado antes de enviar": coluna de fornecedor sem nenhum preço (falha de extração silenciosa), unitário × qtde ≠ total impresso (erro do fornecedor — sinalizado, nunca corrigido), fornecedores com nomes quase idênticos (coluna duplicada), linha com preço sem descrição, extração que produziu zero itens. Tudo vai para o log e para um aviso na interface.

```
extração barata ──ruim?──► reextrai c/ modelo forte     [escalonamento]
      ▼
memória de correções ──► fuzzy ──zona 60-84?──► IA juiz [decisão]
      ▼
autoverificação ──► avisos no log ──► mapa + relatório  [conferência]
      ▼
suas correções na interface ──► aprendizado.db ─┐
      ▲_________________________________________┘       [melhora contínua]
```

---

## 👁️ OCR: Tesseract, alternativas e modelos de visão — análise completa

Documento escaneado é o pior insumo do pipeline: a informação existe, mas como pixels. Há três famílias de solução, e a escolha errada aqui custa ou dinheiro ou dados perdidos.

### As três famílias

**1. OCR clássico open-source.** Tesseract (o adotado), PaddleOCR, EasyOCR, docTR. Transformam pixels em texto localmente, de graça, offline. O Tesseract é o mais maduro para português e o mais leve de instalar; PaddleOCR e EasyOCR são melhores em fotos tortas e layouts difíceis, mas arrastam PyTorch/PaddlePaddle (gigabytes de dependência) — um custo alto de manutenção para um ganho que a família 3 entrega melhor.

**2. APIs comerciais de OCR.** Google Cloud Vision, Azure Document Intelligence, AWS Textract. Excelentes, com detecção nativa de tabelas — mas cobram por página, exigem conta em nuvem separada do OpenRouter, e mandam o documento para fora da máquina. Para este sistema, seriam um segundo fornecedor de nuvem para resolver um problema que o OpenRouter já resolve.

**3. Modelos de visão (multimodais) via OpenRouter.** O mesmo `gemini-2.5-flash` que extrai texto também **enxerga**: a imagem da página vai na chamada e o modelo devolve os itens estruturados diretamente, sem a etapa intermediária de OCR. Não é "OCR melhor" — é a eliminação do OCR: o modelo lê layout de tabela, carimbo por cima do texto, assinatura, foto de tela e manuscrito razoável, coisas em que o Tesseract simplesmente falha.

### Comparação direta

| Critério | Tesseract | PaddleOCR/EasyOCR | API comercial | Visão multimodal |
|---|---|---|---|---|
| Custo por página | zero | zero | ~US$ 0,0015–0,05 | ~US$ 0,0002–0,001* |
| Scan limpo de escritório | muito bom | muito bom | excelente | excelente |
| Foto torta / baixa qualidade | fraco | bom | muito bom | **muito bom** |
| Tabelas (estrutura preservada) | fraco (só texto) | médio | muito bom | **muito bom** |
| Carimbo/manuscrito | quase nulo | fraco | médio | **bom** |
| Funciona offline / dado não sai | ✅ | ✅ | ❌ | ❌ |
| Peso de instalação | leve | pesado (torch) | nenhum | nenhum |
| Devolve itens já estruturados | ❌ | ❌ | parcial | **✅ (JSON direto)** |

*preço de entrada de imagem em modelos flash; um orçamento de 5 páginas custa fração de centavo.

### O que o sistema adota: cascata Tesseract → visão

O melhor retorno não é escolher um — é encadear os dois pelo custo:

```
PDF escaneado / imagem
   ├─ 1º Tesseract (grátis, local, resolve o scan limpo — a maioria)
   │     └─ texto bom? → segue o pipeline normal (estrutural/IA texto)
   └─ 2º texto vazio ou inútil? → páginas renderizadas em PNG e enviadas
         ao modelo multimodal, que devolve os itens em JSON estruturado
         (fonte_processamento = "visao", rastreável na aba Fontes)
```

A lógica econômica: o Tesseract resolve de graça a maioria dos escaneados (documentos de escritório digitalizados razoavelmente); o modelo de visão, que custa centavos, só é acionado nos casos em que antes o sistema devolvia "erro de OCR" — ou seja, **o custo extra só existe onde antes havia perda de dado**. E o dado sensível só sai da máquina no caso minoritário em que sairia de qualquer forma (via texto) para a extração por IA.

### Como a skill faz

A skill usa a mesma filosofia com outros meios: `pdftotext -layout` para PDF nativo, OCR apenas quando necessário — e, quando nada disso funciona, o executor da skill (Claude) **é ele próprio um modelo de visão**: abre a imagem da página e lê diretamente. O fallback de visão implementado aqui replica exatamente esse comportamento, trocando o modelo de ponta por um multimodal barato. É mais uma peça da skill portada para o sistema.

### Recomendações finais

Manter o Tesseract como primeira linha (grátis, local, privado, suficiente no caso comum) e a visão multimodal como segunda — já é o que o sistema faz. **Não** adotar PaddleOCR/EasyOCR: o ganho deles sobre o Tesseract está justamente nos casos que a visão multimodal cobre melhor, e o custo de manutenção (dependências pesadas) recai numa equipe de uma pessoa. **Não** adotar API comercial de OCR: seria um segundo fornecedor de nuvem, pago por página, para um problema que o OpenRouter já resolve por menos. Se um dia o requisito virar "nenhum dado pode sair da máquina, nem escaneado", o caminho é um modelo de visão local (família Qwen-VL/LLaVA via Ollama) — possível, mas só vale o esforço se a restrição de privacidade for absoluta.

---

## 🤖 Mapa das IAs: quando, como e quais modelos são acionados

Pergunta frequente: "é um modelo só ou vários?". A resposta: a arquitetura tem **dois "slots" de modelo** — o **principal** (barato, padrão `google/gemini-2.5-flash`) e o **forte** (padrão `google/gemini-2.5-pro`) — mas o principal é reutilizado em **cinco papéis diferentes**. Ou seja: vários acionamentos, poucos modelos. Ambos os slots são trocáveis pelo administrador (Configurações → Administração) sem mexer em código, pois tudo passa pelo OpenRouter com a mesma interface.

### Os cinco pontos de acionamento

| # | Papel | Onde (módulo) | Gatilho — quando a IA é chamada | Slot | Volume típico |
|---|---|---|---|---|---|
| 1 | Classificar e-mail | `email_classifier.py` | Só quando a heurística (remetente, palavras-chave, assunto) **não** decide a categoria | principal | minoria dos e-mails |
| 2 | Extrair itens de texto | `extract_utils.py` | XLSX/DOCX sem cabeçalho reconhecível; PDF com confiança estrutural <40 (direto) ou 40–84 (dupla checagem parser × IA) | principal | ~30–50% dos PDFs |
| 3 | Extrair por visão | `extract_utils.py` | OCR (Tesseract) falhou ou devolveu texto inútil — a **imagem** da página vai na chamada | principal (multimodal) | raro |
| 4 | Escalonamento | `extract_utils.py` | A extração nº 2/3 veio com erro ou **zero itens** — uma única retentativa | **forte** | ~5% dos documentos |
| 5 | IA juiz do matching | `ai_judge.py` | Pares de descrições com similaridade 60–84 após o matching (lotes de até 20 pares por chamada) | principal | 1–3 chamadas por lote |

### Como cada acionamento funciona

Todos os cinco seguem o mesmo contrato técnico, que é o que torna o sistema previsível: `temperature 0` (mesma entrada → mesma saída, na medida do possível), **JSON Schema estrito** via `response_format` (o modelo é obrigado a devolver os campos certos com os tipos certos — sem parsing frágil de texto livre), fallback automático sem schema se o provedor rejeitar, e `usage: {include: true}` para o OpenRouter devolver o custo real de cada chamada, que o app acumula e exibe. A IA **nunca** faz aritmética final, nunca deduplica e nunca escreve a planilha — ela devolve candidatos estruturados que o código valida (`unitário × qtde ≈ total`, normalização de moeda/unidade, sanity checks).

### Por que `gemini-2.5-flash` como principal

Quatro razões, em ordem de peso. Primeira: é **multimodal** — o mesmo modelo cobre os papéis de texto (1, 2, 5) e o de visão (3); com um modelo texto-puro (como DeepSeek) o papel 3 exigiria um segundo modelo configurado. Segunda: suporta **structured outputs** de verdade (responseSchema), o pilar do contrato acima. Terceira: **custo** — está na faixa mais barata do OpenRouter, e os papéis 1, 2 e 5 são tarefas de "leitura e transcrição estruturada", onde modelos flash têm desempenho próximo dos grandes; um lote inteiro custa centavos. Quarta: janela de contexto grande, útil para PDFs longos em modo texto. As alternativas no seletor existem por trade-offs conscientes: `deepseek/deepseek-chat-v3.2` é ainda mais barato em texto (mas sem visão — o papel 3 deixa de funcionar com ele como principal), e `openai/gpt-5-mini` tem o structured outputs mais rigoroso do mercado (mas custo um pouco maior).

### Por que `gemini-2.5-pro` como escalonamento

O papel 4 só dispara quando o flash **já falhou** — documentos de layout caótico, tabelas destruídas pela conversão, texto entremeado de ruído. Aí o que se precisa é capacidade de raciocínio sobre estrutura ambígua, que é exatamente o que separa um modelo "pro" de um "flash". Ser da mesma família importa: mesmo comportamento com o mesmo schema e o mesmo prompt, sem surpresas de formatação. E a economia fecha: pagar o modelo caro em ~5% dos documentos custa quase nada, enquanto usá-lo em 100% multiplicaria o custo do lote por ~10 sem ganho nos 95% fáceis. Alternativas no seletor admin: `anthropic/claude-sonnet-4.5` (excelente em documentos bagunçados) e `openai/gpt-5`.

### O que NUNCA usa IA (e por quê)

Parsing de `.eml` e threads, hash e deduplicação de anexos, barreira do template em branco, parsers estruturais de Excel/Word/PDF, normalização (moeda, unidade, CNPJ), matching por código PI/NSN e por número, fuzzy matching, memória de correções, detecção de outliers (IQR), autoverificação, geração do `.xlsx` e do `.docx`. Motivo comum: são operações determinísticas onde código é mais barato, mais rápido, mais auditável e **mais correto** que qualquer modelo. A regra de desenho do sistema inteiro: IA só onde há linguagem ambígua; código em todo o resto.

```
e-mail ──heurística──✔──► classificado (0 IA)
          └──ambíguo──► [1] flash

documento ──parser estrutural──✔──► itens (0 IA)
              ├──confiança 40-84──► [2] flash (dupla checagem)
              ├──confiança <40───► [2] flash (decide sozinho)
              ├──OCR inútil─────► [3] flash-visão (imagem)
              └──veio vazio/erro─► [4] PRO (uma retentativa)

matching ──código PI/NSN──✔──► casado (0 IA)
           ├──número──✔──► casado (0 IA)
           ├──fuzzy ≥85──✔──► casado (0 IA)
           └──zona 60-84──► [5] flash-juiz (lote de pares)
```

---

## 🗺️ A rota mais eficiente até a solução completa

Estado atual e o que falta, em ordem de esforço/benefício:

| # | Etapa | Status |
|---|---|---|
| 1 | Pipeline de extração em camadas + score de confiança | ✅ pronto |
| 2 | Classificação de e-mails (heurística + IA) | ✅ pronto |
| 3 | Dedup por hash + barreira de template em branco | ✅ pronto (jul/2026) |
| 4 | Datas de rodadas anteriores via texto citado | ✅ pronto |
| 5 | Modelos atuais + structured outputs + custo real | ✅ pronto (jul/2026) |
| 6 | Relatório Gerencial `.docx` | ✅ pronto (jul/2026) |
| 7 | Workflow v2: memória de correções, IA juiz, escalonamento de modelo, autoverificação | ✅ pronto (jul/2026) |
| 8 | **Validar com lote real** (`arquivos para teste/teste-171.tgz`) e calibrar limiares 85/40/60 | 🔜 próximo passo |
| 9 | Separar núcleo da UI: mover a lógica de `app.py` (2.500+ linhas) para um pacote `core/` importado pelo Streamlit — permite testes automatizados e uso por linha de comando | 🔜 |
| 10 | Suporte a "Aba2 MODELO DE PROPOSTA" multi-linha (bloco de ~10-15 linhas por item) e PDFs longos com regex por bloco | 🔜 |

O caminho 8 → 9 → 10 é deliberado: primeiro provar que o que existe funciona com dados reais (8), depois pagar a dívida técnica que trava testes (9), e só então cobrir os formatos exóticos restantes (10). Inverter essa ordem seria polir peças de um motor que ainda não rodou inteiro.

---

## 🖥️ Onde fazer deploy: análise honesta

| Critério | Local (seu PC) | Streamlit Cloud gratuito | Docker/servidor |
|---|---|---|---|
| RAM | a do seu PC | **1 GB (limita OCR e lotes)** | configurável |
| Cache SQLite persiste | ✅ | ❌ (perdido ao reiniciar) | ✅ |
| Tesseract OCR | instala uma vez | via packages.txt, frágil | no Dockerfile |
| Dados sensíveis saem da máquina | **não** | sim (upload p/ nuvem) | depende do servidor |
| Acesso por colegas | não (sem rede) | ✅ URL pública | ✅ rede interna |
| Custo | zero | zero | servidor |
| Manutenção | mínima | reboots, "Oh no", deps quebrando | média |

**Recomendação: local, sem hesitação.** Três razões: (1) o histórico deste projeto mostra que o limite de 1 GB do Cloud foi a causa das amputações de dependência — o ambiente gratuito estava *moldando a arquitetura para pior*; (2) o cache SQLite é uma das melhores características do sistema (reprocessamento grátis) e o Cloud o apaga a cada reinício; (3) e-mails de fornecedores e preços de processo administrativo são dados que não precisam transitar por nuvem de terceiros sem necessidade.

O Streamlit continua sendo a interface — apenas rodando em `localhost`. Se um dia colegas da seção precisarem acessar, o passo natural é Docker num servidor da rede interna, não o Cloud gratuito:

```dockerfile
FROM python:3.11-slim
RUN apt-get update && apt-get install -y tesseract-ocr tesseract-ocr-por && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["python", "-m", "streamlit", "run", "app.py", "--server.address=0.0.0.0"]
```

---

## 🖼️ E o front-end? Streamlit é a melhor escolha?

Pergunta justa, resposta direta: **para este sistema, neste contexto, sim — e trocar agora seria um erro.** O critério não é "qual framework é o melhor do mercado", é "qual maximiza resultado por hora de manutenção de uma equipe de uma pessoa". O Streamlit ganha porque a interface inteira vive no mesmo arquivo Python que orquestra o pipeline: sem API separada, sem build de JavaScript, sem duas linguagens para manter. Cada melhoria no motor aparece na tela com três linhas de código.

Quando faria sentido trocar, e para quê:

| Cenário futuro | Alternativa indicada | Por quê |
|---|---|---|
| Muitos usuários simultâneos com login individual e permissões | **FastAPI + React/Vue** | Streamlit reprocessa o script a cada interação; autenticação e sessões multiusuário são improvisos nele |
| Interface mais rica (tabelas editáveis complexas, drag-and-drop) mantendo só Python | **NiceGUI** ou **Reflex** | Componentes reativos de verdade, sem abandonar Python |
| Processamentos longos que não podem morrer com o navegador | **FastAPI + fila (Celery/RQ)** + qualquer front | O job roda no servidor, o navegador só consulta o status |

O gatilho objetivo para migrar é um só: **quando o app precisar de login por usuário**. Antes disso, qualquer migração é custo sem benefício. E a arquitetura já prepara esse dia — quanto mais lógica sair do `app.py` para módulos (`match_utils`, `extract_utils`, `report_docx`…), mais barata será a troca de casca no futuro. É exatamente o item 9 do roteiro.

---

## 🔐 Configuração administrada (usuários só fazem upload)

Fluxo para o administrador (uma única vez):

1. Abra a aba **Configurações → 🔐 Administração**.
2. No primeiro acesso, **defina a senha de administrador** (mínimo 6 caracteres).
3. Desbloqueie com a senha e preencha: **chave da API OpenRouter**, **modelo principal** (extração/classificação/juiz) e **modelo forte** (escalonamento automático), além dos padrões de IA juiz e classificação.
4. Salve. Pronto: a barra lateral dos usuários passa a mostrar "✅ Configuração administrada ativa" e o item de IA some — eles só enviam arquivos e clicam em Processar.

Detalhes técnicos e limites honestos: a configuração fica em `config_app.json` ao lado do app (fora do git — já está no `.gitignore`); a senha é guardada como hash PBKDF2-SHA256 com salt (não recuperável — se esquecer, apague `config_app.json` e reconfigure); a chave da API fica em claro nesse arquivo local, então a senha protege contra alteração/leitura **pela interface** — quem tem acesso direto ao disco do servidor consegue ler o arquivo. Para o cenário-alvo (app num computador da seção, usuários acessando pelo navegador), é a proteção certa no lugar certo.

---

## ⚙️ Instalação local

```bash
# 1. Pré-requisito: Tesseract OCR (para PDFs escaneados)
#    Windows: instalador em https://github.com/UB-Mannheim/tesseract/wiki (marcar Portuguese)

# 2. Ambiente e dependências
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt

# 3. Validar o ambiente
python healthcheck.py
python run_test_ocr.py

# 4. Rodar
python -m streamlit run app.py
```

Chave da API: crie conta em https://openrouter.ai/, gere a chave em https://openrouter.ai/keys e adicione crédito (centavos de dólar processam dezenas de orçamentos). A chave é digitada na interface — nunca fica no código.

## 📋 Como usar

Com a **configuração administrada ativa** (recomendado): o usuário só (1) envia os arquivos — pacote `.tgz`/`.zip` com `.eml` e/ou orçamentos avulsos PDF/Word/Excel, (2) opcionalmente sobe a tabela mestre `.xlsx`/`.csv` com colunas `numero_item` e `descricao` — e, se houver, a coluna de **código do material** (PI/NSN/Part Number), que torna o casamento praticamente infalível e (3) clica em **Processar orçamentos**. Chave e modelos já vêm da administração.

Sem configuração administrada, aparece também o passo de colar a chave OpenRouter e escolher o modelo na barra lateral (padrão `google/gemini-2.5-flash`). Em ambos os casos: ao final, baixe o Mapa Comparativo na aba principal, revise e salve decisões de casamento na seção "Revisão de casamentos" (elas viram regras permanentes), e gere o Relatório Gerencial (`.docx` ou PDF) na aba de relatório.

## 🧪 Checklist antes de usar em produção

```bash
pip install -r requirements.txt
python healthcheck.py
python run_test_ocr.py
python -m py_compile app.py extract_utils.py structured_extract.py normalize_utils.py confidence.py match_utils.py export_utils.py db_utils.py email_utils.py email_classifier.py process_db.py report_docx.py
```

Processe o lote de teste (`arquivos para teste/teste-171.tgz`) e confira: classificação dos e-mails no log, anexos filtrados vs aprovados, aba "Revisar Casamentos" e os totais do mapa.

## 📁 Estrutura do projeto

| Arquivo | Papel |
|---|---|
| `app.py` | Interface Streamlit + orquestração do fluxo |
| `email_utils.py` | Parsing de `.eml`, anexos, datas citadas no corpo |
| `email_classifier.py` | Classificação heurística + IA dos e-mails |
| `structured_extract.py` | Roteamento por tipo e extração estrutural (Excel/Word/PDF) |
| `extract_utils.py` | OCR, extração via IA (OpenRouter), pipeline em camadas |
| `confidence.py` | Score de confiança da extração estrutural de PDF |
| `normalize_utils.py` | Unidades, valores BRL, CNPJ, limpeza — determinístico |
| `match_utils.py` | Casamento de itens, tabela mestre por consenso, outliers IQR |
| `export_utils.py` | Geração do Mapa Comparativo `.xlsx` |
| `report_docx.py` | Geração do Relatório Gerencial `.docx` |
| `learning_db.py` | Memória de correções de casamento (aprendizado incremental) |
| `ai_judge.py` | IA juiz: julgamento em lote da zona cinzenta do matching |
| `sanity_check.py` | Autoverificação do mapa antes da entrega |
| `app_config.py` | Configuração administrada: chave, modelos e senha admin |
| `process_db.py` / `db_utils.py` | SQLite: processos, e-mails, participações, cache |
| `healthcheck.py` | Validação do ambiente (Tesseract, dependências) |
| `cleanup.bat` | Limpeza de arquivos acidentais e duplicações do repositório |

---

## 📊 Parecer comparativo: skill com modelo de ponta (Claude Fable 5) × Depurador de Orçamentos

**Objeto.** Este parecer compara as duas vias disponíveis para processar uma pesquisa de preços — (A) a skill `pesquisa-precos-navais` executada por um modelo de fronteira com ambiente de execução (Claude Fable 5 em Cowork/Claude Code) e (B) este sistema, o Depurador de Orçamentos — e conclui pela forma de emprego recomendada.

**Natureza das duas vias.** A distinção fundamental não é de qualidade, é de natureza. A via A é um **analista**: a skill é um roteiro metodológico, e quem o executa é um modelo de ponta que lê o caso concreto, improvisa diante do imprevisto, escreve código sob medida para o layout que nunca viu, confere visualmente o resultado e se corrige. A via B é uma **linha de produção**: um pipeline fixo, auditável e barato, que executa a mesma metodologia com julgamento de IA em pontos cirúrgicos (classificação ambígua, extração difícil, zona cinzenta do casamento) e melhora com o uso através da memória de correções.

**Análise por critério.**

*Robustez ao imprevisto.* Vantagem clara da via A. Diante de um formato de proposta inédito — um PDF de 89 páginas com blocos multi-linha, uma planilha sem cabeçalho reconhecível — o Fable 5 inventa a estratégia de extração na hora. O Depurador, mesmo com escalonamento para modelo forte, está limitado às estratégias programadas; o imprevisto genuíno resulta em aviso de "extração vazia" na autoverificação (comportamento seguro, mas não resolutivo).

*Custo por processo.* Vantagem clara da via B. O Depurador resolve a maior parte do lote com heurísticas e parsers de custo zero e usa modelos de centavos por milhão de tokens nos pontos de linguagem; um lote típico custa **centavos de dólar**. A via A emprega um modelo de fronteira em todas as etapas — o custo por processo é ordens de magnitude maior, além de consumir a cota de uso do plano.

*Velocidade e conveniência operacional.* Vantagem da via B para o uso rotineiro: o usuário faz upload e clica em um botão, sem saber escrever prompt, sem conta na Anthropic, com cache que torna reprocessamentos instantâneos. A via A exige uma sessão de conversa conduzida por alguém que saiba o que pedir.

*Reprodutibilidade e auditoria.* Vantagem da via B. O mesmo lote produz o mesmo mapa, e cada valor carrega fonte de extração e localização (aba Fontes). Na via A, duas execuções podem divergir em detalhes — aceitável para análise, indesejável como padrão de um processo administrativo recorrente.

*Aprendizado.* Empate técnico com mecanismos distintos: a skill acumula lições de campo em texto (que o modelo relê e aplica com flexibilidade); o Depurador acumula regras estruturadas em banco (aplicadas com precisão absoluta, mas sem generalização). A skill generaliza melhor; o banco nunca esquece nem interpreta errado.

**Conclusão.** Não há vencedor absoluto — há emprego correto. **Para a operação recorrente, o melhor é o Depurador**: custo marginal próximo de zero, reprodutível, auditável, operável por qualquer pessoa da seção e cada vez melhor com as correções acumuladas. **Para os casos de exceção, o melhor é a skill com o Fable 5**: o lote com formato inédito, o processo atípico, a validação de qualidade do próprio Depurador (rodar as duas vias no mesmo lote e comparar os mapas é o melhor teste de regressão disponível) e a evolução do sistema — como este próprio ciclo de melhorias demonstra, em que o modelo de ponta atuou como engenheiro do pipeline, não como seu substituto. A configuração recomendada é, portanto, complementar: **a linha de produção para os 95% rotineiros, o analista para os 5% difíceis e para melhorar continuamente a linha.** Quem adota só a via A paga caro pelo trivial; quem adota só a via B fica cego para o excepcional. Juntas, uma cobre exatamente a fraqueza da outra.

---

*escrito por: Claude Fable 5 (Anthropic)*
