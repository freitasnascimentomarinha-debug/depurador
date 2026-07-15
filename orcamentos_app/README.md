# Mapa Comparativo de Orçamentos

App Streamlit que lê orçamentos (PDF, Word, Excel — inclusive PDF escaneado via OCR),
usa extração em camadas (determinística + IA quando necessário) e monta uma planilha
comparativa: itens em linha, empresas em coluna, menor preço destacado.

## 1. Instalação local

### 1.1. Pré-requisitos do sistema
- Python 3.10 ou superior
- **Tesseract OCR** instalado no sistema (necessário para PDFs escaneados/não clicáveis):
  - Windows: baixar o instalador em https://github.com/UB-Mannheim/tesseract/wiki e, durante a instalação, marcar o pacote de idioma **Portuguese**.
  - macOS: `brew install tesseract tesseract-lang`
  - Linux (Debian/Ubuntu): `sudo apt install tesseract-ocr tesseract-ocr-por`

### 1.2. Instalar as dependências Python
```bash
cd orcamentos_app
pip install -r requirements.txt
```

## 2. Configurar a chave da OpenRouter

1. Crie uma conta em https://openrouter.ai/
2. Gere uma chave de API em https://openrouter.ai/keys
3. Adicione algum crédito (o custo por documento processado é bem baixo, especialmente com modelos como `claude-3-haiku` ou `gpt-4o-mini`).

Configure a chave uma única vez em `secrets`:

`./.streamlit/secrets.toml` (local)

```toml
OPENROUTER_API_KEY = "sua-chave-aqui"
```

No Streamlit Community Cloud, use o painel **Settings -> Secrets** com a mesma chave `OPENROUTER_API_KEY`.

## 3. Rodar o app

```bash
streamlit run app.py
```

O navegador abre automaticamente. Na barra lateral:

1. **Arquivos** — escolha como fornecer os orçamentos:
   - **Upload de arquivos**: arraste os arquivos direto na interface (funciona tanto local quanto publicado no Streamlit Cloud).
   - **Pasta local no computador**: cole o caminho de uma pasta no seu computador (ex: uma pasta sincronizada pelo Google Drive para Desktop). Só funciona rodando o app localmente — não funciona se publicado no Streamlit Cloud, já que o servidor não tem acesso ao seu computador.
2. Escolha o modelo (a chave é carregada automaticamente via `secrets`).
3. (Opcional) Suba uma planilha com a lista mestra de itens do edital/TR, com colunas `numero_item` e `descricao` — isso melhora muito a precisão do casamento e garante o nome oficial de cada item na planilha final.
4. Ajuste a sensibilidade do casamento por descrição, se necessário.
5. Ajuste os limiares de confiança estrutural para PDF:
  - confiança alta: usa parser estrutural sem IA;
  - confiança baixa: vai direto para IA;
  - faixa intermediária: roda parser + IA e envia divergências para revisão.
6. Clique em **Processar orçamentos** e aguarde (pode levar alguns minutos para 30-60 arquivos).
7. Baixe a planilha final pelo botão que aparece ao fim.

## 4. Sobre a planilha gerada

- **Aba "Mapa Comparativo"**: itens em linha, empresas em coluna, preço unitário em cada célula.
  O menor preço de cada linha fica destacado em verde. Preços em itálico/laranja indicam que o
  item foi casado por descrição (confiança média), não por número de item.
- **Aba "Revisar Casamentos"**: lista casamentos por similaridade, divergências parser vs IA e
  alertas de preço fora da curva (IQR), com score de confiança.
- **Aba "Fontes"**: mapeia empresa, arquivo, fonte de extração e localização do dado
  (célula/linha/página), para rastreabilidade auditável.

## 5. Histórico local (cache)

O app guarda em um arquivo SQLite (`orcamentos.db`, ou o caminho que você definir) os itens já
extraídos de cada arquivo. Rodando de novo, arquivos que não mudaram são identificados e reaproveitados
automaticamente, sem gastar API de novo — só arquivos novos ou alterados são reprocessados.
- No modo **upload**, a "versão" do arquivo é um hash do conteúdo (se você subir o mesmo arquivo
  de novo sem alterações, ele é reconhecido).
- No modo **pasta local**, a "versão" é a data de modificação do arquivo no disco.
- Use o checkbox "Reprocessar tudo" para ignorar o cache, e o botão "Limpar histórico salvo" para
  zerar o banco.

## 6. Publicar no Streamlit Community Cloud

O app já foi feito para isso: a chave da OpenRouter fica em **Secrets** no Streamlit Cloud,
não é digitada na interface e não fica gravada no código. Por isso o repositório no GitHub pode
ser público sem risco de vazar credenciais. Nesse caso, use o modo **Upload de arquivos**
(o modo de pasta local não funciona num servidor remoto).

### 6.1. Subir o código para o GitHub
```bash
cd orcamentos_app
git init
git add .
git commit -m "App mapa comparativo de orçamentos"
```
Crie um repositório novo em https://github.com/new (pode ser público) e depois:
```bash
git remote add origin https://github.com/SEU_USUARIO/orcamentos-app.git
git branch -M main
git push -u origin main
```

### 6.2. Publicar
1. Acesse https://share.streamlit.io/ e faça login com sua conta do GitHub.
2. Clique em **New app**.
3. Selecione o repositório, a branch `main` e o arquivo principal `app.py`.
4. Clique em **Deploy**.

O arquivo `packages.txt` (incluído neste projeto) faz o Streamlit Cloud instalar automaticamente
o Tesseract OCR e o pacote de idioma português — não precisa configurar nada manualmente para isso.

### 6.3. Limitações do plano gratuito
- 1 GB de RAM e recursos compartilhados — processar 30-60 arquivos com OCR + chamadas de IA
  pode ser mais lento que rodando local, e em lotes muito grandes pode valer a pena dividir
  o processamento em partes (ex: 20 arquivos por vez).
- O app "dorme" após um tempo sem uso e demora alguns segundos para acordar no próximo acesso.
- O cache local (`orcamentos.db`) é temporário no Streamlit Cloud: pode ser perdido quando o
  app reinicia. Rodando localmente, ele persiste normalmente entre execuções.

## 7. Ajustes finos

- **Sensibilidade do casamento por descrição** (barra lateral): controla o quão parecidas duas
  descrições precisam ser para serem consideradas o mesmo item. Comece em 85 e ajuste conforme
  a quantidade de itens que caem na aba de revisão.
- **Modelo da OpenRouter**: `claude-3-haiku` ou `gpt-4o-mini` são mais baratos e rápidos;
  `claude-3.5-sonnet` ou `gpt-4o` tendem a ser mais precisos em textos bagunçados/OCR ruim.

## 8. Checklist de deploy resiliente

Antes de atualizar no Streamlit Cloud:

```bash
pip install -r requirements.txt
python healthcheck.py
python -m py_compile app.py extract_utils.py structured_extract.py normalize_utils.py confidence.py match_utils.py export_utils.py db_utils.py
```

No repositório principal deste app tambem existem:
- `runtime.txt` fixando Python 3.11 no Cloud.
- `.streamlit/config.toml` para mostrar detalhes de erro.
- fallback no `app.py` para exibir erro de import/dependencia na tela.
