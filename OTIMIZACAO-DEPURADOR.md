# Depurador — Plano de otimização de desempenho

**Objetivo:** reduzir o tempo de processamento sem perder assertividade na extração.

**Premissa que guia tudo:** o gargalo do sistema **não é CPU, é espera de rede**.
Cada arquivo de orçamento fica parado num `requests.post` para o OpenRouter com
`timeout=120` (texto) ou `timeout=180` (visão), e o loop principal em `app.py`
processa **um fornecedor por vez**. Com 20 orçamentos a 40 s cada, são ~13 minutos
em que a máquina está 99% ociosa.

Todas as mudanças abaixo são **neutras em assertividade** — nenhuma delas troca
modelo, reduz contexto, pula camada de validação ou afrouxa limiar. Elas só
eliminam espera ociosa e trabalho repetido.

---

## Ranking por impacto

| # | Mudança | Ganho estimado | Risco | Esforço |
|---|---------|----------------|-------|---------|
| 1 | Paralelizar o loop de orçamentos | **5–7×** no trecho dominante | Baixo | Médio |
| 2 | Eliminar a leitura dupla do PDF | 2–4 s por PDF | Nenhum | Baixo |
| 3 | `requests.Session` com pool de conexões | 0,3–0,6 s × nº de chamadas | Nenhum | Trivial |
| 4 | OCR de páginas em paralelo | 3–5× nos PDFs escaneados | Baixo | Baixo |
| 5 | Classificação de e-mails em lote | ~15× no trecho de e-mails | Baixo | Médio |
| 6 | Cache por hash de conteúdo | Elimina reprocessos entre processos | Nenhum | Baixo |
| 7 | Higiene de repositório (`orcamentos_app/`, lixo) | Indireto (evita bugs) | Nenhum | Baixo |

---

## 1. Paralelizar o loop de orçamentos — **a mudança que importa**

### Onde está

`app.py`, linha ~2446:

```python
for i, f in enumerate(budget_candidates):
    ...
    result = extrair_orcamento_em_camadas(...)   # 30–90 s de espera de rede
    ...
    db_utils.save_extraction(conn_budget, ...)   # escrita sqlite
    process_db.vincular_orcamento_ao_processo(...)
```

O loop mistura três responsabilidades: **UI** (`st.progress`, `status.markdown`),
**rede** (a extração) e **banco** (escritas em sqlite). Isso é o que impede a
paralelização direta.

### Por que é seguro paralelizar

Cada chamada de `extrair_orcamento_em_camadas(path=...)` é **independente**: lê um
arquivo, chama a API, devolve um dict. Não compartilha estado mutável com as outras.
O resultado de processar o fornecedor B **não depende** de A. Logo, rodar em paralelo
produz exatamente os mesmos dicts — só que ao mesmo tempo.

### O que **não** pode ir para as threads

- Qualquer `st.*` (Streamlit não é thread-safe para escrita de widget).
- `conn_budget` / `conn_proc` — apesar de `check_same_thread=False`, escrita
  concorrente em sqlite gera `database is locked`.
- `add_log()` — se acumula numa lista de `session_state`, mantenha na thread principal.

### Refatoração em três fases

Substitua o loop único por: **(a) resolver cache → (b) extrair em paralelo →
(c) persistir sequencialmente**.

```python
from concurrent.futures import ThreadPoolExecutor, as_completed

MAX_WORKERS_EXTRACAO = 6   # ver nota sobre rate limit abaixo

# ---------- FASE A: cache (sequencial, rápido, mexe em sqlite) ----------
pendentes = []          # itens que precisam de IA
for f in budget_candidates:
    file_id = f["file_id"]
    cached = None if forcar_reprocessamento else db_utils.get_cached_file(conn_budget, file_id)
    cached_sem_dados_fornecedor = bool(
        cached
        and not (cached.get("cnpj") or "").strip()
        and not (cached.get("telefone") or "").strip()
    )
    if cached_sem_dados_fornecedor:
        cached = None
    if (cached
        and cached["modified_time"] == f["modified_time"]
        and cached.get("extraction_version") == db_utils.EXTRACTION_VERSION):
        # ---- exatamente o mesmo bloco de cache que já existe hoje ----
        _aplicar_resultado_cache(f, cached)     # extraia o bloco atual para esta função
        n_budget_cache += 1
    else:
        pendentes.append(f)

_set_pipeline_progress(0.38, f"{len(pendentes)} orçamento(s) para extrair com IA")

# ---------- FASE B: extração (paralela, só rede/CPU, ZERO sqlite e ZERO st.*) ----------
def _extrair_um(f: dict) -> tuple[dict, dict | None, str | None]:
    """Roda em thread. Devolve (f, result, erro). Não toca em banco nem em UI."""
    try:
        if f.get("anexos_complementares"):
            result = extrair_anexos_orcamento_em_camadas(
                f["anexos_complementares"], api_key, model, pre_filtrar,
                int(limiar_confianca_alta), int(limiar_confianca_baixa),
                lista_referencia_extracao,
            )
        else:
            result = extrair_orcamento_em_camadas(
                path=f["path"], api_key=api_key, model=model,
                pre_filtrar=pre_filtrar,
                limiar_alto=int(limiar_confianca_alta),
                limiar_baixo=int(limiar_confianca_baixa),
                lista_referencia=lista_referencia_extracao,
            )
        return f, result, None
    except Exception as exc:
        return f, None, f"{exc}\n{traceback.format_exc()}"

resultados_brutos = []
if pendentes:
    workers = max(1, min(MAX_WORKERS_EXTRACAO, len(pendentes)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futuros = {pool.submit(_extrair_um, f): f for f in pendentes}
        concluidos = 0
        for fut in as_completed(futuros):
            f_ret, result, erro = fut.result()
            resultados_brutos.append((f_ret, result, erro))
            concluidos += 1
            # UI na thread principal — as_completed roda aqui, não na worker
            _set_pipeline_progress(
                0.38 + (0.50 * (concluidos / len(pendentes))),
                f"Extraídos {concluidos}/{len(pendentes)} • último: "
                f"{f_ret.get('fornecedor_nome_hint') or f_ret['name']}",
            )
            status.markdown(f"**Concluído:** `{html.escape(str(f_ret['name']))}`")

# ---------- FASE C: persistência (sequencial, thread principal) ----------
# Ordena para manter a ordem determinística do relatório, independente de
# qual thread terminou primeiro — importante para reprodutibilidade do mapa.
ordem = {id(f): i for i, f in enumerate(budget_candidates)}
resultados_brutos.sort(key=lambda t: ordem.get(id(t[0]), 0))

for f, result, erro in resultados_brutos:
    if erro:
        falhas.append(f"{f['name']}: {erro.splitlines()[0]}")
        add_log(f"ERRO ao processar {f['name']}: {erro}")
        continue
    # ---- daqui para baixo é EXATAMENTE o código que já existe hoje ----
    # add_log dos debug_events, resolução de empresa, all_extractions.append,
    # somatório de usage/custo, diagnostico_arquivos.append,
    # db_utils.save_extraction, process_db.vincular_orcamento_ao_processo, etc.
```

### Notas importantes

**Ordem determinística.** `as_completed` devolve fora de ordem. A reordenação na
Fase C garante que o mapa comparativo saia idêntico ao de hoje — sem isso, a ordem
das colunas de fornecedor no XLSX ficaria variando entre execuções, o que é ruim
num documento oficial.

**Quantos workers.** Comece com **6**. O limitador é o rate limit do OpenRouter,
não a sua máquina. Se começar a aparecer HTTP 429 nos logs, caia para 4. Deixe
configurável no painel admin (`cfg_max_workers`) para você ajustar sem mexer no código.

**Retry vs. paralelismo.** `_call_openrouter_extract_once` já faz `time.sleep(2 ** attempt)`.
Com 6 threads, um 429 em várias ao mesmo tempo faz todas dormirem juntas e
reenviarem juntas (thundering herd). Adicione jitter:

```python
import random
time.sleep((2 ** attempt) + random.uniform(0, 1.5))
```

E trate 429 explicitamente, respeitando `Retry-After` quando vier:

```python
if resp.status_code == 429:
    espera = float(resp.headers.get("Retry-After") or (2 ** attempt))
    time.sleep(espera + random.uniform(0, 1.0))
    continue
```

**Ganho real esperado:** com 20 orçamentos e 6 workers, o trecho cai de ~13 min
para ~2,5 min. Não é 6× exato porque um arquivo pesado (visão, 8 páginas, 180 s)
domina a cauda.

---

## 2. Eliminar a leitura dupla do PDF

### O problema (medido)

`extrair_orcamento_em_camadas` faz, na primeira linha:

```python
tipo = detectar_tipo_e_rotear(path)
```

E `detectar_tipo_e_rotear` (em `structured_extract.py:101`) abre o PDF e roda
`page.extract_text()` em **todas as páginas** só para decidir se é `pdf_texto` ou
`pdf_escaneado`. Logo em seguida, `_extract_text_pages_pdf(path)` abre o PDF **de
novo** e roda `page.extract_text()` em todas as páginas outra vez.

Medição no PDF de teste do próprio repositório (`Cotação ISSARTEL`, 16 páginas):

```
passo 1 (detecção): 2,40 s
passo 2 (extração): 2,22 s   ← 100% desperdício
```

Em PDFs escaneados é pior: o OCR a 250 dpi pode rodar duas vezes dependendo do caminho.

### Correção — passe único com memoização

Adicione em `extract_utils.py`:

```python
import hashlib

_PAGINAS_CACHE: dict[str, list[dict]] = {}
_PAGINAS_CACHE_LOCK = threading.Lock()   # necessário com o item 1

def _chave_arquivo(path: str) -> str:
    st_ = os.stat(path)
    return f"{os.path.abspath(path)}|{st_.st_size}|{st_.st_mtime_ns}"

def _extract_text_pages_pdf(path: str, ocr_lang: str = "por"):
    chave = _chave_arquivo(path)
    with _PAGINAS_CACHE_LOCK:
        if chave in _PAGINAS_CACHE:
            return _PAGINAS_CACHE[chave]
    # ... corpo atual da função, produzindo `pages` ...
    with _PAGINAS_CACHE_LOCK:
        _PAGINAS_CACHE[chave] = pages
    return pages
```

E reescreva a detecção para **consumir** essas páginas em vez de reler o arquivo:

```python
def detectar_tipo_por_paginas(pages: list[dict]) -> str:
    if not pages:
        return "pdf_escaneado"
    com_texto = sum(1 for p in pages if (p.get("texto") or "").strip())
    return "pdf_texto" if com_texto >= max(1, len(pages) // 2) else "pdf_escaneado"
```

No `extrair_orcamento_em_camadas`, para extensão `.pdf`, chame
`_extract_text_pages_pdf` **primeiro** e derive o tipo dali. Para os demais formatos,
`detectar_tipo_e_rotear` continua sendo só olhar a extensão (custo zero) — mantenha.

**Cuidado com memória:** o cache guarda o texto de todas as páginas. Com 50 PDFs
grandes numa sessão isso cresce. Limpe ao final do processamento:

```python
extract_utils._PAGINAS_CACHE.clear()
```

Ou troque por `functools.lru_cache(maxsize=64)` sobre uma função que receba a chave.

---

## 3. `requests.Session` com pool de conexões

Hoje cada `requests.post` abre conexão TCP + handshake TLS do zero. São **4 pontos**
no código: `extract_utils` (texto e visão), `ai_judge`, `email_classifier`.

Com retries e escalonamento, um único arquivo pode gerar 2–4 handshakes. Multiplicado
por 20 arquivos, são dezenas de handshakes de ~300–600 ms cada.

Crie um módulo `http_client.py`:

```python
"""Cliente HTTP compartilhado com pool de conexões e retry de transporte."""
from __future__ import annotations

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_TIMEOUT_PADRAO = (10, 180)   # (connect, read)

def criar_sessao(pool_size: int = 16) -> requests.Session:
    s = requests.Session()
    retry = Retry(
        total=0,                    # retry de negócio continua no código chamador
        connect=2, read=0,
        backoff_factor=0.5,
        status_forcelist=[],
        allowed_methods=frozenset(["POST"]),
    )
    adapter = HTTPAdapter(
        pool_connections=pool_size,
        pool_maxsize=pool_size,
        max_retries=retry,
    )
    s.mount("https://", adapter)
    return s

SESSAO = criar_sessao()
```

Depois troque, nos quatro pontos, `requests.post(...)` por `SESSAO.post(...)`.

`requests.Session` é thread-safe para requisições concorrentes desde que você não
mute atributos da sessão durante o uso — o que não é o caso aqui, já que os headers
vão por chamada. **Importante:** `pool_maxsize` precisa ser ≥ ao número de workers,
senão as threads ficam disputando conexões e você perde o ganho do item 1.

Aproveite e troque `timeout=120` por tupla `(10, 120)`: hoje, se o OpenRouter demorar
a aceitar a conexão, você espera 120 s antes de descobrir. Com `(10, 120)` você
descobre em 10 s e vai para o retry.

---

## 4. OCR de páginas em paralelo

`_extract_text_pages_pdf` roda OCR página a página, sequencialmente, a 250 dpi.
Num PDF escaneado de 10 páginas isso é facilmente 40–60 s.

`pytesseract.image_to_string` invoca o binário `tesseract` num **subprocesso** — ou
seja, o GIL é liberado durante a espera e `ThreadPoolExecutor` funciona de verdade aqui.

```python
from concurrent.futures import ThreadPoolExecutor

def _ocr_paginas_paralelo(paginas_para_ocr: list[tuple[int, object]], ocr_lang: str,
                          max_workers: int = 4) -> dict[int, tuple[str, str | None]]:
    if not paginas_para_ocr:
        return {}
    resultados = {}
    with ThreadPoolExecutor(max_workers=min(max_workers, len(paginas_para_ocr))) as pool:
        futs = {pool.submit(_ocr_image, img, ocr_lang): idx
                for idx, img in paginas_para_ocr}
        for fut in futs:
            idx = futs[fut]
            try:
                resultados[idx] = fut.result()
            except Exception as exc:
                resultados[idx] = ("", str(exc))
    return resultados
```

Fluxo revisado em `_extract_text_pages_pdf`: primeiro extraia o texto nativo de todas
as páginas e marque quais são pobres (`_texto_pobre_para_ocr`); renderize só essas
para imagem; mande o lote para `_ocr_paginas_paralelo`; reintegre pelo índice.

**Atenção quando combinado com o item 1:** se 6 arquivos rodam em paralelo e cada um
abre 4 threads de OCR, são 24 processos `tesseract` competindo. Limite globalmente:

```python
_OCR_SEMAFORO = threading.BoundedSemaphore(os.cpu_count() or 4)
```

E envolva cada chamada de `_ocr_image` com esse semáforo.

**Sobre resolução:** não baixe de 250 dpi. Tabelas de preço com fonte pequena
perdem dígito abaixo disso, e trocar `1.500` por `1.50` num mapa comparativo oficial
é exatamente o erro que não pode acontecer. A paralelização dá o ganho sem esse risco.

---

## 5. Classificação de e-mails em lote

Hoje, em `app.py:2209`, cada e-mail vira **uma chamada de API**:

```python
clf = email_classifier.classificar_email(parsed, api_key, model)
```

Num `.tgz` com 60 e-mails, e considerando que `_heuristica` resolve talvez metade,
sobram ~30 chamadas sequenciais de ~2–4 s = 1 a 2 minutos só classificando.

Você **já tem o padrão certo implementado** em `ai_judge.py`: lotes de 20 pares com
JSON Schema estrito, numa chamada. Replique.

```python
CLASSIFY_BATCH_SIZE = 15

CLASSIFY_SCHEMA = {
    "name": "classificacao_emails",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "classificacoes": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "id": {"type": "integer"},
                        "tipo": {"type": "string", "enum": sorted(TIPOS_VALIDOS)},
                        "confianca": {"type": "number"},
                        "resumo": {"type": "string"},
                        "numero_processo": {"type": ["string", "null"]},
                    },
                    "required": ["id", "tipo", "confianca", "resumo", "numero_processo"],
                    "additionalProperties": False,
                },
            }
        },
        "required": ["classificacoes"],
        "additionalProperties": False,
    },
}

def classificar_emails_em_lote(parseds: list[dict], api_key: str,
                               model: str) -> dict[int, dict]:
    """Classifica vários e-mails por chamada. Heurística primeiro; só o resto vai à IA."""
    decisoes, pendentes = {}, []
    for idx, parsed in enumerate(parseds):
        tipo_heur = _heuristica(parsed)
        if tipo_heur:
            decisoes[idx] = {"tipo": tipo_heur, "confianca": 90, "resumo": "",
                             "numero_processo": None, "uso_ia": False, "usage": {}}
        else:
            pendentes.append((idx, parsed))

    for inicio in range(0, len(pendentes), CLASSIFY_BATCH_SIZE):
        lote = pendentes[inicio:inicio + CLASSIFY_BATCH_SIZE]
        # monte o user_msg com "E-MAIL {idx}: assunto / remetente / corpo truncado"
        # chame a API com CLASSIFY_SCHEMA, reintegre por `id`
        ...
    return decisoes
```

**Cuidado com a assertividade aqui:** trunque o corpo de cada e-mail (uns 1.500
caracteres) para o lote não estourar contexto. E **mantenha os dois guarda-corpos
que já existem** no `classificar_email` — o de `duvida` sem `?` e o de `duvida` com
`R$` virando `orcamento_recebido`. Aplique-os por item depois de receber o lote.

Alternativa mais simples, se preferir não mexer no prompt: mantenha uma chamada por
e-mail, mas rode os e-mails pendentes num `ThreadPoolExecutor` de 6 workers. Ganho
menor que o lote, risco praticamente zero.

---

## 6. Cache por hash de conteúdo

`db_utils.get_cached_file` usa `file_id` + `modified_time`. Isso significa que o
**mesmo PDF do mesmo fornecedor**, chegando em outro processo ou reenviado por
e-mail, é reprocessado do zero.

Adicione uma coluna `content_sha256` e consulte por ela como fallback:

```python
def _sha256_arquivo(path: str, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for bloco in iter(lambda: fh.read(chunk), b""):
            h.update(bloco)
    return h.hexdigest()
```

Lógica: procura por `file_id` (rápido, caminho atual); se não achar, procura por
hash; se achar por hash **e** `extraction_version` bate, reaproveita os itens e só
grava o vínculo novo. Ganho: em processos com fornecedores recorrentes, boa parte
dos arquivos vira cache hit puro.

O hash de um PDF de 5 MB leva ~20 ms. É irrelevante perto de uma chamada de IA.

---

## 7. Higiene do repositório

Não é desempenho, mas é o tipo de coisa que gera bug caro depois:

- **`orcamentos_app/` é uma cópia paralela** de `app.py`, `extract_utils.py`,
  `match_utils.py`, `normalize_utils.py`, etc. Duas implementações do mesmo pipeline
  **divergem** — você corrige um bug de parsing num lado e ele continua vivo no outro.
  Decida qual é a canônica e apague ou transforme a outra em `import`.
- **`__pycache__/` está versionado**, inclusive com `.pyc` de duas versões de Python
  (3.11 e 3.12). Adicione ao `.gitignore` e rode `git rm -r --cached`.
- **Arquivos-lixo de shell commitados**: `tatus --short`, `f_`, `_app`,
  `cript..._)`, `t 120 lines) ===_`, `qlite3; conn = sqlite3.connect(...)`.
  São restos de comandos colados no terminal com o cursor no lugar errado. Remova.
- **`orcamentos.db` e `processos_emails.db` versionados** — bancos com dados de
  fornecedores dentro do repositório. Tire do git e ponha no `.gitignore`.

---

## Ordem de execução sugerida

1. **Item 3** (`requests.Session`) — 20 minutos, risco zero, ganho imediato.
2. **Item 2** (leitura dupla do PDF) — 1 hora, risco zero, ganho medido.
3. **Item 1** (paralelizar orçamentos) — meio dia, é o ganho de verdade.
4. **Item 4** (OCR paralelo) — só depois do item 1, porque precisa do semáforo global.
5. **Item 5** (e-mails em lote) — quando os `.tgz` grandes começarem a incomodar.
6. **Itens 6 e 7** — manutenção.

---

## Como provar que não perdeu assertividade

Antes de mexer em qualquer coisa, gere a **linha de base** com o `teste-171.tgz` que
já está no repositório:

```bash
python run_test_ai_fallback_teste171.py > baseline.json
```

Depois de cada mudança, rode de novo e compare **item a item**, não só o total:

```python
# compara_baseline.py
import json, sys

def chave(it):
    return (str(it.get("numero_item") or ""), (it.get("descricao") or "").strip().lower()[:60])

a = {chave(i): i for i in json.load(open(sys.argv[1]))["itens"]}
b = {chave(i): i for i in json.load(open(sys.argv[2]))["itens"]}

print("só na baseline:", len(a.keys() - b.keys()))
print("só no novo    :", len(b.keys() - a.keys()))
for k in a.keys() & b.keys():
    for campo in ("preco_unitario", "quantidade", "unidade"):
        if a[k].get(campo) != b[k].get(campo):
            print(f"DIVERGE {k} {campo}: {a[k].get(campo)} -> {b[k].get(campo)}")
```

Nenhuma das mudanças 1–4 e 6 deve produzir **uma única linha** de divergência. Se
produzir, alguma delas encostou em lógica de negócio por engano — e aí você sabe
exatamente onde olhar. A mudança 5 (lote) pode gerar divergência legítima em
classificação de e-mail, porque o contexto muda; valide essa separado.
