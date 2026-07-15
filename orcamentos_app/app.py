"""
Mapa Comparativo de Orçamentos
------------------------------
Lê orçamentos (PDF/Word/Excel, inclusive PDFs escaneados via OCR) de uma pasta
local ou de arquivos enviados por upload, extrai os itens usando um LLM via
OpenRouter, casa os itens entre empresas e gera uma planilha .xlsx comparativa.
"""
import os
import tempfile
import traceback
import csv
import time

import streamlit as st

IMPORT_ERROR = None
IMPORT_ERROR_TRACE = ""
try:
    import db_utils
    import file_utils
    from export_utils import build_excel
    from extract_utils import extrair_orcamento_em_camadas, ocr_runtime_status
    from match_utils import build_comparison_table, build_master_from_consensus
except Exception as exc:  # pragma: no cover - caminho de diagnostico em deploy
    IMPORT_ERROR = exc
    IMPORT_ERROR_TRACE = traceback.format_exc()

st.set_page_config(page_title="Depurador de Orçamentos", layout="wide")
st.title("🤖 Depurador de Orçamentos")
st.caption(
    "Extrai orçamentos de PDF/Word/Excel (inclusive PDFs escaneados, via OCR) "
    "e monta uma planilha comparativa por item."
)
if IMPORT_ERROR is None:
    st.caption(f"Motor de extração ativo: v{db_utils.EXTRACTION_VERSION}")

if IMPORT_ERROR is not None:
    st.error(
        "Falha ao iniciar o app por erro de dependência/importação. "
        "Veja os detalhes técnicos abaixo para corrigir o deploy."
    )
    with st.expander("Detalhes técnicos do erro", expanded=True):
        st.code(IMPORT_ERROR_TRACE)
    st.stop()


def _default_db_path() -> str:
    # Em Streamlit Cloud, use /tmp para evitar reaproveitar banco versionado no repo.
    if os.path.exists("/mount/src"):
        return "/tmp/orcamentos.db"
    return "orcamentos.db"


def _ler_master_items(master_file) -> list[dict]:
    nome = (master_file.name or "").lower()
    if nome.endswith(".csv"):
        conteudo = master_file.getvalue().decode("utf-8", errors="ignore")
        reader = csv.DictReader(conteudo.splitlines())
        return [
            {"numero_item": r.get("numero_item"), "descricao": r.get("descricao")}
            for r in reader
        ]

    if nome.endswith(".xls") and not nome.endswith(".xlsx"):
        try:
            import xlrd
        except Exception as exc:
            raise RuntimeError("Leitura de .xls indisponível. Instale a dependência 'xlrd'.") from exc

        wb = xlrd.open_workbook(file_contents=master_file.getvalue())
        if wb.nsheets <= 0:
            return []
        ws = wb.sheet_by_index(0)
        if ws.nrows <= 0:
            return []

        headers = [str(h).strip().lower() if h is not None else "" for h in ws.row_values(0)]
        idx_num = headers.index("numero_item") if "numero_item" in headers else None
        idx_desc = headers.index("descricao") if "descricao" in headers else None
        if idx_num is None and idx_desc is None:
            return []

        out = []
        for ridx in range(1, ws.nrows):
            row = ws.row_values(ridx)
            out.append({
                "numero_item": row[idx_num] if idx_num is not None and idx_num < len(row) else None,
                "descricao": row[idx_desc] if idx_desc is not None and idx_desc < len(row) else None,
            })
        return out

    import openpyxl

    wb = openpyxl.load_workbook(master_file, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    idx_num = headers.index("numero_item") if "numero_item" in headers else None
    idx_desc = headers.index("descricao") if "descricao" in headers else None
    if idx_num is None and idx_desc is None:
        return []
    out = []
    for row in rows[1:]:
        out.append({
            "numero_item": row[idx_num] if idx_num is not None and idx_num < len(row) else None,
            "descricao": row[idx_desc] if idx_desc is not None and idx_desc < len(row) else None,
        })
    return out


def _load_openrouter_api_key() -> str:
    for secret_name in ("OPENROUTER_API_KEY", "openrouter_api_key"):
        api_key = st.secrets.get(secret_name)
        if api_key:
            return str(api_key).strip()
    return os.getenv("OPENROUTER_API_KEY", "").strip()


api_key = _load_openrouter_api_key()
ocr_ok, ocr_err = ocr_runtime_status()


def _to_int(value) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _to_float(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0

with st.sidebar:
    st.header("1. Arquivos")
    modo_entrada = st.radio(
        "Como você vai fornecer os orçamentos?",
        ["Upload de arquivos", "Pasta local no computador"],
        help=(
            "Upload: arraste os arquivos aqui (funciona local ou publicado no Streamlit Cloud). "
            "Pasta local: informe um caminho no seu computador (só funciona rodando localmente)."
        ),
    )

    uploaded_files = None
    local_folder = None

    if modo_entrada == "Upload de arquivos":
        uploaded_files = st.file_uploader(
            "Arraste os orçamentos aqui",
            type=["pdf", "docx", "doc", "xlsx", "xls", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"],
            accept_multiple_files=True,
        )
    else:
        local_folder = st.text_input(
            "Caminho da pasta",
            placeholder=r"Ex: C:\Users\voce\Google Drive\Orcamentos ou /home/voce/orcamentos",
        )

    st.header("2. Modelo")
    if api_key:
        st.markdown(
            "<p style='font-size:0.78rem;color:#b71c1c;font-weight:700;margin:0;'>"
            "Chave de API carregada via secrets.</p>",
            unsafe_allow_html=True,
        )
    else:
        st.warning("Chave de API não encontrada em secrets.")
    model = st.selectbox(
        "Modelo",
        [
            "deepseek/deepseek-v4-flash",
            "anthropic/claude-3-haiku",
            "openai/gpt-4o-mini",
        ],
        index=0,
        help="Modelos mais baratos para este fluxo: Claude Haiku e GPT-4o-mini.",
    )

    st.header("3. Lista mestra de itens (opcional)")
    st.caption("Planilha com colunas: numero_item, descricao")
    master_file = st.file_uploader("Lista mestra", type=["xls", "xlsx", "csv"], label_visibility="collapsed")

    gerar_master_auto = st.checkbox(
        "Se eu não subir lista mestra, gerar uma automaticamente por consenso", value=True,
        help="Quando 3 ou mais orçamentos concordam no número e na descrição de um item, "
             "o sistema usa isso como referência fixa — funciona como se você tivesse "
             "subido a lista mestra, mas montada a partir dos próprios orçamentos."
    )
    min_agree = st.number_input(
        "Mínimo de orçamentos que precisam concordar", min_value=2, max_value=10, value=3,
        disabled=not gerar_master_auto,
    )
    consensus_threshold = st.slider(
        "Confiança mínima de concordância", min_value=60, max_value=100, value=80,
        disabled=not gerar_master_auto,
    )

    st.header("4. Ajustes")
    if not ocr_ok:
        st.error(
            "OCR indisponivel neste ambiente: PDFs escaneados/imagem podem falhar. "
            f"Detalhe: {ocr_err}"
        )
    limiar_confianca_alta = st.slider(
        "Limiar de confiança alta (PDF estrutural)", min_value=60, max_value=100, value=85,
        help="Acima deste valor, o PDF usa extração estrutural sem chamar IA."
    )
    limiar_confianca_baixa = st.slider(
        "Limiar de confiança baixa (PDF estrutural)", min_value=0, max_value=60, value=40,
        help="Abaixo deste valor, o PDF vai direto para IA. Entre os limiares ocorre dupla checagem.",
    )
    fuzzy_threshold = st.slider(
        "Sensibilidade do casamento por descrição", min_value=70, max_value=100, value=85,
        help="Quanto maior, mais rigoroso (menos itens casados automaticamente, mais precisão)."
    )
    pre_filtrar = st.checkbox(
        "Pré-filtrar texto antes de enviar à IA (remove ruído/boilerplate)", value=True,
        help="Remove parágrafos legais longos sem números e cabeçalhos/rodapés repetidos, "
             "reduzindo o custo de API. Desative se desconfiar que algum item está sendo perdido."
    )
    sanity_threshold = st.slider(
        "Sensibilidade do alerta 'número bate, descrição diferente'", min_value=0, max_value=80, value=50,
        help="Quando dois itens têm o mesmo número mas a descrição é muito diferente, o sistema "
             "sinaliza para revisão em vez de confiar cegamente no número. Quanto maior o valor, "
             "mais fácil disparar o alerta (mais rigoroso)."
    )
    bloquear_numero_incoerente = st.checkbox(
        "Bloquear casamento automático quando número bate mas descrição diverge",
        value=True,
        help="Quando ativado, itens com mesmo número mas descrição incompatível são separados "
             "em linhas distintas e enviados para revisão manual.",
    )
    usar_uf_casamento = st.checkbox(
        "Usar UF (unidade de fornecimento) na identificação dos itens",
        value=True,
        help="Quando ativado:\n"
             "• Casamento por descrição: UF igual facilita o match; UF diferente (ex: PCT vs KG) bloqueia — "
             "itens com unidades incompatíveis nunca são o mesmo produto.\n"
             "• Casamento por número: se UF divergir entre fornecedores, o item é sinalizado para revisão.",
    )
    usar_qtd_casamento = st.checkbox(
        "Usar quantidade para validar casamentos",
        value=False,
        help="Quando ativado, sinaliza para revisão os itens cujas quantidades diferem mais de 2× entre "
             "fornecedores — pode indicar que um fornecedor cotou item diferente ou parcial. "
             "Não bloqueia o casamento, apenas avisa.",
    )

    st.header("5. Histórico (cache local)")
    db_path = st.text_input("Arquivo do banco de dados", value=_default_db_path())
    forcar_reprocessamento = st.checkbox(
        "Reprocessar tudo (ignorar cache)", value=False,
        help="Marque se quiser forçar a extração de novo em todos os arquivos, mesmo os já processados antes."
    )

    conn_preview = db_utils.get_connection(db_path)
    removidos_antigos = db_utils.purge_old_versions(conn_preview)
    n_cached = db_utils.count_files(conn_preview)
    if removidos_antigos:
        st.caption(f"🧹 {removidos_antigos} entrada(s) antigas de cache removidas automaticamente.")
    st.caption(f"📦 {n_cached} arquivo(s) já salvos no histórico local.")
    if n_cached and st.button("🗑️ Limpar histórico salvo"):
        db_utils.clear_db(conn_preview)
        st.success("Histórico limpo.")
        st.rerun()

    processar = st.button("🚀 Processar orçamentos", type="primary", use_container_width=True)

if processar:
    try:
        if modo_entrada == "Upload de arquivos" and not uploaded_files:
            st.error("Envie pelo menos um arquivo.")
            st.stop()
        if modo_entrada == "Pasta local no computador" and not local_folder:
            st.error("Informe o caminho da pasta.")
            st.stop()
        if not api_key:
            st.error(
                "Chave da OpenRouter ausente. Configure OPENROUTER_API_KEY em "
                "secrets (arquivo .streamlit/secrets.toml ou painel Secrets do Streamlit Cloud)."
            )
            st.stop()

        with tempfile.TemporaryDirectory() as tmpdir:
            master_items = None
            if master_file:
                master_items = _ler_master_items(master_file)

            # Monta uma lista única de arquivos a processar, com {name, path, modified_time},
            # independente de terem vindo de upload ou de pasta local.
            files = []
            if modo_entrada == "Upload de arquivos":
                for uf in uploaded_files:
                    content = uf.getbuffer().tobytes()
                    local_path = os.path.join(tmpdir, uf.name)
                    with open(local_path, "wb") as fh:
                        fh.write(content)
                    files.append({
                        "name": uf.name,
                        "path": local_path,
                        "modified_time": file_utils.hash_bytes(content),
                    })
            else:
                try:
                    files = file_utils.list_local_files(local_folder)
                except NotADirectoryError as exc:
                    st.error(str(exc))
                    st.stop()

            if not files:
                st.warning("Nenhum arquivo suportado (.pdf, .docx, .xlsx, .png, .jpg, .jpeg, .tif, .tiff, .bmp, .webp) encontrado.")
                st.stop()

            st.success(f"{len(files)} arquivo(s) encontrado(s). Iniciando processamento...")

            conn = db_utils.get_connection(db_path)

            progress = st.progress(0)
            status = st.empty()
            log_expander = st.expander("📜 Logs de processamento (tempo real)", expanded=True)
            log_placeholder = log_expander.empty()
            log_lines = []

            def add_log(message: str):
                stamp = time.strftime("%H:%M:%S")
                log_lines.append(f"[{stamp}] {message}")
                log_placeholder.code("\n".join(log_lines[-250:]), language="text")

            add_log(f"Iniciando processamento de {len(files)} arquivo(s).")
            all_extractions = []
            sources = []
            review_extracao = []
            falhas = []
            diagnostico_arquivos = []
            n_do_cache = 0
            n_novos = 0
            avisos_truncamento = []
            prompt_tokens_total = 0
            completion_tokens_total = 0
            total_tokens_total = 0
            custo_total_usd = 0.0
            arquivos_com_ia = 0
            custo_estimado = False

            for i, f in enumerate(files):
            # file_id: para pasta local usa o caminho completo; para upload usa o nome do arquivo
                file_id = f["path"] if modo_entrada == "Pasta local no computador" else f["name"]
                modified_time = f["modified_time"]
                add_log(f"[{i + 1}/{len(files)}] Preparando arquivo: {f['name']}")
                cached = None if forcar_reprocessamento else db_utils.get_cached_file(conn, file_id)

                if (
                    cached
                    and cached["modified_time"] == modified_time
                    and cached.get("extraction_version") == db_utils.EXTRACTION_VERSION
                ):
                # Arquivo não mudou desde a última execução: usa o que já está salvo, sem gastar API
                    status.text(f"(cache) {f['name']} ({i + 1}/{len(files)})")
                    add_log(f"{f['name']}: usando cache local (sem chamada de IA).")
                    itens = db_utils.get_items_for_file(conn, file_id)
                    empresa = cached["empresa"]
                    all_extractions.append({"empresa": empresa, "arquivo": f['name'], "itens": itens})
                    diagnostico_arquivos.append({
                        "arquivo": f["name"],
                        "empresa": empresa,
                        "itens": len(itens),
                        "fonte": "cache",
                        "confianca": cached.get("extraction_version"),
                    })
                    for item in itens:
                        sources.append({
                            "empresa": empresa,
                            "arquivo": f['name'],
                            "fonte_extracao": item.get("fonte_extracao", "cache"),
                            "origem": item.get("origem", "desconhecida"),
                            "numero_item": item.get("numero_item"),
                            "descricao": item.get("descricao"),
                        })
                    n_do_cache += 1
                    progress.progress((i + 1) / len(files))
                    continue

                status.text(f"Processando: {f['name']} ({i + 1}/{len(files)})")
                try:
                    add_log(f"{f['name']}: extraindo itens (estrutural/IA conforme confiança).")
                    result = extrair_orcamento_em_camadas(
                        path=f["path"],
                        api_key=api_key,
                        model=model,
                        pre_filtrar=pre_filtrar,
                        limiar_alto=int(limiar_confianca_alta),
                        limiar_baixo=int(limiar_confianca_baixa),
                    )
                    if result.get("erro") and not result.get("itens"):
                        falhas.append(f"{f['name']}: {result['erro']}")
                        add_log(f"{f['name']}: falha na extração -> {result['erro']}")
                        progress.progress((i + 1) / len(files))
                        continue

                    for event in result.get("debug_events", []):
                        add_log(f"{f['name']}: {event}")
                    for event in result.get("debug_highlights", []):
                        add_log(f"{f['name']}: !!! {event}")

                    empresa = result.get("empresa") or os.path.splitext(f['name'])[0]
                    itens = result.get("itens", [])
                    if result.get("texto_truncado"):
                        avisos_truncamento.append(f['name'])
                    all_extractions.append({"empresa": empresa, "arquivo": f['name'], "itens": itens})
                    review_extracao.extend(result.get("review", []))
                    usage = result.get("usage") or {}
                    prompt_tokens_total += _to_int(usage.get("prompt_tokens"))
                    completion_tokens_total += _to_int(usage.get("completion_tokens"))
                    total_tokens_total += _to_int(usage.get("total_tokens"))
                    custo_total_usd += _to_float(usage.get("cost_usd"))
                    custo_estimado = custo_estimado or bool(usage.get("estimated"))
                    if _to_int(usage.get("total_tokens")) > 0:
                        arquivos_com_ia += 1
                    add_log(
                        f"{f['name']}: concluído | itens={len(itens)} | "
                        f"fonte={result.get('fonte_processamento', 'ia')} | "
                        f"tokens={_to_int(usage.get('total_tokens'))} | "
                        f"custo=US$ {_to_float(usage.get('cost_usd')):.6f}"
                    )
                    diagnostico_arquivos.append({
                        "arquivo": f["name"],
                        "empresa": empresa,
                        "itens": len(itens),
                        "fonte": result.get("fonte_processamento", "ia"),
                        "confianca": result.get("confianca_estrutural"),
                    })
                    for item in itens:
                        sources.append({
                            "empresa": empresa,
                            "arquivo": f['name'],
                            "fonte_extracao": item.get("fonte_extracao", result.get("fonte_processamento", "ia")),
                            "origem": item.get("origem", "desconhecida"),
                            "numero_item": item.get("numero_item"),
                            "descricao": item.get("descricao"),
                        })
                    db_utils.save_extraction(conn, file_id, f['name'], empresa, modified_time, itens)
                    n_novos += 1
                except Exception as exc:
                    falhas.append(f"{f['name']}: {exc}")
                    add_log(f"{f['name']}: exceção durante processamento -> {exc}")

                progress.progress((i + 1) / len(files))

            status.text("Casando itens e montando a planilha...")
            add_log("Montando tabela comparativa e consolidando dados finais.")

            if master_items is None and gerar_master_auto:
                auto_master = build_master_from_consensus(
                    all_extractions, min_agree=int(min_agree), consensus_threshold=consensus_threshold
                )
                if auto_master:
                    master_items = auto_master
                    add_log(f"Lista mestra automática gerada por consenso com {len(auto_master)} item(ns).")
                    st.info(
                        f"📋 Tabela mestre gerada automaticamente por consenso: "
                        f"{len(auto_master)} item(ns) confirmados por {int(min_agree)}+ orçamentos concordantes."
                    )

            rows, row_index, matrix, review = build_comparison_table(
                all_extractions, master_items=master_items,
                fuzzy_threshold=fuzzy_threshold, sanity_threshold=sanity_threshold,
                usar_uf=usar_uf_casamento, usar_qtd=usar_qtd_casamento,
                bloquear_numero_incoerente=bloquear_numero_incoerente,
            )
            review.extend(review_extracao)
            empresas = sorted({e["empresa"] for e in all_extractions})

            if not rows:
                add_log("Nenhum item válido após consolidação. Encerrando com erro.")
                st.error("Nenhum item foi extraído com sucesso. Verifique os arquivos e a chave de API.")
                st.stop()

            add_log("Gerando arquivo Excel final.")
            excel_buffer = build_excel(rows, matrix, empresas, review, sources)

            st.success("✅ Processamento concluído!")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Itens únicos", len(rows))
            col2.metric("Empresas", len(empresas))
            col3.metric("Novos processados", n_novos)
            col4.metric("Vindos do cache", n_do_cache)

            media_custo_por_arquivo = custo_total_usd / arquivos_com_ia if arquivos_com_ia else 0.0
            suo = " (estimado)" if custo_estimado else ""
            add_log(
                "Processamento finalizado | "
                f"tokens={total_tokens_total} | custo_total=US$ {custo_total_usd:.6f}{suo} | "
                f"media_por_arquivo=US$ {media_custo_por_arquivo:.6f}"
            )
            st.markdown(
                "<p style='font-size:0.80rem;color:#b71c1c;font-weight:700;margin-top:0.45rem;'>"
                f"Consumo IA: {total_tokens_total:,} tokens "
                f"(entrada {prompt_tokens_total:,} | saída {completion_tokens_total:,}) | "
                f"Gasto: US$ {custo_total_usd:.6f}{suo} | "
                f"Média por arquivo com IA: US$ {media_custo_por_arquivo:.6f}"
                "</p>",
                unsafe_allow_html=True,
            )

            if diagnostico_arquivos:
                with st.expander("Diagnóstico da extração", expanded=False):
                    for info in diagnostico_arquivos:
                        st.write(
                            f"- {info['arquivo']}: {info['itens']} item(ns), "
                            f"fonte={info['fonte']}, confiança={info['confianca']}, empresa={info['empresa']}"
                        )

            if review:
                n_desc = sum(1 for r in review if r["tipo"] == "casamento por descrição")
                n_susp = sum(1 for r in review if r["tipo"] == "número igual, descrição muito diferente")
                n_out = sum(1 for r in review if r["tipo"] == "preço fora da curva")
                n_div = sum(1 for r in review if r["tipo"] == "divergência parser vs. IA")
                partes = []
                if n_desc:
                    partes.append(f"{n_desc} casamento(s) por descrição")
                if n_susp:
                    partes.append(f"{n_susp} caso(s) suspeito(s) de número trocado")
                if n_out:
                    partes.append(f"{n_out} preço(s) fora da curva")
                if n_div:
                    partes.append(f"{n_div} divergência(s) parser vs IA")
                st.info(
                    ", ".join(partes) + " precisam de revisão manual "
                    "(veja a aba 'Revisar Casamentos' na planilha)."
                )
            if falhas:
                with st.expander(f"⚠️ {len(falhas)} arquivo(s) com problema"):
                    for msg in falhas:
                        st.write(f"- {msg}")
            if avisos_truncamento:
                with st.expander(f"✂️ {len(avisos_truncamento)} arquivo(s) com texto muito longo (cortado)"):
                    st.write(
                        "Esses arquivos têm mais texto do que o limite enviado à IA — "
                        "alguns itens do final do documento podem não ter sido extraídos:"
                    )
                    for nome in avisos_truncamento:
                        st.write(f"- {nome}")

            st.download_button(
                "⬇️ Baixar Mapa Comparativo (.xlsx)",
                data=excel_buffer,
                file_name="mapa_comparativo_orcamentos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    except Exception:
        st.error("O processamento falhou com uma exceção não tratada.")
        with st.expander("Traceback técnico", expanded=True):
            st.code(traceback.format_exc())
