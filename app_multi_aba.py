import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path # Importar Path para lidar com caminhos de arquivo

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Processador de Abas Excel",
    page_icon="📂",
    layout="wide",
)

# ── Caminho do template pré-definido ──────────────────────────────────────────
# O template TEMPLATE_WHATS_COBRANA_.xlsx deve estar na mesma pasta do app_multi_aba.py
TEMPLATE_FILE_NAME = "TEMPLATE_WHATS_COBRANA_.xlsx"
TEMPLATE_PATH = Path(__file__).parent / TEMPLATE_FILE_NAME

# ── Parse de datas robusto ────────────────────────────────────────────────────
DATE_FORMATS = [
    "%d/%m/%Y",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d-%m-%Y",
    "%d-%m-%Y %H:%M:%S",
    "%Y-%m-%d",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d",
    "%m/%d/%Y",
    "%d.%m.%Y",
]

def parse_dates_robust(series: pd.Series) -> pd.Series:
    result    = pd.Series([pd.NaT] * len(series), index=series.index)
    remaining = series.copy()
    for fmt in DATE_FORMATS:
        mask = result.isna() & remaining.notna()
        if not mask.any():
            break
        parsed = pd.to_datetime(remaining[mask], format=fmt, errors="coerce")
        result[mask] = parsed
    still_null = result.isna() & series.notna()
    if still_null.any():
        result[still_null] = pd.to_datetime(
            series[still_null], infer_datetime_format=True, errors="coerce"
        )
    return result


# ── Carga de arquivo (agora com suporte a múltiplas abas) ─────────────────────
def load_excel_with_sheets(uploaded_file) -> dict[str, pd.DataFrame] | None:
    name = uploaded_file.name.lower()
    try:
        if not (name.endswith(".xlsx") or name.endswith(".xls")):
            st.error("Formato não suportado. Envie um arquivo .xlsx ou .xls.")
            return None

        xls = pd.ExcelFile(uploaded_file)
        sheets_data = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            df.columns = df.columns.str.strip()
            sheets_data[sheet_name] = df
        return sheets_data
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None


# ── Popula o template Excel ───────────────────────────────────────────────────
def populate_template(df_data: pd.DataFrame, template_path: Path) -> bytes:
    try:
        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado em: {template_path}")

        # Carrega o template do caminho especificado
        template_wb = load_workbook(template_path)
        ws = template_wb.active # Assume que o template tem uma aba ativa para preencher

        # Colunas do template (definidas com base no TEMPLATE_WHATS_COBRANA_.xlsx)
        template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

        # Cria um DataFrame com as colunas do template, preenchendo com o que tiver no df_data
        # e vazio para o que não tiver.
        df_to_write = pd.DataFrame(columns=template_cols)
        for col in template_cols:
            if col in df_data.columns:
                df_to_write[col] = df_data[col]
            else:
                df_to_write[col] = '' # Preenche com vazio se a coluna não existir no df_data

        # Escreve os dados no template a partir da segunda linha (assumindo cabeçalho na linha 1)
        for r_idx, row_data in enumerate(df_to_write.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row_data, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Salva o workbook em um buffer de bytes
        output_buffer = io.BytesIO()
        template_wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer.getvalue()
    except Exception as e:
        st.error(f"Erro ao popular o template: {e}")
        return None


# ── Paginação para DataFrames grandes ─────────────────────────────────────────
def paginar(df: pd.DataFrame, key: str, page_size: int = 500):
    if df.empty:
        st.write("Nenhum dado para exibir.")
        return

    total_rows = len(df)
    total_pages = (total_rows - 1) // page_size + 1

    if total_pages > 1:
        page = st.number_input(
            "Página",
            min_value=1,
            max_value=total_pages,
            value=st.session_state.get(f"{key}_page", 1),
            step=1,
            key=f"{key}_page_input",
        )
        st.session_state[f"{key}_page"] = page
        start_idx = (page - 1) * page_size
        end_idx = min(start_idx + page_size, total_rows)
        df_display = df.iloc[start_idx:end_idx]
        st.dataframe(df_display, use_container_width=True, height=450)
        st.caption(f"Exibindo {start_idx + 1}–{end_idx} de {total_rows:,} registros")
    else:
        st.dataframe(df, use_container_width=True, height=450)
        st.caption(f"Exibindo {total_rows:,} registros")


# ── Detecção de duplicidades (corrigida e completa) ───────────────────────────
def detect_duplicates(
    df: pd.DataFrame,
    col_cliente: str,
    col_servico: str,
    col_data: str,
    janela_dias: int,
    col_os: str | None = None,
    cols_extras: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna (df_det, df_resumo_grupos, df_resumo_servico).

    Regras:
      - Duplicidade = mesmo cliente + mesmo tipo de serviço
        com datas dentro da janela de janela_dias dias.
      - A PRIMEIRA OS do grupo = ORIGINAL, não entra na contagem.
      - As demais = DUPLICATA, entram na contagem.
      - Uma OS já marcada DUPLICATA não pode virar âncora de novo grupo.
      - Ponteiro avança i += 1 sempre (não pula o cluster inteiro),
        garantindo que cada OS seja avaliada como potencial âncora.
    """

    work = df.copy()

    # ── Parsing de datas ──────────────────────────────────────────────────────
    work["_data_parsed"] = parse_dates_robust(work[col_data].astype(str).str.strip())

    total     = len(work)
    invalidas = work["_data_parsed"].isna().sum()
    validas   = total - invalidas

    with st.expander("📅 Diagnóstico de parsing de datas", expanded=invalidas > 0):
        d1, d2, d3 = st.columns(3)
        d1.metric("Total de registros",       f"{total:,}")
        d2.metric("Datas reconhecidas",        f"{validas:,}")
        d3.metric("Datas inválidas/ignoradas", f"{invalidas:,}",
                  delta=f"-{invalidas}" if invalidas else None,
                  delta_color="inverse")
        if invalidas > 0:
            exemplos = work[work["_data_parsed"].isna()][col_data].dropna().unique()[:10]
            st.warning(f"Exemplos não reconhecidos: `{'`, `'.join(map(str, exemplos))}`")
        else:
            amostra = work[[col_data, "_data_parsed"]].dropna().head(5).copy()
            amostra["_data_parsed"] = amostra["_data_parsed"].dt.strftime("%d/%m/%Y")
            amostra.columns = ["Valor original", "Interpretado como"]
            st.success("Todas as datas foram reconhecidas com sucesso.")
            st.dataframe(amostra, use_container_width=True)

    work = work.dropna(subset=["_data_parsed"]).copy()
    work = work.reset_index(drop=True)
    work["_row_id"]       = work.index
    work["_cliente_norm"] = work[col_cliente].astype(str).str.strip().str.upper()
    work["_servico_norm"] = work[col_servico].astype(str).str.strip().str.upper()

    # Remover linhas onde cliente ou serviço são nulos após normalização
    original_len = len(work)
    work.dropna(subset=["_cliente_norm", "_servico_norm"], inplace=True)
    if len(work) < original_len:
        st.warning(f"⚠️ {original_len - len(work)} linhas foram removidas da análise por terem Cliente ou Serviço vazios/nulos.")

    # ── Algoritmo de detecção ─────────────────────────────────────────────────
    #
    # Para cada par (cliente, serviço):
    #   1. Ordena por data crescente
    #   2. Percorre cada OS como possível âncora (ORIGINAL)
    #   3. Se já foi marcada DUPLICATA → pula, não vira âncora
    #   4. Busca todas as OS posteriores não classificadas dentro da janela
    #   5. Se achou ≥ 1 → cria grupo: âncora = ORIGINAL, demais = DUPLICATA
    #   6. Avança i += 1 (não pula o cluster inteiro!)
    #

    classificacoes = [] # Lista para armazenar as classificações de todas as OS
    global_grupo_counter = 0 # Contador GLOBAL para IDs de grupo únicos

    # Agrupa por cliente e serviço normalizados, depois ordena por data
    grouped = work.groupby(["_cliente_norm", "_servico_norm"], sort=False)

    for (cliente_norm, servico_norm), grp in grouped:
        grp = grp.sort_values("_data_parsed").reset_index(drop=True)

        # Mapeamento local de _row_id para a classificação já feita
        # Isso evita que uma OS já marcada como DUPLICATA seja reprocessada como ORIGINAL
        local_classificacao_map = {c["_row_id"]: c for c in classificacoes if c["_cliente_norm"] == cliente_norm and c["_servico_norm"] == servico_norm}

        i = 0
        while i < len(grp):
            current_os = grp.iloc[i]
            current_row_id = current_os["_row_id"]

            # Se esta OS já foi classificada como DUPLICATA em um grupo anterior, pule-a
            if current_row_id in local_classificacao_map and local_classificacao_map[current_row_id]["tipo_registro"] == "DUPLICATA":
                i += 1
                continue

            # Inicia um novo grupo de duplicidades
            global_grupo_counter += 1
            grupo_id = global_grupo_counter

            # Adiciona a OS atual como ORIGINAL
            classificacoes.append({
                "_row_id": current_row_id,
                "grupo_duplicidade": grupo_id,
                "tipo_registro": "ORIGINAL",
                "_cliente_norm": cliente_norm,
                "_servico_norm": servico_norm,
            })
            local_classificacao_map[current_row_id] = classificacoes[-1] # Atualiza o mapa local

            # Busca por duplicatas posteriores dentro da janela
            found_duplicates_in_group = False
            for j in range(i + 1, len(grp)):
                next_os = grp.iloc[j]
                next_row_id = next_os["_row_id"]

                # Se a próxima OS já foi classificada, ou se já está muito distante no tempo, pare
                if next_row_id in local_classificacao_map:
                    # Se já é ORIGINAL de outro grupo, não pode ser duplicata deste
                    if local_classificacao_map[next_row_id]["tipo_registro"] == "ORIGINAL":
                        continue
                    # Se já é DUPLICATA deste grupo, não precisa adicionar de novo
                    if local_classificacao_map[next_row_id]["grupo_duplicidade"] == grupo_id:
                        continue
                    # Se já é DUPLICATA de outro grupo, não pode ser duplicata deste
                    if local_classificacao_map[next_row_id]["tipo_registro"] == "DUPLICATA":
                        continue

                time_diff = (next_os["_data_parsed"] - current_os["_data_parsed"]).days
                if 0 <= time_diff <= janela_dias:
                    # Marca como DUPLICATA
                    classificacoes.append({
                        "_row_id": next_row_id,
                        "grupo_duplicidade": grupo_id,
                        "tipo_registro": "DUPLICATA",
                        "_cliente_norm": cliente_norm,
                        "_servico_norm": servico_norm,
                    })
                    local_classificacao_map[next_row_id] = classificacoes[-1] # Atualiza o mapa local
                    found_duplicates_in_group = True
                elif time_diff > janela_dias:
                    # Se a diferença de tempo excedeu a janela, não há mais duplicatas para esta OS
                    break

            # Se a OS atual foi uma ORIGINAL mas não teve nenhuma duplicata,
            # e ela não foi marcada como ORIGINAL de outro grupo,
            # podemos considerá-la como não-duplicada (ou ORIGINAL única).
            # No entanto, a lógica atual já a classifica como ORIGINAL e a inclui no df_det.
            # O importante é que ela não seja contada como "duplicata" no resumo.

            i += 1 # Avança para a próxima OS como potencial âncora

    # ── Montagem dos DataFrames de saída ──────────────────────────────────────
    if not classificacoes:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_classificacao = pd.DataFrame(classificacoes)

    # 1) Detalhamento
    # Mescla as classificações de volta com o DataFrame original
    df_det = pd.merge(work, df_classificacao, on="_row_id", how="left")

    # Preenche OSs não duplicadas (sem grupo) com tipo_registro 'ÚNICA'
    df_det["tipo_registro"].fillna("ÚNICA", inplace=True)
    df_det["grupo_duplicidade"].fillna(0, inplace=True) # Grupo 0 para únicas

    # Seleciona e ordena as colunas para o detalhamento
    output_cols = ["grupo_duplicidade", "tipo_registro"]
    if col_os:
        output_cols.append(col_os)
    output_cols += [col_cliente, col_servico, col_data]
    if cols_extras:
        output_cols += [c for c in cols_extras if c not in output_cols]

    # Garante que as colunas existam no df_det antes de selecionar
    final_det_cols = [c for c in output_cols if c in df_det.columns]
    df_det = df_det[final_det_cols].copy()
    df_det.rename(columns={"grupo_duplicidade": "grupo"}, inplace=True)
    df_det.sort_values(["grupo", "_data_parsed"], inplace=True)

    # 2) Resumo por grupo
    resumo_grupos = []
    # Filtra apenas os grupos que têm pelo menos uma DUPLICATA
    grupos_com_duplicata = df_det[df_det["tipo_registro"] == "DUPLICATA"]["grupo"].unique()

    for gid in sorted(grupos_com_duplicata):
        grupo_df = df_det[df_det["grupo"] == gid].copy()
        original_os = grupo_df[grupo_df["tipo_registro"] == "ORIGINAL"].iloc[0]
        dups = grupo_df[grupo_df["tipo_registro"] == "DUPLICATA"]

        datas_dup = dups["_data_parsed"] if not dups.empty else pd.Series(dtype='datetime64[ns]')

        row = {
            "grupo":                   int(gid),
            "cliente":                 original_os[col_cliente],
            "tipo_servico":            original_os[col_servico],
            "data_os_original":        original_os["_data_parsed"].strftime("%d/%m/%Y"),
            "qtd_duplicatas":          len(dups),
            "data_primeira_duplicata": datas_dup.min().strftime("%d/%m/%Y") if not dups.empty else "—",
            "data_ultima_duplicata":   datas_dup.max().strftime("%d/%m/%Y") if not dups.empty else "—",
            "intervalo_dias":          (datas_dup.max() - datas_dup.min()).days if not dups.empty else 0,
        }
        if col_os:
            row["os_original"]   = str(original_os[col_os])
            row["os_duplicadas"] = ", ".join(dups[col_os].astype(str).tolist()) if not dups.empty else "—"
        resumo_grupos.append(row)

    df_resumo_grupos = pd.DataFrame(resumo_grupos)

    # 3) Resumo por tipo de serviço
    cont_cli_serv = (
        work.groupby(["_servico_norm", "_cliente_norm"])
        .size()
        .reset_index(name="total_os_cliente")
    )

    resumo_servico = []
    dups_only = df_det[df_det["tipo_registro"] == "DUPLICATA"]

    for servico_norm, grp_serv in dups_only.groupby("_servico_norm"):
        servico_val  = grp_serv[col_servico].iloc[0]
        total_dup    = grp_serv.shape[0]
        clientes_dup = grp_serv["_cliente_norm"].nunique()
        total_os     = work[work["_servico_norm"] == servico_norm].shape[0]

        dist = cont_cli_serv[
            cont_cli_serv["_servico_norm"] == servico_norm
        ]["total_os_cliente"]

        resumo_servico.append({
            "tipo_servico":                 servico_val,
            "total_os_no_periodo":          int(total_os),
            "total_duplicatas":             int(total_dup),
            "clientes_com_duplicata":       int(clientes_dup),
            "media_duplicatas_por_cliente": round(total_dup / clientes_dup, 2) if clientes_dup else 0,
            "clientes_1_pedido":            int((dist == 1).sum()),
            "clientes_2_pedidos":           int((dist == 2).sum()),
            "clientes_3_pedidos":           int((dist == 3).sum()),
            "clientes_4_a_6_pedidos":       int(((dist >= 4) & (dist <= 6)).sum()),
            "clientes_7_a_10_pedidos":      int(((dist >= 7) & (dist <= 10)).sum()),
            "clientes_mais_10_pedidos":     int((dist > 10)).sum()),
        })

    df_resumo_servico = pd.DataFrame(resumo_servico).sort_values(
        "total_duplicatas", ascending=False
    ).reset_index(drop=True)

    return df_det, df_resumo_grupos, df_resumo_servico


# ── Interface do Streamlit ────────────────────────────────────────────────────
st.title("📂 Processador de Abas Excel com Detecção de Duplicidades")
st.markdown(
    "Este aplicativo processa um arquivo Excel com múltiplas abas, "
    "detecta ordens de serviço duplicadas em cada aba e gera arquivos "
    "Excel individuais preenchidos com as duplicatas, usando um template pré-definido."
)

# ── Sidebar para upload e configurações ───────────────────────────────────────
with st.sidebar:
    st.header("Upload de Arquivo")
    uploaded_file = st.file_uploader(
        "Arraste e solte seu arquivo Excel (.xlsx ou .xls) aqui",
        type=["xlsx", "xls"],
        help="O arquivo pode conter múltiplas abas, cada uma será processada individualmente."
    )

    if uploaded_file:
        st.session_state["uploaded_file_name"] = uploaded_file.name
        st.session_state["sheets_data"] = load_excel_with_sheets(uploaded_file)
        if st.session_state["sheets_data"]:
            st.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")
            st.write(f"Abas encontradas: {', '.join(st.session_state['sheets_data'].keys())}")
        else:
            st.error("Nenhuma aba válida encontrada no arquivo.")
            st.session_state["sheets_data"] = None
    else:
        st.session_state["sheets_data"] = None
        st.session_state["uploaded_file_name"] = None

    st.markdown("---")
    st.info(f"Template de saída: **{TEMPLATE_FILE_NAME}** (deve estar na mesma pasta do app)")
    if not TEMPLATE_PATH.exists():
        st.error(f"⚠️ O arquivo de template '{TEMPLATE_FILE_NAME}' não foi encontrado na pasta do aplicativo. Por favor, coloque-o lá para que o app funcione.")


# ── Corpo principal do app ────────────────────────────────────────────────────
if not st.session_state.get("sheets_data"):
    st.warning("Por favor, faça o upload de um arquivo Excel na barra lateral para começar.")
else:
    st.subheader("Seleção de Abas para Processar")
    all_sheet_names = list(st.session_state["sheets_data"].keys())
    selected_sheets = st.multiselect(
        "Selecione as abas que deseja processar:",
        options=all_sheet_names,
        default=all_sheet_names, # Seleciona todas por padrão
        help="Apenas as abas selecionadas serão analisadas e terão arquivos gerados."
    )

    if not selected_sheets:
        st.warning("Por favor, selecione pelo menos uma aba para processar.")
    else:
        st.info(f"Serão processadas {len(selected_sheets)} abas.")

        # Configurações de análise de duplicidades (com base na primeira aba selecionada)
        st.subheader("⚙️ Configurações da análise de duplicidades (aplicadas a todas as abas)")

        # Usamos a primeira aba selecionada para inferir as colunas
        first_sheet_df = st.session_state["sheets_data"][selected_sheets[0]]
        all_cols = first_sheet_df.columns.tolist()

        def suggest(keywords, cols):
            for kw in keywords:
                for col in cols:
                    if kw.lower() in col.lower():
                        return col
            return cols[0] if cols else None

        c1, c2, c3 = st.columns(3)
        col_cliente = c1.selectbox(
            "👤 Coluna de cliente",
            all_cols,
            index=all_cols.index(suggest(["cliente", "matricula", "cpf", "cod"], all_cols) or all_cols[0]),
        )
        col_servico = c2.selectbox(
            "🔧 Coluna de tipo de serviço",
            all_cols,
            index=all_cols.index(suggest(["servico", "serviço", "tipo", "categoria"], all_cols) or all_cols[0]),
        )
        col_data = c3.selectbox(
            "📅 Coluna de data da OS",
            all_cols,
            index=all_cols.index(suggest(["data", "dt_", "abertura", "criacao", "criação"], all_cols) or all_cols[0]),
        )

        c4, c5 = st.columns(2)
        none_opt = "— nenhuma —"
        col_os_raw = c4.selectbox(
            "🔢 Coluna de número da OS (opcional)",
            [none_opt] + all_cols,
            index=0,
        )
        col_os = None if col_os_raw == none_opt else col_os_raw

        janela_dias = c5.number_input(
            "📆 Janela de tempo para considerar duplicidade (dias)",
            min_value=1, max_value=3650, value=30, step=1,
        )

        extras_disponiveis = [c for c in all_cols if c not in [col_cliente, col_servico, col_data, col_os]]
        cols_extras = st.multiselect(
            "➕ Colunas adicionais para exibir no relatório (opcional)",
            extras_disponiveis,
        )

        st.divider()

        if st.button("🚀 Processar Abas e Gerar Arquivos", type="primary", use_container_width=True):
            st.session_state["processed_files"] = {}
            progress_bar = st.progress(0)
            status_text = st.empty()

            if not TEMPLATE_PATH.exists():
                st.error(f"Erro: O arquivo de template '{TEMPLATE_FILE_NAME}' não foi encontrado. Por favor, coloque-o na mesma pasta do aplicativo.")
            else:
                for i, sheet_name in enumerate(selected_sheets):
                    status_text.text(f"Processando aba: {sheet_name} ({i+1}/{len(selected_sheets)})")
                    df_aba = st.session_state["sheets_data"][sheet_name].copy()

                    # 1. Detectar duplicidades para a aba
                    df_det, df_grupos, df_servico = detect_duplicates(
                        df=df_aba,
                        col_cliente=col_cliente,
                        col_servico=col_servico,
                        col_data=col_data,
                        janela_dias=int(janela_dias),
                        col_os=col_os,
                        cols_extras=cols_extras if cols_extras else None,
                    )

                    # Filtrar apenas as duplicatas para o template de cobrança
                    df_cobranca = df_det[df_det["tipo_registro"] == "DUPLICATA"].copy()

                    # Mapear e preparar df_cobranca para o template
                    # Este é o ponto CRÍTICO que você precisa ajustar para suas colunas
                    # Exemplo de mapeamento:
                    df_cobranca_template = pd.DataFrame()
                    if not df_cobranca.empty:
                        # Adapte estas linhas para mapear as colunas do seu df_cobranca
                        # para as colunas do template (MATRICULA, TELEFONE, etc.)
                        # Exemplo:
                        df_cobranca_template['MATRICULA'] = df_cobranca[col_cliente] # Usa a coluna de cliente selecionada
                        df_cobranca_template['TELEFONE'] = df_cobranca['TELEFONE_CONTATO'] if 'TELEFONE_CONTATO' in df_cobranca.columns else '' # Exemplo: se tiver coluna de telefone
                        df_cobranca_template['CONCESSIONARIA'] = 'SUA_CONCESSIONARIA' # Valor fixo ou de outra coluna
                        df_cobranca_template['CIDADE'] = df_cobranca['CIDADE_OS'] if 'CIDADE_OS' in df_cobranca.columns else '' # Exemplo
                        df_cobranca_template['DIRETORIA'] = 'SUA_DIRETORIA' # Valor fixo ou de outra coluna
                        df_cobranca_template['SITUACAO'] = df_cobranca[col_servico] # Usa a coluna de serviço selecionada

                        # Garante que todas as colunas do template existam, preenchendo com vazio se faltar
                        for t_col in ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']:
                            if t_col not in df_cobranca_template.columns:
                                df_cobranca_template[t_col] = ''

                    # 2. Popular o template com os dados da aba (apenas duplicatas)
                    if not df_cobranca_template.empty:
                        output_file_bytes = populate_template(df_cobranca_template, TEMPLATE_PATH)
                        if output_file_bytes:
                            st.session_state["processed_files"][sheet_name] = output_file_bytes
                    else:
                        st.info(f"Nenhuma duplicata encontrada na aba '{sheet_name}' para gerar arquivo de cobrança.")

                    progress_bar.progress((i + 1) / len(selected_sheets))
                status_text.success("Processamento concluído para todas as abas selecionadas!")
                st.rerun()

    if "processed_files" in st.session_state and st.session_state["processed_files"]:
        st.subheader("Arquivos Gerados")
        st.markdown("Clique para baixar os arquivos de cobrança gerados por aba:")
        for sheet_name, file_bytes in st.session_state["processed_files"].items():
            st.download_button(
                label=f"⬇️ Baixar {sheet_name}.xlsx",
                data=file_bytes,
                file_name=f"{sheet_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{sheet_name}"
            )
