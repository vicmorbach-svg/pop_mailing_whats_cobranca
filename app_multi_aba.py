import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from pathlib import Path

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Populador de Template Excel por Aba",
    page_icon="📂",
    layout="wide",
)

# ── Caminho do template pré-definido ──────────────────────────────────────────
TEMPLATE_FILE_NAME = "TEMPLATE_WHATS_COBRANCA.xlsx"
TEMPLATE_PATH = Path(__file__).parent / TEMPLATE_FILE_NAME

# ── Carga de arquivo (suporte a múltiplas abas e Parquet) ───────────────────────────────
@st.cache_data(ttl=3600)
def load_data_from_file(uploaded_file_bytes: bytes, file_name: str) -> dict[str, pd.DataFrame] | None:
    """
    Carrega dados de um arquivo Excel (xlsx, xls) ou Parquet.
    Retorna um dicionário onde a chave é o nome da aba/arquivo e o valor é o DataFrame.
    """
    name = file_name.lower()
    try:
        sheets_data = {}
        if name.endswith(".parquet"):
            st.info("Lendo arquivo Parquet... Isso é ótimo para desempenho com arquivos grandes!")
            df = pd.read_parquet(io.BytesIO(uploaded_file_bytes))
            df.columns = df.columns.str.strip()
            sheets_data["Dados_Parquet"] = df
            st.success("Arquivo Parquet carregado com sucesso!")
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            st.info("Lendo arquivo Excel... Para arquivos muito grandes, Parquet é mais recomendado.")
            xls = pd.ExcelFile(io.BytesIO(uploaded_file_bytes))
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
                df.columns = df.columns.str.strip()
                sheets_data[sheet_name] = df
            st.success("Arquivo Excel carregado com sucesso!")
        else:
            st.error("Formato não suportado. Envie um arquivo .xlsx, .xls ou .parquet.")
            return None
        return sheets_data
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None


# ── Popula o template Excel ───────────────────────────────────────────────────
def populate_template(df_data: pd.DataFrame, template_path: Path, column_mapping: dict) -> tuple[bytes | None, int]:
    try:
        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado em: {template_path}")

        template_wb = load_workbook(template_path)
        ws = template_wb.active

        template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

        processed_rows_list = []
        for _, row_data_input in df_data.iterrows():
            new_row_dict = {}
            for t_col in template_cols:
                mapped_value = column_mapping.get(t_col)
                if mapped_value:
                    if mapped_value in row_data_input.index:
                        new_row_dict[t_col] = row_data_input[mapped_value]
                    else:
                        new_row_dict[t_col] = mapped_value
                else:
                    new_row_dict[t_col] = ''
            processed_rows_list.append(new_row_dict)

        df_to_write = pd.DataFrame(processed_rows_list, columns=template_cols)

        # Remove linhas completamente vazias do DataFrame antes de escrever
        df_to_write.dropna(how='all', inplace=True)

        # Contagem de clientes (linhas válidas)
        client_count = len(df_to_write)

        # Adiciona os dados ao template
        for r_idx, row_data in enumerate(df_to_write.itertuples(index=False)):
            ws.append(list(row_data))

        # Limpeza rigorosa: Remove qualquer linha em branco no final do arquivo Excel
        while ws.max_row > 1:
            # Verifica se todos os valores da última linha são nulos ou strings vazias
            if all(cell.value is None or str(cell.value).strip() == '' for cell in ws[ws.max_row]):
                ws.delete_rows(ws.max_row)
            else:
                break

        # Salva o workbook em um buffer de bytes
        output_buffer = io.BytesIO()
        template_wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer.getvalue(), client_count

    except FileNotFoundError as e:
        st.error(f"Erro: {e}. Por favor, verifique se o arquivo '{TEMPLATE_FILE_NAME}' está na mesma pasta do aplicativo.")
        return None, 0
    except Exception as e:
        st.error(f"Erro ao popular o template: {e}")
        return None, 0


# ── Interface do Streamlit ────────────────────────────────────────────────────
st.title("Populador de Template Excel por Aba")
st.markdown("---")

# Verifica se o template existe
if not TEMPLATE_PATH.exists():
    st.error(f"**Erro:** O arquivo de template '{TEMPLATE_FILE_NAME}' não foi encontrado na mesma pasta do aplicativo.")
    st.info("Por favor, coloque o arquivo `TEMPLATE_WHATS_COBRANCA.xlsx` ao lado do script.")
else:
    st.sidebar.header("1. Carregar Arquivo de Entrada")
    uploaded_file = st.sidebar.file_uploader(
        "Selecione um arquivo (Excel ou Parquet)", type=["xlsx", "xls", "parquet"]
    )

    sheets_data = None
    if uploaded_file:
        sheets_data = load_data_from_file(uploaded_file.getvalue(), uploaded_file.name)
        if sheets_data:
            st.sidebar.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")
            st.sidebar.write(f"Abas/Datasets encontrados: {', '.join(sheets_data.keys())}")

            st.sidebar.header("2. Selecionar Abas/Datasets para Processar")
            selected_sheets = st.sidebar.multiselect(
                "Escolha as abas/datasets que deseja processar:",
                options=list(sheets_data.keys()),
                default=list(sheets_data.keys())
            )

            if selected_sheets:
                st.sidebar.header("3. Mapeamento de Colunas")
                st.sidebar.info("Mapeie as colunas do seu arquivo de entrada para as colunas do template.")

                first_selected_df_cols = ['']
                if selected_sheets and sheets_data:
                    first_selected_df_cols.extend(sheets_data[selected_sheets[0]].columns.tolist())

                template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']
                column_mapping = {}

                for t_col in template_cols:
                    mapping_type = st.sidebar.radio(
                        f"Como preencher '{t_col}'?",
                        ('Mapear Coluna', 'Valor Fixo'),
                        key=f"type_map_{t_col}",
                        index=0 if t_col not in ['CONCESSIONARIA'] else 1
                    )
                    if mapping_type == 'Mapear Coluna':
                        selected_source_col = st.sidebar.selectbox(
                            f"Coluna do seu arquivo para '{t_col}':",
                            options=first_selected_df_cols,
                            key=f"map_{t_col}"
                        )
                        if selected_source_col:
                            column_mapping[t_col] = selected_source_col
                    else:
                        fixed_value = st.sidebar.text_input(
                            f"Valor fixo para '{t_col}':",
                            value="" if t_col not in ['CONCESSIONARIA'] else f"Corsan",
                            key=f"fixed_val_{t_col}"
                        )
                        column_mapping[t_col] = fixed_value

                st.sidebar.markdown("---")
                st.sidebar.header("4. Gerar Arquivos")
                if st.sidebar.button("🚀 Processar Abas e Gerar Arquivos"):
                    if not TEMPLATE_PATH.exists():
                        st.error(f"Erro: O arquivo de template '{TEMPLATE_FILE_NAME}' não foi encontrado.")
                    elif not column_mapping:
                        st.warning("Por favor, configure o mapeamento de colunas antes de processar.")
                    else:
                        st.session_state["processed_files"] = {}
                        status_text = st.empty()
                        progress_bar = st.progress(0)

                        for i, sheet_name in enumerate(selected_sheets):
                            status_text.info(f"Processando aba/dataset: '{sheet_name}'...")
                            df_aba = sheets_data[sheet_name]

                            output_file_bytes, client_count = populate_template(df_aba, TEMPLATE_PATH, column_mapping)

                            if output_file_bytes:
                                st.session_state["processed_files"][sheet_name] = {
                                    "bytes": output_file_bytes,
                                    "count": client_count
                                }

                            progress_bar.progress((i + 1) / len(selected_sheets))

                        status_text.success("Processamento concluído para todas as abas/datasets selecionados!")
                        st.rerun()

    if "processed_files" in st.session_state and st.session_state["processed_files"]:
        st.subheader("Arquivos Gerados")
        st.markdown("Clique para baixar os arquivos Excel gerados por aba/dataset:")

        for sheet_name, file_data in st.session_state["processed_files"].items():
            col1, col2 = st.columns([3, 1]) # Divide o espaço para o botão e a contagem

            with col1:
                st.download_button(
                    label=f"⬇️ Baixar {sheet_name}.xlsx",
                    data=file_data["bytes"],
                    file_name=f"{sheet_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{sheet_name}"
                )
            with col2:
                # Exibe a contagem de clientes processados
                st.info(f"👥 Clientes: **{file_data['count']}**")

    st.sidebar.markdown("---")
    st.sidebar.header("Opções de Cache")
    if st.sidebar.button("Limpar Cache de Dados Carregados"):
        st.cache_data.clear()
        st.sidebar.success("Cache de dados limpo! O aplicativo será recarregado para garantir novos dados.")
        st.rerun()
