import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Populador de Template Excel por Aba",
    page_icon="📄",
    layout="wide",
)

# ── Caminho do template pré-definido ──────────────────────────────────────────
# O template TEMPLATE_WHATS_COBRANCA.xlsx deve estar na mesma pasta do app.py
TEMPLATE_FILE_NAME = "TEMPLATE_WHATS_COBRANCA.xlsx"
TEMPLATE_PATH = Path(__file__).parent / TEMPLATE_FILE_NAME

# ── Carga de arquivo Excel com múltiplas abas ─────────────────────────────────
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
            df.columns = df.columns.str.strip() # Limpa espaços em branco dos nomes das colunas
            sheets_data[sheet_name] = df
        return sheets_data
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None


# ── Popula o template Excel ───────────────────────────────────────────────────
def populate_template(df_data: pd.DataFrame, template_path: Path, column_mapping: dict) -> bytes | None:
    try:
        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado em: {template_path}")

        template_wb = load_workbook(template_path)
        ws = template_wb.active # Assume que o template tem uma aba ativa para preencher

        # Colunas do template (definidas com base no TEMPLATE_WHATS_COBRANCA.xlsx)
        template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

        # Cria um DataFrame temporário para organizar os dados conforme o template
        df_to_write = pd.DataFrame(columns=template_cols)

        # Preenche o df_to_write usando o mapeamento fornecido
        for template_col, source_col in column_mapping.items():
            if template_col in template_cols: # Garante que a coluna existe no template
                if source_col in df_data.columns:
                    df_to_write[template_col] = df_data[source_col]
                else:
                    # Se a coluna de origem não existe, preenche com vazio ou valor padrão
                    df_to_write[template_col] = ''

        # Para colunas do template que não foram mapeadas, preenche com vazio
        for t_col in template_cols:
            if t_col not in df_to_write.columns:
                df_to_write[t_col] = ''

        # Escreve os dados no template a partir da segunda linha (assumindo cabeçalho na linha 1)
        for r_idx, row_data in enumerate(df_to_write.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row_data, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Salva o workbook em um buffer de bytes
        output_buffer = io.BytesIO()
        template_wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer.getvalue()

    except FileNotFoundError as e:
        st.error(f"Erro: {e}. Por favor, verifique se o arquivo '{TEMPLATE_FILE_NAME}' está na mesma pasta do aplicativo.")
        return None
    except Exception as e:
        st.error(f"Erro ao popular o template: {e}")
        return None


# ── Interface do Streamlit ────────────────────────────────────────────────────
st.title("Populador de Template Excel por Aba")
st.markdown("---")

# Verifica se o template existe
if not TEMPLATE_PATH.exists():
    st.error(f"**Erro:** O arquivo de template '{TEMPLATE_FILE_NAME}' não foi encontrado na mesma pasta do aplicativo.")
    st.info("Por favor, coloque o arquivo `TEMPLATE_WHATS_COBRANCA.xlsx` ao lado do `app_populador_template.py`.")
else:
    st.sidebar.header("1. Carregar Arquivo de Entrada")
    uploaded_file = st.sidebar.file_uploader(
        "Selecione um arquivo Excel com múltiplas abas", type=["xlsx", "xls"]
    )

    sheets_data = None
    if uploaded_file:
        sheets_data = load_excel_with_sheets(uploaded_file)
        if sheets_data:
            st.sidebar.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")
            st.sidebar.write(f"Abas encontradas: {', '.join(sheets_data.keys())}")

            st.sidebar.header("2. Selecionar Abas para Processar")
            selected_sheets = st.sidebar.multiselect(
                "Escolha as abas que deseja processar:",
                options=list(sheets_data.keys()),
                default=list(sheets_data.keys()) # Seleciona todas por padrão
            )

            if selected_sheets:
                st.sidebar.header("3. Mapeamento de Colunas")
                st.sidebar.info("Mapeie as colunas do seu arquivo de entrada para as colunas do template.")

                # Pega as colunas da primeira aba selecionada para sugerir no mapeamento
                # (assumindo que as colunas são consistentes entre as abas)
                first_selected_df_cols = [''] # Opção vazia
                if selected_sheets and sheets_data:
                    first_selected_df_cols.extend(sheets_data[selected_sheets[0]].columns.tolist())

                # Colunas do template
                template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

                # Dicionário para armazenar o mapeamento
                column_mapping = {}

                for t_col in template_cols:
                    # Se for MATRICULA, TELEFONE, CIDADE, SITUACAO, permite mapear para uma coluna do input
                    if t_col in ['MATRICULA', 'TELEFONE', 'CIDADE', 'SITUACAO']:
                        selected_source_col = st.sidebar.selectbox(
                            f"Coluna do template '{t_col}' será preenchida por:",
                            options=first_selected_df_cols,
                            key=f"map_{t_col}"
                        )
                        if selected_source_col:
                            column_mapping[t_col] = selected_source_col
                    # Para CONCESSIONARIA e DIRETORIA, permite um valor fixo ou mapeamento
                    elif t_col in ['CONCESSIONARIA', 'DIRETORIA']:
                        mapping_type = st.sidebar.radio(
                            f"Como preencher '{t_col}'?",
                            ('Valor Fixo', 'Mapear Coluna'),
                            key=f"type_map_{t_col}"
                        )
                        if mapping_type == 'Valor Fixo':
                            fixed_value = st.sidebar.text_input(
                                f"Valor fixo para '{t_col}':",
                                value=f"Minha {t_col}" if t_col == 'CONCESSIONARIA' else f"Minha {t_col}",
                                key=f"fixed_val_{t_col}"
                            )
                            column_mapping[t_col] = fixed_value # Armazena o valor fixo
                        else: # Mapear Coluna
                            selected_source_col = st.sidebar.selectbox(
                                f"Coluna do template '{t_col}' será preenchida por:",
                                options=first_selected_df_cols,
                                key=f"map_{t_col}"
                            )
                            if selected_source_col:
                                column_mapping[t_col] = selected_source_col

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
                            status_text.info(f"Processando aba: '{sheet_name}'...")
                            df_aba = sheets_data[sheet_name]

                            # Popula o template com os dados da aba usando o mapeamento
                            output_file_bytes = populate_template(df_aba, TEMPLATE_PATH, column_mapping)

                            if output_file_bytes:
                                st.session_state["processed_files"][sheet_name] = output_file_bytes

                            progress_bar.progress((i + 1) / len(selected_sheets))

                        status_text.success("Processamento concluído para todas as abas selecionadas!")
                        st.rerun()

    if "processed_files" in st.session_state and st.session_state["processed_files"]:
        st.subheader("Arquivos Gerados")
        st.markdown("Clique para baixar os arquivos Excel gerados por aba:")
        for sheet_name, file_bytes in st.session_state["processed_files"].items():
            st.download_button(
                label=f"⬇️ Baixar {sheet_name}.xlsx",
                data=file_bytes,
                file_name=f"{sheet_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{sheet_name}"
            )
