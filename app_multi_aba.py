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
@st.cache_data(ttl=3600) # Cache por 1 hora, ou ajuste conforme a necessidade
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
            # Otimização de tipos de dados: Se souber os tipos, defina-os aqui para economizar memória.
            # Ex: df['MATRICULA'] = pd.to_numeric(df['MATRICULA'], errors='coerce').astype('Int64')
            # df['TELEFONE'] = df['TELEFONE'].astype(str) # Garante que telefone é string
            sheets_data["Dados_Parquet"] = df
            st.success("Arquivo Parquet carregado com sucesso!")
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            st.info("Lendo arquivo Excel... Para arquivos muito grandes, Parquet é mais recomendado.")
            xls = pd.ExcelFile(io.BytesIO(uploaded_file_bytes))
            for sheet_name in xls.sheet_names:
                # Usecols pode ser útil aqui se você não precisar de todas as colunas
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
def populate_template(df_data: pd.DataFrame, template_path: Path, column_mapping: dict) -> bytes | None:
    try:
        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado em: {template_path}")

        template_wb = load_workbook(template_path)
        ws = template_wb.active

        template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

        # --- INÍCIO DA OTIMIZAÇÃO: Construir lista de dicionários e criar DataFrame de uma vez ---
        processed_rows_list = []
        for _, row_data_input in df_data.iterrows():
            new_row_dict = {}
            for t_col in template_cols:
                mapped_value = column_mapping.get(t_col)
                if mapped_value:
                    if mapped_value in row_data_input.index: # Verifica se é uma coluna do input
                        new_row_dict[t_col] = row_data_input[mapped_value]
                    else: # Valor fixo
                        new_row_dict[t_col] = mapped_value
                else:
                    new_row_dict[t_col] = ''
            processed_rows_list.append(new_row_dict)

        df_to_write = pd.DataFrame(processed_rows_list, columns=template_cols)
        # --- FIM DA OTIMIZAÇÃO ---

        # Escreve os dados no template a partir da segunda linha (assumindo cabeçalho na linha 1)
        # Usando ws.append para um desempenho potencialmente melhor em grandes volumes
        # Primeiro, limpa as linhas existentes se necessário (cuidado para não apagar o cabeçalho)
        # Se o template já tem cabeçalho na linha 1 e você quer adicionar a partir da linha 2:
        # Você pode apagar as linhas existentes a partir da linha 2 antes de adicionar
        # for row in ws.iter_rows(min_row=2):
        #     for cell in row:
        #         cell.value = None # Limpa o conteúdo
        # Ou, se o template é sempre vazio abaixo do cabeçalho, apenas comece a adicionar.

        # Adiciona os dados do DataFrame ao template
        # O openpyxl.Workbook.append() é geralmente mais rápido para adicionar muitas linhas
        # do que iterar célula por célula.
        for r_idx, row_data in enumerate(df_to_write.itertuples(index=False)):
            ws.append(list(row_data)) # Adiciona cada linha como uma lista

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





# ── Interface do Streamlit ────────────────────────────────────────────────────
st.title("Populador de Template Excel por Aba")
st.markdown("---")

# Verifica se o template existe
if not TEMPLATE_PATH.exists():
    st.error(f"**Erro:** O arquivo de template '{TEMPLATE_FILE_NAME}' não foi encontrado na mesma pasta do aplicativo.")
    st.info("Por favor, coloque o arquivo `TEMPLATE_WHATS_COBRANCA.xlsx` ao lado do `app_populador_template_simples.py`.")
else:
    st.sidebar.header("1. Carregar Arquivo de Entrada")
    uploaded_file = st.sidebar.file_uploader(
        "Selecione um arquivo (Excel ou Parquet)", type=["xlsx", "xls", "parquet"]
    )

    sheets_data = None
    if uploaded_file:
        # Passa os bytes e o nome do arquivo para a função cacheada
        sheets_data = load_data_from_file(uploaded_file.getvalue(), uploaded_file.name)
        if sheets_data:
            st.sidebar.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")
            st.sidebar.write(f"Abas/Datasets encontrados: {', '.join(sheets_data.keys())}")

            st.sidebar.header("2. Selecionar Abas/Datasets para Processar")
            selected_sheets = st.sidebar.multiselect(
                "Escolha as abas/datasets que deseja processar:",
                options=list(sheets_data.keys()),
                default=list(sheets_data.keys()) # Seleciona todas por padrão
            )

            if selected_sheets:
                st.sidebar.header("3. Mapeamento de Colunas")
                st.sidebar.info("Mapeie as colunas do seu arquivo de entrada para as colunas do template.")

                # Pega as colunas da primeira aba selecionada para sugerir no mapeamento
                first_selected_df_cols = [''] # Opção vazia para "não mapear"
                if selected_sheets and sheets_data:
                    first_selected_df_cols.extend(sheets_data[selected_sheets[0]].columns.tolist())

                # Colunas do template
                template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

                # Dicionário para armazenar o mapeamento
                column_mapping = {}

                for t_col in template_cols:
                    # Para todas as colunas, permite mapear para uma coluna do input ou usar valor fixo
                    mapping_type = st.sidebar.radio(
                        f"Como preencher '{t_col}'?",
                        ('Mapear Coluna', 'Valor Fixo'),
                        key=f"type_map_{t_col}",
                        index=0 if t_col not in ['CONCESSIONARIA', 'DIRETORIA'] else 1 # Sugere mapear para a maioria, fixo para concessionaria/diretoria
                    )
                    if mapping_type == 'Mapear Coluna':
                        selected_source_col = st.sidebar.selectbox(
                            f"Coluna do seu arquivo para '{t_col}':",
                            options=first_selected_df_cols,
                            key=f"map_{t_col}"
                        )
                        if selected_source_col:
                            column_mapping[t_col] = selected_source_col
                    else: # Valor Fixo
                        fixed_value = st.sidebar.text_input(
                            f"Valor fixo para '{t_col}':",
                            value="" if t_col not in ['CONCESSIONARIA', 'DIRETORIA'] else f"Minha {t_col}",
                            key=f"fixed_val_{t_col}"
                        )
                        column_mapping[t_col] = fixed_value # Armazena o valor fixo

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

                            # Popula o template com os dados da aba usando o mapeamento
                            output_file_bytes = populate_template(df_aba, TEMPLATE_PATH, column_mapping)

                            if output_file_bytes:
                                st.session_state["processed_files"][sheet_name] = output_file_bytes

                            progress_bar.progress((i + 1) / len(selected_sheets))

                        status_text.success("Processamento concluído para todas as abas/datasets selecionados!")
                        st.rerun()

    if "processed_files" in st.session_state and st.session_state["processed_files"]:
        st.subheader("Arquivos Gerados")
        st.markdown("Clique para baixar os arquivos Excel gerados por aba/dataset:")
        for sheet_name, file_bytes in st.session_state["processed_files"].items():
            st.download_button(
                label=f"⬇️ Baixar {sheet_name}.xlsx",
                data=file_bytes,
                file_name=f"{sheet_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{sheet_name}"
            )

    # Opção para limpar o cache
    st.sidebar.markdown("---")
    st.sidebar.header("Opções de Cache")
    if st.sidebar.button("Limpar Cache de Dados Carregados"):
        st.cache_data.clear()
        st.sidebar.success("Cache de dados limpo! O aplicativo será recarregado para garantir novos dados.")
        st.rerun()
