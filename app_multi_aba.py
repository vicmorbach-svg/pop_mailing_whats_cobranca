import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Processador de Abas Excel",
    page_icon="📂",
    layout="wide",
)

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
def populate_template(df_data: pd.DataFrame, template_bytes: bytes) -> bytes:
    try:
        # Carrega o template
        template_wb = load_workbook(io.BytesIO(template_bytes))
        ws = template_wb.active # Assume que o template tem uma aba ativa para preencher

        # Adiciona o cabeçalho do DataFrame se não existir
        # (O template tem cabeçalho fixo, então vamos apenas garantir que as colunas existam)
        template_headers = [cell.value for cell in ws[1]]

        # Mapeia as colunas do DataFrame para as colunas do template
        # Assumimos que as colunas do template são 'MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO'
        # Você precisará ajustar essa lógica se as colunas do seu df_data tiverem nomes diferentes
        # e precisar de um mapeamento explícito. Por enquanto, faremos um mapeamento direto.

        # Exemplo de mapeamento (ajuste conforme suas colunas de entrada)
        # Se o seu df_data tiver colunas como 'cliente_id', 'contato', 'local', etc.
        # você precisará mapear para 'MATRICULA', 'TELEFONE', 'CIDADE', etc.

        # Para este exemplo, vamos assumir que df_data já tem as colunas do template
        # ou que você vai passar um df_data já transformado.
        # Se não, você precisará adicionar uma etapa de transformação aqui.

        # Vamos usar um mapeamento simples para o template fornecido:
        # MATRICULA -> Cliente (ou alguma coluna de ID)
        # TELEFONE -> Alguma coluna de telefone
        # CIDADE -> Localidade
        # SITUACAO -> Tipo de Serviço (ou alguma coluna de status)

        # Para o propósito deste app, vamos assumir que o df_data já está
        # com as colunas renomeadas para MATRICULA, TELEFONE, etc.
        # OU que você quer que o app preencha as colunas do template com
        # as colunas do df_data que tiverem nomes correspondentes.

        # Para o template "TEMPLATE_WHATS_COBRANA_.xlsx":
        # | MATRICULA | TELEFONE | CONCESSIONARIA | CIDADE | DIRETORIA | SITUACAO |

        # Vamos criar um DataFrame de exemplo para o template
        # Você precisará adaptar isso para as colunas reais do seu df_data

        # Exemplo: se df_data tem 'CLIENTE', 'TELEFONE_CONTATO', 'LOCALIDADE', 'STATUS_OS'
        # df_template_ready = df_data.rename(columns={
        #     'CLIENTE': 'MATRICULA',
        #     'TELEFONE_CONTATO': 'TELEFONE',
        #     'LOCALIDADE': 'CIDADE',
        #     'STATUS_OS': 'SITUACAO'
        # })
        # df_template_ready['CONCESSIONARIA'] = 'Sua Concessionaria' # Exemplo de coluna fixa
        # df_template_ready['DIRETORIA'] = 'Sua Diretoria' # Exemplo de coluna fixa

        # Para simplificar, vamos assumir que o df_data já tem as colunas
        # que queremos no template, ou que faremos um mapeamento básico.

        # Colunas do template
        template_cols = ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']

        # Cria um DataFrame com as colunas do template, preenchendo com o que tiver no df_data
        # e NaN para o que não tiver.
        df_to_write = pd.DataFrame(columns=template_cols)
        for col in template_cols:
            if col in df_data.columns:
                df_to_write[col] = df_data[col]
            else:
                df_to_write[col] = '' # Preenche com vazio se a coluna não existir no df_data

        # Escreve os dados no template a partir da segunda linha
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
# Esta função é a mesma que te enviei na última correção, com o grupo_counter global
# e as validações de cliente/serviço não nulos.
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

    # ── Remove linhas sem data, sem cliente ou sem serviço ───────────────────
    work = work.dropna(subset=["_data_parsed"]).copy()
    work = work[
        work[col_cliente].notna() & (work[col_cliente] != '') &
        work[col_servico].notna() & (work[col_servico] != '')
    ].copy()

    if work.empty:
        st.warning("Após a limpeza de dados inválidos (data, cliente ou serviço), não restaram registros para análise.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    work = work.reset_index(drop=True)
    work["_row_id"]       = work.index
    work["_cliente_norm"] = work[col_cliente].astype(str).str.strip().str.upper()
    work["_servico_norm"] = work[col_servico].astype(str).str.strip().str.upper()

    # ── Algoritmo de detecção ─────────────────────────────────────────────────
    registros = []
    global_grupo_counter = 0 # Contador GLOBAL para IDs de grupo únicos

    for (_, _), grp in work.groupby(["_cliente_norm", "_servico_norm"], sort=False):
        if len(grp) < 2:
            continue

        grp     = grp.sort_values("_data_parsed").reset_index(drop=True)
        datas   = grp["_data_parsed"].tolist()
        row_ids = grp["_row_id"].tolist()
        n       = len(grp)

        classificacao = {} # row_id -> {"grupo": int, "tipo": str}
        i = 0

        while i < n:
            rid_i = row_ids[i]

            # Se esta OS já foi classificada como DUPLICATA, não pode ser âncora de um novo grupo
            if rid_i in classificacao and classificacao[rid_i]["tipo"] == "DUPLICATA":
                i += 1
                continue

            duplicatas_j = []
            for j in range(i + 1, n):
                delta = (datas[j] - datas[i]).days
                if delta <= janela_dias:
                    rid_j = row_ids[j]
                    # Só adiciona se ainda não foi classificada
                    if rid_j not in classificacao:
                        duplicatas_j.append(j)
                else:
                    break

            if duplicatas_j:
                global_grupo_counter += 1 # Incrementa o contador GLOBAL

                # Classifica a OS âncora como ORIGINAL
                if rid_i not in classificacao:
                    classificacao[rid_i] = {"grupo": global_grupo_counter, "tipo": "ORIGINAL"}

                # Classifica as OSs encontradas como DUPLICATA
                for j_idx in duplicatas_j:
                    classificacao[row_ids[j_idx]] = {"grupo": global_grupo_counter, "tipo": "DUPLICATA"}

            i += 1 # Avança para a próxima OS, mesmo que tenha formado um cluster

        # Adiciona os resultados deste grupo (cliente, serviço) à lista global
        for rid, info in classificacao.items():
            registros.append({
                "_row_id": rid,
                "grupo_duplicidade": info["grupo"],
                "tipo_registro": info["tipo"],
            })

    if not registros:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_classificacao = pd.DataFrame(registros)

    # ── Monta o DataFrame de detalhamento ────────────────────────────────────
    df_merged = pd.merge(
        work, df_classificacao, on="_row_id", how="left"
    )
    # Preenche as OSs que não foram duplicadas
    df_merged["tipo_registro"] = df_merged["tipo_registro"].fillna("ÚNICA")
    df_merged["grupo_duplicidade"] = df_merged["grupo_duplicidade"].fillna(0).astype(int)

    # Filtra apenas os registros que fazem parte de um grupo de duplicidade
    df_merged = df_merged[df_merged["grupo_duplicidade"] > 0].copy()

    # Colunas para o relatório de detalhamento
    output_cols = ["grupo_duplicidade", "tipo_registro"]
    if col_os:
        output_cols.append(col_os)
    output_cols += [col_cliente, col_servico, col_data]
    if cols_extras:
        output_cols += [c for c in cols_extras if c not in output_cols]

    # Garante que as colunas selecionadas existam no df_merged
    cols_existentes = [c for c in output_cols if c in df_merged.columns]
    df_det = df_merged[cols_existentes].copy()

    # Adiciona a coluna de data formatada para exibição
    df_det["DATA_DA_OS"] = df_merged["_data_parsed"].dt.strftime("%d/%m/%Y")

    # Renomeia colunas para o relatório final
    df_det = df_det.rename(columns={
        col_cliente: "CLIENTE",
        col_servico: "DESCRICAO_DO_SERVICO",
        col_os: "Nº DO PEDIDO" if col_os else "Nº DA OS", # Ajuste o nome da coluna OS
        col_data: "DATA_ORIGINAL_INPUT" # Mantém a coluna original de data para referência
    })

    # Reordena as colunas para o detalhamento
    final_det_cols = ["grupo_duplicidade", "tipo_registro", "Nº DO PEDIDO", "DESCRICAO_DO_SERVICO", "DATA_DA_OS", "CLIENTE"]
    if cols_extras:
        final_det_cols.extend([c for c in cols_extras if c not in final_det_cols])

    df_det = df_det[[c for c in final_det_cols if c in df_det.columns]]


    # ── Resumo por grupo ──────────────────────────────────────────────────────
    resumo_grupos = []
    for gid, grp in df_merged.groupby("grupo_duplicidade"):
        orig = grp[grp["tipo_registro"] == "ORIGINAL"]
        dups = grp[grp["tipo_registro"] == "DUPLICATA"]
        if orig.empty: # Deve sempre ter uma original se o grupo > 0
            continue

        o         = orig.iloc[0]
        datas_dup = dups["_data_parsed"]

        row = {
            "grupo":                   int(gid),
            "cliente":                 o[col_cliente],
            "tipo_servico":            o[col_servico],
            "data_os_original":        o["_data_parsed"].strftime("%d/%m/%Y"),
            "qtd_duplicatas":          len(dups),
            "data_primeira_duplicata": datas_dup.min().strftime("%d/%m/%Y") if not dups.empty else "—",
            "data_ultima_duplicata":   datas_dup.max().strftime("%d/%m/%Y") if not dups.empty else "—",
            "intervalo_dias":          (datas_dup.max() - datas_dup.min()).days if not dups.empty else 0,
        }
        if col_os:
            row["os_original"]   = str(o[col_os])
            row["os_duplicadas"] = ", ".join(dups[col_os].astype(str).tolist()) if not dups.empty else "—"
        resumo_grupos.append(row)

    df_resumo_grupos = pd.DataFrame(resumo_grupos)

    # ── Resumo por tipo de serviço ────────────────────────────────────────────
    cont_cli_serv = (
        work.groupby(["_servico_norm", "_cliente_norm"])
        .size()
        .reset_index(name="total_os_cliente")
    )
    resumo_servico = []
    dups_only = df_merged[df_merged["tipo_registro"] == "DUPLICATA"]

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
            "clientes_mais_10_pedidos":     int((dist > 10).sum()),
        })

    df_resumo_servico = pd.DataFrame(resumo_servico).sort_values(
        "total_duplicatas", ascending=False
    ).reset_index(drop=True)

    return df_det, df_resumo_grupos, df_resumo_servico


# ── Exportação Excel ──────────────────────────────────────────────────────────
def to_excel_bytes(
    df_det: pd.DataFrame,
    df_grupos: pd.DataFrame,
    df_servico: pd.DataFrame,
    janela_dias: int,
    sheet_name: str = "Configurações"
) -> bytes:
    wb = Workbook()

    header_fill   = PatternFill("solid", fgColor="1F4E79")
    header_font   = Font(bold=True, color="FFFFFF", size=11)
    header_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin          = Side(style="thin", color="BFBFBF")
    border        = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_original = PatternFill("solid", fgColor="C6EFCE")
    palette       = ["FFF2CC", "FDEBD0", "D5F5E3", "D6EAF8", "F9EBEA",
                     "EAF2FF", "FDF2F8", "E8F8F5", "FDFEFE", "F4ECF7"]

    def auto_col_width(ws, df):
        for c_idx, col in enumerate(df.columns, 1):
            max_length = 0
            column = df.iloc[:, c_idx - 1]
            for cell in column:
                try:
                    if len(str(cell)) > max_length:
                        max_length = len(str(cell))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[get_column_letter(c_idx)].width = adjusted_width

    # Detalhamento por OS
    ws1 = wb.active
    ws1.title = "Detalhamento por OS"
    if not df_det.empty:
        ws1.append(df_det.columns.tolist())
        for r_idx, row in enumerate(ws1.iter_rows(min_row=1, max_row=1), start=1):
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border

        for r_idx, row_data in enumerate(df_det.itertuples(index=False), start=2):
            ws1.append(row_data)
            # Aplica formatação condicional para ORIGINAL vs DUPLICATA
            if row_data[1] == "ORIGINAL": # tipo_registro é a segunda coluna
                for cell in ws1[r_idx]:
                    cell.fill = fill_original
                    cell.border = border
            elif row_data[0] > 0: # Se for duplicata e tiver grupo
                fill_dup = PatternFill("solid", fgColor=palette[(row_data[0] - 1) % len(palette)])
                for cell in ws1[r_idx]:
                    cell.fill = fill_dup
                    cell.border = border
        auto_col_width(ws1, df_det)

    # Resumo por Grupo
    ws2 = wb.create_sheet("Resumo por Grupo")
    if not df_grupos.empty:
        ws2.append(df_grupos.columns.tolist())
        for r_idx, row in enumerate(ws2.iter_rows(min_row=1, max_row=1), start=1):
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border
        for row_data in df_grupos.itertuples(index=False):
            ws2.append(row_data)
        auto_col_width(ws2, df_grupos)

    # Visão por Tipo de Serviço
    ws3 = wb.create_sheet("Visao por Tipo de Servico")
    if not df_servico.empty:
        ws3.append(df_servico.columns.tolist())
        for r_idx, row in enumerate(ws3.iter_rows(min_row=1, max_row=1), start=1):
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border
        for row_data in df_servico.itertuples(index=False):
            ws3.append(row_data)
        auto_col_width(ws3, df_servico)

    # Configurações da Análise
    ws4 = wb.create_sheet(sheet_name)
    meta = [
        ("Data da Análise", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
        ("Janela de Duplicidade", f"{janela_dias} dias"),
        ("Total de OS no período", f"{len(df_det):,}"),
        ("Total de Grupos Duplicados", f"{df_grupos['grupo'].nunique():,}"),
        ("Total de OS Duplicadas", f"{len(df_det[df_det['tipo_registro'] == 'DUPLICATA']):,}"),
    ]
    for r, (k, v) in enumerate(meta, 1):
        ws4.cell(r, 1, k).font = Font(bold=True)
        ws4.cell(r, 2, str(v))
    ws4.column_dimensions['A'].width = 25
    ws4.column_dimensions['B'].width = 30

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Interface Streamlit ───────────────────────────────────────────────────────
st.title("📂 Processador de Abas Excel para Cobrança")
st.markdown(
    "Faça upload de um arquivo Excel com múltiplas abas. "
    "Para cada aba, um arquivo Excel será gerado usando o template padrão."
)

# Upload do template
st.sidebar.header("Template de Saída")
uploaded_template = st.sidebar.file_uploader(
    "Faça upload do TEMPLATE_WHATS_COBRANA_.xlsx",
    type=["xlsx"],
    key="template_uploader"
)

template_content = None
if uploaded_template:
    template_content = uploaded_template.read()
    st.sidebar.success("Template carregado com sucesso!")
else:
    st.sidebar.warning("Por favor, carregue o template Excel para continuar.")

st.sidebar.divider()

# Upload do arquivo com múltiplas abas
st.sidebar.header("Arquivo de Entrada")
uploaded_file = st.sidebar.file_uploader(
    "Faça upload do arquivo Excel com as abas para processar",
    type=["xlsx", "xls"],
    key="main_file_uploader"
)

sheets_data = None
if uploaded_file and template_content:
    sheets_data = load_excel_with_sheets(uploaded_file)
    if sheets_data:
        st.sidebar.success(f"Arquivo '{uploaded_file.name}' carregado com {len(sheets_data)} abas.")
        st.session_state["sheets_data"] = sheets_data
    else:
        st.sidebar.error("Não foi possível carregar os dados das abas.")

if "sheets_data" in st.session_state and st.session_state["sheets_data"] and template_content:
    st.subheader("Abas Encontradas e Configurações de Análise")

    # Seleção de abas para processar
    selected_sheets = st.multiselect(
        "Selecione as abas para processar",
        list(st.session_state["sheets_data"].keys()),
        default=list(st.session_state["sheets_data"].keys())
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
                    df_cobranca_template['MATRICULA'] = df_cobranca['CLIENTE'] # Assumindo que 'CLIENTE' do df_det mapeia para 'MATRICULA'
                    df_cobranca_template['TELEFONE'] = '' # Preencher com a coluna de telefone real do seu df_aba
                    df_cobranca_template['CONCESSIONARIA'] = 'SUA_CONCESSIONARIA' # Valor fixo ou de outra coluna
                    df_cobranca_template['CIDADE'] = df_cobranca['LOCALIDADE'] if 'LOCALIDADE' in df_cobranca.columns else '' # Exemplo
                    df_cobranca_template['DIRETORIA'] = 'SUA_DIRETORIA' # Valor fixo ou de outra coluna
                    df_cobranca_template['SITUACAO'] = df_cobranca['DESCRICAO_DO_SERVICO'] # Exemplo

                    # Se houver outras colunas no template que não estão no df_cobranca, adicione-as vazias
                    for t_col in ['MATRICULA', 'TELEFONE', 'CONCESSIONARIA', 'CIDADE', 'DIRETORIA', 'SITUACAO']:
                        if t_col not in df_cobranca_template.columns:
                            df_cobranca_template[t_col] = ''

                # 2. Popular o template com os dados da aba (apenas duplicatas)
                if not df_cobranca_template.empty:
                    output_file_bytes = populate_template(df_cobranca_template, template_content)
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
