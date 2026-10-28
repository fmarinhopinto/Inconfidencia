"""
========================================================================================
Dashboard Comercial Executivo - Grupo Fortpel & Inconfidência
Arquitetura Modular de Alta Performance | Motor Analítico DuckDB Colunar C++
========================================================================================
"""

import os
import io
import datetime
from PIL import Image
import streamlit as st
import pandas as pd

# Módulos Especializados da Arquitetura
from modules.db_manager import (
    DB_PATH,
    init_db,
    table_exists,
    get_table_row_count,
    ingest_sheet,
    ingest_default_files_if_needed,
    reset_database,
    run_query
)
from modules.ui_components import (
    inject_custom_css,
    render_header,
    render_kpi_cards,
    format_currency_brl,
    format_date_br,
    exibir_tela_login,
    CREDENCIAIS
)
from modules.data_service import (
    obter_base_consolidada,
    processar_aba_trimestral,
    load_uploaded_file,
    DATA_CORTE,
    limpar_texto_corrompido
)

# ---------------------------------------------------------
# 1. INICIALIZAÇÃO ANALÍTICA & CONFIGURAÇÃO DA PÁGINA
# ---------------------------------------------------------
try:
    init_db()
    ingest_default_files_if_needed()
except Exception:
    pass

st.set_page_config(
    layout="wide",
    page_title="Dashboard Comercial - Grupo Fortpel & Inconfidência",
    page_icon="🏢",
    initial_sidebar_state="collapsed"
)

# Injeção da Identidade Visual Corporativa (Design System)
inject_custom_css()

# ---------------------------------------------------------
# 2. CONTROLE DE AUTENTICAÇÃO E PERFIS DE ACESSO
# ---------------------------------------------------------
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_conectado" not in st.session_state:
    st.session_state["usuario_conectado"] = ""
if "perfil" not in st.session_state:
    st.session_state["perfil"] = "cliente"

if not st.session_state["autenticado"]:
    exibir_tela_login()
    st.stop()

# ---------------------------------------------------------
# 3. BARRA LATERAL (SIDEBAR): IDENTIDADE, GESTÃO & FILTROS
# ---------------------------------------------------------
col1, col2 = st.sidebar.columns(2)
try:
    path_fortpel = "logo_fortpel.png" if os.path.exists("logo_fortpel.png") else ("logo_fortpel.png.png" if os.path.exists("logo_fortpel.png.png") else None)
    if path_fortpel:
        col1.image(Image.open(path_fortpel), use_container_width=True)
    else:
        col1.markdown("<div style='background:rgba(255,255,255,0.1); border-radius:6px; padding:6px; text-align:center;'><small style='color:#FFFFFF;'><b>[Fortpel]</b></small></div>", unsafe_allow_html=True)

    path_inconfidencia = "logo_inconfidencia.png" if os.path.exists("logo_inconfidencia.png") else ("logo_inconfidencia.png.jfif" if os.path.exists("logo_inconfidencia.png.jfif") else None)
    if path_inconfidencia:
        col2.image(Image.open(path_inconfidencia), use_container_width=True)
    else:
        col2.markdown("<div style='background:rgba(255,255,255,0.1); border-radius:6px; padding:6px; text-align:center;'><small style='color:#FFFFFF;'><b>[Inconfidência]</b></small></div>", unsafe_allow_html=True)
except Exception as e:
    st.sidebar.warning(f"Erro ao carregar logos: {e}")

usuario_chave = st.session_state.get('usuario_conectado', 'cliente')
info_usuario = CREDENCIAIS.get(usuario_chave, {"nome": usuario_chave.capitalize(), "perfil": "cliente"})
nome_exibicao = info_usuario["nome"]
perfil_atual = st.session_state.get('perfil', 'cliente')
badge_perfil = "Administrador" if perfil_atual == "admin" else "Cliente"

if perfil_atual != "admin":
    st.markdown("""
    <style>
        [data-testid="stSidebarNav"] { display: none !important; }
        div[data-testid="stSidebarNav"] { display: none !important; }
        ul[data-testid="stSidebarNavItems"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

st.sidebar.markdown(f"""
<div style="background: rgba(255,255,255,0.12); padding: 6px 10px; border-radius: 6px; margin: 8px 0; text-align: center;">
    <small style="color: #FFFFFF;">👤 <b>{nome_exibicao}</b> ({badge_perfil})</small>
</div>
""", unsafe_allow_html=True)

if perfil_atual == "admin":
    tamanho_duckdb_mb = round(os.path.getsize(DB_PATH) / (1024 * 1024), 2) if os.path.exists(DB_PATH) else 0.0
    total_linhas_db = get_table_row_count("acompanhamento") + get_table_row_count("dados_bi") + get_table_row_count("unidades")
    st.sidebar.markdown(f"""
    <div style="background: rgba(15, 23, 42, 0.45); border: 1px solid rgba(56, 189, 248, 0.3); padding: 6px 10px; border-radius: 6px; margin-bottom: 8px; text-align: center;">
        <small style="color: #38BDF8;">⚡ <b>DuckDB:</b> {total_linhas_db:,} reg ({tamanho_duckdb_mb} MB)</small>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

if st.sidebar.button("🚪 Sair (Logout)", use_container_width=True):
    st.session_state["autenticado"] = False
    st.session_state["usuario_conectado"] = ""
    st.session_state["perfil"] = "cliente"
    st.rerun()

# Uploads Exclusivos de Administrador
upload_unidades = None
upload_pedidos = None
upload_faturamento = None
upload_controle = None
upload_emergencial = None

if perfil_atual == "admin":
    with st.sidebar:
        st.markdown("---")
        st.header("📂 Carregamento de Bases")
        upload_unidades = st.file_uploader("1. Listagem de Unidades (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="up_unid")
        upload_pedidos = st.file_uploader("2. Relatório de Pedidos / Itens (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="up_ped")
        upload_faturamento = st.file_uploader("3. Base de Faturamento / BI (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="up_fat")
        upload_controle = st.file_uploader("4. Controle de Emissão (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="up_ctrl")
        upload_emergencial = st.file_uploader("5. Base Emergencial (NFs Pré-Corte)", type=["xlsx", "xls", "csv"], key="up_emerg")

        if st.button("💾 Salvar Bases na Nuvem & DuckDB", use_container_width=True):
            arquivos_salvos = []
            if upload_unidades is not None:
                ext_unid = os.path.splitext(upload_unidades.name)[1].lower()
                ext_unid = ext_unid if ext_unid in [".xlsx", ".xls", ".csv"] else ".xlsx"
                with open(f"listagem_unidades{ext_unid}", "wb") as f:
                    f.write(upload_unidades.getbuffer())
                ingest_sheet(upload_unidades, table_name="listagem_unidades", mode="replace")
                arquivos_salvos.append("1. Listagem de Unidades")

            if upload_pedidos is not None:
                with open("acompanhamento.xlsx", "wb") as f:
                    f.write(upload_pedidos.getbuffer())
                ingest_sheet(upload_pedidos, table_name="acompanhamento", mode="replace")
                arquivos_salvos.append("2. Relatório de Pedidos / Itens")

            if upload_faturamento is not None:
                with open("dados_bi.xlsx", "wb") as f:
                    f.write(upload_faturamento.getbuffer())
                ingest_sheet(upload_faturamento, table_name="dados_bi", mode="replace")
                arquivos_salvos.append("3. Base de Faturamento / BI")

            if upload_controle is not None:
                with open("unidades.xlsx", "wb") as f:
                    f.write(upload_controle.getbuffer())
                ingest_sheet(upload_controle, table_name="unidades", mode="replace")
                arquivos_salvos.append("4. Controle de Emissão")

            if upload_emergencial is not None:
                ext_nf = os.path.splitext(upload_emergencial.name)[1].lower()
                ext_nf = ext_nf if ext_nf in [".xlsx", ".xls", ".csv"] else ".xlsx"
                with open(f"nf_emergencial{ext_nf}", "wb") as f:
                    f.write(upload_emergencial.getbuffer())
                ingest_sheet(upload_emergencial, table_name="nf_emergencial", mode="replace")
                arquivos_salvos.append("5. Base Emergencial")

            if arquivos_salvos:
                st.cache_data.clear()
                st.sidebar.success(f"✅ Bases gravadas com sucesso: {', '.join(arquivos_salvos)}!")
                st.rerun()
            else:
                st.sidebar.warning("⚠️ Selecione pelo menos uma planilha para salvar.")

        if st.button("🗑️ Limpar / Resetar Base de Dados", use_container_width=True, type="secondary"):
            reset_database()
            for fpath in ['acompanhamento.xlsx', 'dados_bi.xlsx', 'unidades.xlsx', 'listagem_unidades.xlsx', 'listagem_unidades.xls', 'listagem_unidades.csv', 'nf_emergencial.xlsx', 'nf_emergencial.xls', 'nf_emergencial.csv']:
                if os.path.exists(fpath):
                    try: os.remove(fpath)
                    except Exception: pass
            st.cache_data.clear()
            st.sidebar.success("✅ Base de dados e DuckDB limpos com sucesso!")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Período do Relatório")

data_inicio = st.sidebar.date_input("Data Inicial", value=DATA_CORTE, min_value=DATA_CORTE, format="DD/MM/YYYY")
data_fim = st.sidebar.date_input("Data Final", value=None, min_value=DATA_CORTE, format="DD/MM/YYYY")

data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
st.sidebar.markdown(
    f"<div style='margin-top: 8px; margin-bottom: 8px; color: #FFFFFF; font-size: 14px; font-weight: 500; text-align: center;'>"
    f"🔄 Atualizado {data_atual} por Filipe"
    f"</div>", 
    unsafe_allow_html=True
)
st.sidebar.markdown("---")

# ---------------------------------------------------------
# 4. CARREGAMENTO & PIPELINE ANALÍTICO CONSOLIDADO
# ---------------------------------------------------------
file_acomp = upload_pedidos if upload_pedidos is not None else 'acompanhamento.xlsx'
file_bi = upload_faturamento if upload_faturamento is not None else 'dados_bi.xlsx'
file_unid = upload_controle if upload_controle is not None else 'unidades.xlsx'

tem_duckdb = table_exists("acompanhamento") and table_exists("dados_bi") and table_exists("unidades")
tem_arquivos = os.path.exists('acompanhamento.xlsx') and os.path.exists('dados_bi.xlsx') and os.path.exists('unidades.xlsx')

if not (tem_duckdb or tem_arquivos):
    st.warning("⚠️ Os arquivos de banco de dados não foram localizados. Por favor, contate o administrador.")
    st.stop()

# Execução consolidada de alta performance
df_merged, df_bi_raw, df_unidades_unique, col_ped, col_uni, nfs_emergenciais, pedidos_emergenciais = obter_base_consolidada(
    file_acomp, file_bi, file_unid, uploaded_nf_emergencial=upload_emergencial
)

# ---------------------------------------------------------
# 5. FILTROS NA SIDEBAR
# ---------------------------------------------------------
st.sidebar.subheader("🔍 Filtros de Visualização")

lista_unidades = sorted(df_merged["UNIDADE"].dropna().unique().tolist())
selected_unidades = st.sidebar.multiselect("Unidade de Entrega", options=lista_unidades, default=[])

lista_status = sorted(df_merged["Status"].dropna().unique().tolist())
selected_status = st.sidebar.multiselect("Status do Acompanhamento", options=lista_status, default=[])

lista_tipo = sorted(df_merged["Tipo"].dropna().unique().tolist())
selected_tipo = st.sidebar.multiselect("Tipo de Pedido", options=lista_tipo, default=[])

busca_texto = st.sidebar.text_input("Buscar por Pedido ou Nota:", value="").strip()

# Aplicação dos Filtros
df_filtrado = df_merged.copy()

data_ini_efetiva = max(data_inicio, DATA_CORTE) if data_inicio else DATA_CORTE
ts_ini_efetiva = pd.Timestamp(data_ini_efetiva).normalize()
serie_dt = pd.to_datetime(df_filtrado['Dt_inser'], dayfirst=True, errors='coerce').dt.normalize()

cond_inicio = serie_dt >= ts_ini_efetiva
if nfs_emergenciais:
    cond_inicio = cond_inicio | df_filtrado['NF_Clean'].isin(nfs_emergenciais)
if pedidos_emergenciais:
    cond_inicio = cond_inicio | df_filtrado['Pedido_Clean'].isin(pedidos_emergenciais)

if data_fim:
    ts_fim = pd.Timestamp(data_fim).normalize()
    cond_fim = serie_dt <= ts_fim
    if nfs_emergenciais:
        cond_fim = cond_fim | df_filtrado['NF_Clean'].isin(nfs_emergenciais)
    if pedidos_emergenciais:
        cond_fim = cond_fim | df_filtrado['Pedido_Clean'].isin(pedidos_emergenciais)
    df_filtrado = df_filtrado[cond_inicio & cond_fim]
else:
    df_filtrado = df_filtrado[cond_inicio]

if selected_unidades:
    df_filtrado = df_filtrado[df_filtrado["UNIDADE"].isin(selected_unidades)]
if selected_status:
    df_filtrado = df_filtrado[df_filtrado["Status"].isin(selected_status)]
if selected_tipo:
    df_filtrado = df_filtrado[df_filtrado["Tipo"].isin(selected_tipo)]
if busca_texto:
    df_filtrado = df_filtrado[
        df_filtrado["Pedido"].astype(str).str.contains(busca_texto, case=False, na=False) |
        df_filtrado["Nota"].astype(str).str.contains(busca_texto, case=False, na=False)
    ]

# ---------------------------------------------------------
# 6. CABEÇALHO DO DASHBOARD E CARDS DE KPIS EXECUTIVOS
# ---------------------------------------------------------
render_header()

kpi_total_pedidos = len(df_filtrado)
kpi_faturados = int((df_filtrado["Status_Faturamento"] == "Faturado").sum())
kpi_pendentes = int((df_filtrado["Status_Faturamento"] != "Faturado").sum())
kpi_valor_total = float(df_filtrado["Vendas Total"].sum())

render_kpi_cards(kpi_total_pedidos, kpi_faturados, kpi_pendentes, kpi_valor_total)

# Painel de Diagnóstico de Cruzamento
with st.expander("🛠️ Debug de Cruzamento - Por que as unidades não aparecem?"):
    st.warning("Diagnóstico do Cruzamento de Pedidos x Unidades")
    st.write(f"- **Coluna lida como 'Pedido' na 3ª planilha:** `{col_ped}`")
    st.write(f"- **Coluna lida como 'Unidade' na 3ª planilha:** `{col_uni}`")
    st.write("- **Amostra de Pedidos Limpos (Acompanhamento):**", df_merged['Pedido_Clean'].head(5).tolist())
    st.write("- **Amostra de Pedidos Limpos (3ª Planilha):**", df_unidades_unique['Pedido_Clean'].head(5).tolist())
    sucessos = df_merged[col_uni].notna().sum()
    st.metric("Pedidos com 'Match' exato", f"{sucessos} de {len(df_merged)}")

st.markdown("<div style='margin-bottom: 25px;'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 7. ABAS PRINCIPAIS: VISÃO POR UNIDADE, PEDIDOS & CONTROLE
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 Visão por Unidade", "📦 Pedidos e Itens", "Controle de Emissão de pedidos"])

# --- ABA 1: VISÃO POR UNIDADE ---
with tab1:
    st.markdown("### Relatório Detalhado de Faturamento")

    df_final = df_filtrado if ('df_filtrado' in locals() and df_filtrado is not None and not df_filtrado.empty) else df_merged
    col_valor_matches = [c for c in df_final.columns if 'vendas total' in str(c).lower() or 'valor total' in str(c).lower()]
    col_valor = col_valor_matches[0] if col_valor_matches else df_final.columns[-1]

    df_tabela = df_final[['Pedido', 'Unidade_Final', 'Dt_inser', col_valor]].copy()
    df_tabela.rename(columns={
        'Unidade_Final': 'Unidade', 
        'Dt_inser': 'Data de Inserção', 
        col_valor: 'Valor (R$)'
    }, inplace=True)

    df_tabela['Data_Formatada'] = pd.to_datetime(df_tabela['Data de Inserção'], errors='coerce', dayfirst=True).dt.normalize()
    ts_tab_ini = pd.Timestamp(data_ini_efetiva).normalize()
    cond_tab = df_tabela['Data_Formatada'] >= ts_tab_ini
    if nfs_emergenciais or pedidos_emergenciais:
        peds_liberados = df_merged['Pedido'].unique()
        cond_tab = cond_tab | df_tabela['Pedido'].isin(peds_liberados)

    if data_fim:
        ts_tab_fim = pd.Timestamp(data_fim).normalize()
        cond_tab_fim = df_tabela['Data_Formatada'] <= ts_tab_fim
        if nfs_emergenciais or pedidos_emergenciais:
            cond_tab_fim = cond_tab_fim | df_tabela['Pedido'].isin(peds_liberados)
        df_tabela = df_tabela[cond_tab & cond_tab_fim]
    else:
        df_tabela = df_tabela[cond_tab]

    df_tabela = df_tabela.drop(columns=['Data_Formatada'], errors='ignore')
    df_tabela['Valor (R$)'] = pd.to_numeric(df_tabela['Valor (R$)'], errors='coerce').fillna(0)
    df_tabela['Valor (R$)'] = df_tabela['Valor (R$)'].apply(format_currency_brl)
    df_tabela['Data de Inserção'] = df_tabela['Data de Inserção'].fillna('-')
    df_tabela['Pedido'] = df_tabela['Pedido'].fillna('-')

    colunas_exibicao = ['Pedido', 'Unidade', 'Data de Inserção', 'Valor (R$)']
    st.dataframe(df_tabela[colunas_exibicao], use_container_width=True, height=600, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_tabela[colunas_exibicao].to_excel(writer, index=False, sheet_name='Relatorio_Filtrado')

    st.download_button(
        label="📥 Baixar Relatório Filtrado (Excel)",
        data=buffer.getvalue(),
        file_name="Relatorio_Comercial_Inconfidencia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    pedidos_orfaos = df_merged[df_merged['Unidade_Final'] == df_merged['cidade']][['Pedido', 'cidade', 'Status', 'Dt_inser']]
    if not pedidos_orfaos.empty:
        with st.expander("⚠️ Atenção: Pedidos sem correspondência no Relatório de Unidades"):
            st.write("Os pedidos abaixo não constam na planilha de Unidades (ou possuem numeração divergente). A cidade foi usada como Unidade provisória para não perder o registro do faturamento:")
            st.dataframe(pedidos_orfaos, use_container_width=True, hide_index=True)

# --- ABA 2: PEDIDOS E ITENS ---
with tab2:
    st.subheader("Relação Detalhada de Pedidos")

    if df_filtrado.empty:
        st.warning("Nenhum pedido encontrado para os filtros selecionados.")
    else:
        colunas_pedidos = ["Pedido", "Nota", "UNIDADE", "Dt_inser", "dt_nota_hist", "Status", "Tipo", "Vendas Total", "Status_Faturamento"]
        df_pedidos_exibicao = df_filtrado[colunas_pedidos].copy()
        df_pedidos_exibicao["Dt_inser"] = df_pedidos_exibicao["Dt_inser"].apply(format_date_br)
        df_pedidos_exibicao["dt_nota_hist"] = df_pedidos_exibicao["dt_nota_hist"].apply(format_date_br)
        df_pedidos_exibicao["Vendas Total"] = df_pedidos_exibicao["Vendas Total"].apply(format_currency_brl)
        df_pedidos_exibicao = df_pedidos_exibicao.rename(columns={
            "Pedido": "Pedido",
            "Nota": "Nota Fiscal",
            "UNIDADE": "Unidade",
            "Dt_inser": "Data Pedido",
            "dt_nota_hist": "Data Faturamento",
            "Status": "Status Acomp.",
            "Tipo": "Tipo",
            "Vendas Total": "Valor Faturado",
            "Status_Faturamento": "Situação"
        })

        st.dataframe(df_pedidos_exibicao, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("🔍 Detalhamento e Drill-Down do Pedido")

        pedidos_disponiveis = df_filtrado["Pedido"].astype(str).unique().tolist()
        pedido_selecionado = st.selectbox("Selecione um Pedido para visualizar os detalhes e itens:", options=pedidos_disponiveis, index=0)

        if pedido_selecionado:
            pedido_info = df_filtrado[df_filtrado["Pedido"].astype(str) == str(pedido_selecionado)].iloc[0]
            nota_pedido = str(pedido_info.get("Nota", "0")).strip()
            status_faturamento = pedido_info.get("Status_Faturamento", "Pendente")
            valor_faturado = format_currency_brl(pedido_info.get("Vendas Total", 0.0))
            data_pedido = format_date_br(pedido_info.get("Dt_inser", "-"))
            data_fat = format_date_br(pedido_info.get("dt_nota_hist", "-"))
            unidade_pedido = pedido_info.get("UNIDADE", "Não Informada")
            status_acomp = pedido_info.get("Status", "-")
            tipo_pedido = pedido_info.get("Tipo", "-")

            st.markdown(f"""
            <div class="order-card">
                <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E2E8F0; padding-bottom: 12px; margin-bottom: 15px;">
                    <h4 style="margin: 0; color: #1E293B; font-weight: 700;">
                        📄 Pedido: <span style="color:#3B6077;">{pedido_selecionado}</span>
                    </h4>
                    <span style="background-color: {'#E6F4EA' if status_faturamento == 'Faturado' else '#FCE8E6'}; color: {'#137333' if status_faturamento == 'Faturado' else '#C5221F'}; padding: 4px 12px; border-radius: 12px; font-weight: 700; font-size: 0.85rem;">
                        {status_faturamento}
                    </span>
                </div>
                <div style="display: flex; flex-wrap: wrap; gap: 24px;">
                    <div><span style="color:#64748B; font-size:0.85rem;">Nota Fiscal:</span><br><b>{nota_pedido if nota_pedido != '0' else 'Sem Nota'}</b></div>
                    <div><span style="color:#64748B; font-size:0.85rem;">Unidade de Entrega:</span><br><b style="color:#3B6077;">{unidade_pedido}</b></div>
                    <div><span style="color:#64748B; font-size:0.85rem;">Data do Pedido:</span><br><b>{data_pedido}</b></div>
                    <div><span style="color:#64748B; font-size:0.85rem;">Data Faturamento:</span><br><b>{data_fat}</b></div>
                    <div><span style="color:#64748B; font-size:0.85rem;">Status Acompanhamento:</span><br><b>{status_acomp}</b></div>
                    <div><span style="color:#64748B; font-size:0.85rem;">Tipo:</span><br><b>{tipo_pedido}</b></div>
                    <div><span style="color:#64748B; font-size:0.85rem;">Valor Total Faturado:</span><br><b style="color:#E6192B; font-size:1.1rem;">{valor_faturado}</b></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if nota_pedido in ["0", "", "nan", "None"]:
                st.warning("⚠️ **Atenção:** Este pedido ainda aguarda faturamento e não possui nota fiscal vinculada.")
            else:
                itens_nota = df_bi_raw[df_bi_raw["n_nota_obs"].astype(str) == nota_pedido].copy()
                if not itens_nota.empty:
                    st.markdown(f"#### 🛒 Itens Vinculados à Nota Fiscal `{nota_pedido}`")
                    cols_itens = ["Cod", "Descr. produto", "QTD", "Preço Venda", "Vendas Total"]
                    itens_exibicao = itens_nota[cols_itens].copy()
                    itens_exibicao["Preço Venda"] = itens_exibicao["Preço Venda"].apply(format_currency_brl)
                    itens_exibicao["Vendas Total"] = itens_exibicao["Vendas Total"].apply(format_currency_brl)
                    itens_exibicao["QTD"] = itens_exibicao["QTD"].apply(lambda q: f"{q:,.0f}".replace(",", "."))
                    st.dataframe(itens_exibicao, use_container_width=True, hide_index=True)
                else:
                    st.warning(f"⚠️ A Nota Fiscal `{nota_pedido}` não foi encontrada na base de dados do BI.")

# --- ABA 3: CONTROLE DE EMISSÃO DE PEDIDOS ---
with tab3:
    st.markdown("### Controle de Emissão de Pedidos - Visão Trimestral")
    st.caption("Consolidação trimestral automática (últimos 3 meses até a data presente) cruzando a listagem mestre de unidades com os relatórios de pedidos do portal.")

    tem_listagem_db = table_exists("listagem_unidades")
    tem_unidades_db = table_exists("unidades")
    file_listagem_unidades_fixo = 'listagem_unidades.xlsx' if os.path.exists('listagem_unidades.xlsx') else ('listagem_unidades.csv' if os.path.exists('listagem_unidades.csv') else ('listagem_unidades.xls' if os.path.exists('listagem_unidades.xls') else None))
    arquivo_listagem_alvo = upload_unidades if upload_unidades is not None else file_listagem_unidades_fixo
    arquivo_pedidos_alvo = file_unid if file_unid is not None else ('unidades.xlsx' if os.path.exists('unidades.xlsx') else None)

    if not tem_listagem_db and not arquivo_listagem_alvo:
        st.warning("⚠️ Por favor, faça o upload da 'Listagem de Unidades' na área administrativa da barra lateral.")
    elif not tem_unidades_db and not arquivo_pedidos_alvo:
        st.warning("⚠️ O relatório de pedidos do portal ('unidades.xlsx') não foi encontrado no servidor.")
    else:
        try:
            if not (isinstance(arquivo_listagem_alvo, io.BytesIO) or hasattr(arquivo_listagem_alvo, "read")) and tem_listagem_db:
                df_cad_raw = run_query("SELECT * FROM listagem_unidades")
            else:
                df_cad_raw = load_uploaded_file(arquivo_listagem_alvo)

            if not (isinstance(arquivo_pedidos_alvo, io.BytesIO) or hasattr(arquivo_pedidos_alvo, "read")) and tem_unidades_db:
                df_ped_raw = run_query("SELECT * FROM unidades")
            else:
                df_ped_raw = load_uploaded_file(arquivo_pedidos_alvo)

            matriz_exibicao, meses_ordenados, total_cad, total_ativas, unidades_com_compra, unidades_sem_compra = processar_aba_trimestral(
                df_cad_raw, df_ped_raw
            )

            c_kp1, c_kp2, c_kp3, c_kp4 = st.columns(4)
            c_kp1.metric("Unidades Cadastradas", f"{total_cad:,}".replace(",", "."))
            c_kp2.metric("Unidades Ativas", f"{total_ativas:,}".replace(",", "."))
            c_kp3.metric("Com Compra no Período", f"{unidades_com_compra:,}".replace(",", "."))
            c_kp4.metric("Sem Compra no Período", f"{unidades_sem_compra:,}".replace(",", "."))

            st.markdown("---")
            st.markdown("#### 🔍 Filtros de Consulta")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                opcoes_ativo = ["TODOS"] + list(matriz_exibicao['ATIVO'].dropna().unique()) if 'ATIVO' in matriz_exibicao.columns else ["TODOS"]
                filtro_ativo = st.selectbox("Filtrar por Status (ATIVO)", opcoes_ativo)
            with col_f2:
                busca_unidade = st.text_input("🔍 Buscar Unidade por Nome")

            df_controle_f = matriz_exibicao.copy()
            if filtro_ativo != "TODOS" and 'ATIVO' in df_controle_f.columns:
                df_controle_f = df_controle_f[df_controle_f['ATIVO'] == filtro_ativo]

            if busca_unidade and 'Unidade' in df_controle_f.columns:
                busca_norm = limpar_texto_corrompido(busca_unidade)
                df_controle_f = df_controle_f[
                    df_controle_f['Unidade'].apply(limpar_texto_corrompido).str.contains(busca_norm, case=False, na=False) |
                    df_controle_f['Unidade'].astype(str).str.contains(busca_unidade, case=False, na=False)
                ]

            st.markdown(f"Exibindo {len(df_controle_f)} unidades encontradas:")
            st.dataframe(df_controle_f, use_container_width=True, height=600, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)
            buf_matriz = io.BytesIO()
            with pd.ExcelWriter(buf_matriz, engine='xlsxwriter') as writer:
                df_controle_f.to_excel(writer, index=False, sheet_name='Controle_Emissao')

            st.download_button(
                label="📥 Baixar Controle de Emissão Trimestral (Excel)",
                data=buf_matriz.getvalue(),
                file_name="Controle_Emissao_Pedidos_Trimestral.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as err:
            st.error(f"Erro ao processar visão trimestral de emissão de pedidos: {err}")
