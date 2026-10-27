"""
Página de Gestão e Ingestão de Dados (DuckDB)
Permite upload, validação de schema, ingestão em alta performance e histórico de cargas.
"""

import os
import sys
import datetime
import streamlit as st
import pandas as pd

# Adiciona a raiz do projeto ao sys.path para importar módulos
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from modules.db_manager import (
    DB_PATH,
    init_db,
    ingest_sheet,
    run_query,
    get_last_ingestion_info,
    get_tables_summary,
    get_table_row_count,
    table_exists,
    reset_database
)

st.set_page_config(
    page_title="Ingestão de Dados - DuckDB",
    page_icon="⚡",
    layout="wide"
)

# Inicializa o banco de dados
init_db()

# Cabeçalho Principal
st.markdown("""
<div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); padding: 22px 28px; border-radius: 12px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.15);">
    <h1 style="color: #FFFFFF; font-size: 1.8rem; font-weight: 700; margin: 0 0 6px 0;">
        ⚡ Ingestão e Banco Analítico Local (DuckDB)
    </h1>
    <p style="color: #94A3B8; font-size: 0.95rem; margin: 0;">
        Carregamento de planilhas de alto desempenho, padronização automática e persistência em banco colunar local.
    </p>
</div>
""", unsafe_allow_html=True)

# Métricas do Banco DuckDB no topo
col_m1, col_m2, col_m3, col_m4 = st.columns(4)

tamanho_mb = 0.0
if os.path.exists(DB_PATH):
    tamanho_mb = round(os.path.getsize(DB_PATH) / (1024 * 1024), 2)

df_resumo_tabelas = get_tables_summary()
total_tabelas = len(df_resumo_tabelas)
total_linhas = int(df_resumo_tabelas["Linhas"].sum()) if not df_resumo_tabelas.empty else 0

with col_m1:
    st.metric("📦 Tamanho do Banco (DuckDB)", f"{tamanho_mb} MB")
with col_m2:
    st.metric("🗄️ Tabelas Ativas", f"{total_tabelas}")
with col_m3:
    st.metric("📄 Total de Linhas Gravadas", f"{total_linhas:,}".replace(",", "."))
with col_m4:
    status_banco = "🟢 Ativo e Persistido" if os.path.exists(DB_PATH) else "🟡 Não Inicializado"
    st.metric("Status do Motor", status_banco)

st.markdown("---")

# Definição dos Schemas Esperados para cada base
BASES_CONFIG = {
    "acompanhamento": {
        "titulo": "1. Relatório de Pedidos / Acompanhamento",
        "tabela": "acompanhamento",
        "descricao": "Base principal de pedidos, status, itens de separação e datas de inserção.",
        "colunas_chave": ["Pedido", "Nota", "Dt_inser", "cidade", "Status", "Tipo"]
    },
    "dados_bi": {
        "titulo": "2. Base de Faturamento / Dados BI",
        "tabela": "dados_bi",
        "descricao": "Itens faturados, vendas totais, quantidades e histórico de notas fiscais emitidas.",
        "colunas_chave": ["Cod", "Descr. produto", "dt_nota_hist", "QTD", "Preço Venda", "Vendas Total", "n_nota_obs"]
    },
    "unidades": {
        "titulo": "3. Controle de Emissão / Relatório de Unidades",
        "tabela": "unidades",
        "descricao": "Relatório de pedidos do portal com a associação de cada pedido à sua unidade real.",
        "colunas_chave": ["Codigo pedido", "Unidade", "Data", "Descricao", "Quantidade"]
    },
    "listagem_unidades": {
        "titulo": "4. Listagem Mestre de Unidades (Cadastro)",
        "tabela": "listagem_unidades",
        "descricao": "Cadastro oficial completo de todas as unidades ativas para o cruzamento trimestral.",
        "colunas_chave": ["Nome", "Ativo", "Data de Cadastro", "Codigo Integracao Cliente"]
    },
    "nf_emergencial": {
        "titulo": "5. Base Emergencial / NFs Pré-Corte",
        "tabela": "nf_emergencial",
        "descricao": "Planilha com pedidos e notas fiscais faturadas antes do corte com a coluna 'Local de entrega'.",
        "colunas_chave": ["Local de entrega", "NF"]
    }
}

# Seção de Upload e Carga
st.subheader("📤 Upload e Ingestão de Planilhas")

tab_upload, tab_historico, tab_consulta_sql = st.tabs([
    "📥 Ingestão de Novas Planilhas", 
    "📜 Histórico de Cargas & Metadados", 
    "🔍 Explorador SQL (DuckDB)"
])

with tab_upload:
    col_cfg, col_up = st.columns([1, 1.4])
    
    with col_cfg:
        base_selecionada = st.selectbox(
            "Selecione a Tabela de Destino:",
            options=list(BASES_CONFIG.keys()),
            format_func=lambda k: BASES_CONFIG[k]["titulo"]
        )
        
        cfg = BASES_CONFIG[base_selecionada]
        st.info(f"**Finalidade:** {cfg['descricao']}")
        
        modo_carga = st.radio(
            "Modo de Ingestão:",
            options=["replace", "append"],
            format_func=lambda m: "🔄 Substituir base inteira (replace)" if m == "replace" else "➕ Adicionar / Incrementar registros (append)",
            help="Substituir remove os dados antigos e recria a tabela. Adicionar insere novos registros preservando o histórico."
        )

        st.markdown("**Colunas Chave Esperadas:**")
        cols_html = " ".join([f"<span style='background:#E2E8F0; color:#1E293B; padding:3px 8px; border-radius:6px; font-size:12px; margin-right:4px; display:inline-block; margin-bottom:4px;'>{c}</span>" for c in cfg["colunas_chave"]])
        st.markdown(cols_html, unsafe_allow_html=True)

    with col_up:
        arquivo_enviado = st.file_uploader(
            f"Selecione o arquivo para '{cfg['titulo']}'",
            type=["xlsx", "xls", "csv"],
            key=f"uploader_{base_selecionada}"
        )

        if arquivo_enviado is not None:
            st.markdown(f"**Arquivo selecionado:** `{arquivo_enviado.name}` ({round(arquivo_enviado.size / 1024, 1)} KB)")
            
            # Botão de disparo da ingestão
            if st.button("⚡ Processar e Gravar no DuckDB", type="primary", use_container_width=True):
                with st.spinner("Processando dados e gravando na tabela DuckDB..."):
                    try:
                        inicio_tempo = datetime.datetime.now()
                        resultado = ingest_sheet(
                            file_path_or_buffer=arquivo_enviado,
                            table_name=cfg["tabela"],
                            mode=modo_carga,
                            original_filename=arquivo_enviado.name
                        )
                        tempo_gasto = (datetime.datetime.now() - inicio_tempo).total_seconds()
                        
                        st.success(f"""
                        ✅ **Carga concluída com sucesso em {tempo_gasto:.2f} segundos!**
                        - **Tabela no DuckDB:** `{resultado['table_name']}`
                        - **Linhas inseridas:** **{resultado['row_count']:,}**
                        - **Colunas processadas:** {resultado['columns_count']}
                        - **Modo:** `{resultado['mode']}`
                        - **Timestamp:** {resultado['timestamp']}
                        """.replace(",", "."))
                        
                        # Limpa cache de dados do Streamlit para refletir imediatamente
                        st.cache_data.clear()
                        
                    except Exception as e:
                        st.error(f"❌ Erro durante o processo de ingestão: {e}")

    # Exibição do resumo das tabelas logo abaixo
    st.markdown("---")
    st.markdown("#### 📊 Estado Atual das Tabelas no DuckDB")
    if not df_resumo_tabelas.empty:
        st.dataframe(df_resumo_tabelas, use_container_width=True, hide_index=True)
    else:
        st.warning("Nenhuma tabela foi criada ainda no DuckDB.")


with tab_historico:
    st.subheader("📜 Histórico de Cargas & Metadados")
    st.caption("Logs de todas as operações de ingestão registradas na tabela `ingestion_metadata`.")
    
    df_historico = get_last_ingestion_info()
    if not df_historico.empty:
        st.dataframe(df_historico, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma carga foi registrada no histórico até o momento.")

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        if st.button("🔄 Atualizar Histórico", use_container_width=True):
            st.rerun()
    with col_btn2:
        if st.button("🗑️ Resetar / Limpar Todo o Banco DuckDB", type="secondary", use_container_width=True):
            if reset_database():
                st.cache_data.clear()
                st.success("✅ Banco DuckDB resetado com sucesso!")
                st.rerun()


with tab_consulta_sql:
    st.subheader("🔍 Explorador SQL (DuckDB)")
    st.caption("Execute consultas analíticas diretamente no motor DuckDB e veja os resultados em tempo real.")

    tabelas_existentes = df_resumo_tabelas["Tabela"].tolist() if not df_resumo_tabelas.empty else []
    
    query_exemplo = "SELECT table_name, estimated_size FROM duckdb_tables();"
    if tabelas_existentes:
        t_padrao = tabelas_existentes[0]
        query_exemplo = f"SELECT * FROM {t_padrao} LIMIT 10;"

    sql_input = st.text_area(
        "Digite sua query SQL:",
        value=query_exemplo,
        height=120
    )

    if st.button("▶️ Executar Query SQL", type="primary"):
        with st.spinner("Executando consulta no DuckDB..."):
            try:
                inicio_q = datetime.datetime.now()
                df_sql_res = run_query(sql_input)
                tempo_q = (datetime.datetime.now() - inicio_q).total_seconds()
                
                st.success(f"Query retornou {len(df_sql_res)} linhas em {tempo_q:.3f} segundos.")
                st.dataframe(df_sql_res, use_container_width=True)
            except Exception as e_sql:
                st.error(f"Erro na execução da query SQL: {e_sql}")
