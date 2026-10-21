import os
import io
import re
import zipfile
import xlrd
from PIL import Image
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ---------------------------------------------------------
# 1. CONFIGURAÇÃO DA PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    layout="wide",
    page_title="Dashboard Comercial | Grupo Fortpel & Inconfidência",
    page_icon="🏢"
)

# ---------------------------------------------------------
# IDENTIDADE VISUAL CORPORATIVA (CSS PERSONALIZADO)
# Cores Base:
# - Azul Petróleo (Inconfidência): #3B6077
# - Vermelho Fortpel: #E6192B
# - Laranja Detalhes: #F26D21
# - Fundo Geral: #F8F9FA
# ---------------------------------------------------------
st.markdown("""
<style>
    /* 1. FUNDO PRINCIPAL E TIPOGRAFIA */
    [data-testid="stAppViewContainer"], .main {
        background-color: #F8F9FA !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    [data-testid="stHeader"] {
        background-color: rgba(248, 249, 250, 0.85) !important;
    }

    /* 2. BARRA LATERAL (SIDEBAR) - AZUL PETRÓLEO */
    section[data-testid="stSidebar"] {
        background-color: #3B6077 !important;
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] h4, 
    section[data-testid="stSidebar"] p, 
    section[data-testid="stSidebar"] span, 
    section[data-testid="stSidebar"] label {
        color: #FFFFFF !important;
        font-weight: 500;
    }
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.25) !important;
    }

    /* Uploader na Sidebar */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
        background-color: rgba(255, 255, 255, 0.08);
        border: 1px dashed rgba(255, 255, 255, 0.35);
        border-radius: 8px;
        padding: 8px;
        margin-bottom: 12px;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] small {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
        background-color: #E6192B !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button:hover {
        background-color: #C21323 !important;
        box-shadow: 0 2px 6px rgba(230, 25, 43, 0.4);
    }

    /* Widgets na Sidebar (Inputs e Multiselect) */
    section[data-testid="stSidebar"] .stTextInput input {
        background-color: #FFFFFF !important;
        color: #1E293B !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] .stTextInput input:focus {
        border-color: #E6192B !important;
        box-shadow: 0 0 0 2px rgba(230, 25, 43, 0.25) !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        border-color: #CBD5E1 !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="select"]:focus-within > div {
        border-color: #E6192B !important;
        box-shadow: 0 0 0 2px rgba(230, 25, 43, 0.25) !important;
    }
    section[data-testid="stSidebar"] [data-baseweb="tag"] {
        background-color: #3B6077 !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }

    /* 3. CARTÕES DE KPIS (st.metric) */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        padding: 18px 20px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
        border-left: 5px solid #3B6077 !important;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -2px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stMetricLabel"] p {
        font-size: 0.82rem !important;
        font-weight: 700 !important;
        color: #64748B !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.85rem !important;
        font-weight: 800 !important;
        color: #3B6077 !important;
    }

    /* Métrica 3 (Alerta/Pendente): Vermelho Fortpel */
    div[data-testid="column"]:nth-of-type(3) div[data-testid="stMetric"] {
        border-left: 5px solid #E6192B !important;
    }
    div[data-testid="column"]:nth-of-type(3) div[data-testid="stMetricValue"] {
        color: #E6192B !important;
    }

    /* Métrica 4 (Valor Total Faturado): Acento em Laranja Inconfidência */
    div[data-testid="column"]:nth-of-type(4) div[data-testid="stMetric"] {
        border-left: 5px solid #F26D21 !important;
    }
    div[data-testid="column"]:nth-of-type(4) div[data-testid="stMetricValue"] {
        color: #1E293B !important;
    }

    /* 4. ABAS E BOTÕES */
    button[data-baseweb="tab"] {
        color: #3B6077 !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding-top: 10px !important;
        padding-bottom: 10px !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #E6192B !important;
    }
    div[data-baseweb="tab-highlight"] {
        background-color: #E6192B !important;
        height: 3px !important;
    }

    button[kind="primary"], .stButton > button {
        background-color: #E6192B !important;
        color: #FFFFFF !important;
        border-color: #E6192B !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    button[kind="primary"]:hover, .stButton > button:hover {
        background-color: #C21323 !important;
        border-color: #C21323 !important;
    }

    /* 5. CARDS DE DETALHES */
    .order-card {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 22px;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #3B6077;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 25px;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# FUNÇÕES UTILITÁRIAS
# ---------------------------------------------------------
def format_currency_brl(val):
    """Formata valor float para moeda brasileira R$ 1.234,56"""
    try:
        val = float(val)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def format_date_br(val):
    """Formata data para DD/MM/AAAA de maneira segura"""
    if pd.isna(val) or val is None or str(val).strip() in ["", "-", "NaT", "nan", "None"]:
        return "-"
    s_val = str(val).strip()
    if len(s_val) == 10 and s_val[2] == "/" and s_val[5] == "/":
        return s_val
    try:
        dt = pd.to_datetime(val, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return s_val

def find_column(df, target_names, partial=False):
    """Localiza uma coluna de forma insensível a maiúsculas e espaços"""
    cols = {str(c).strip().lower(): c for c in df.columns}
    for target in target_names:
        t = target.strip().lower()
        if t in cols:
            return cols[t]
    if partial:
        for target in target_names:
            t = target.strip().lower()
            for col_lower, original_col in cols.items():
                if t in col_lower:
                    return original_col
    return None

def sanitize_order_id(val):
    """
    Higienização profunda da chave de Pedido:
    - Converte para string
    - Remove '.0' adicionado pelo Excel
    - Remove espaços em branco
    - Extrai os dígitos numéricos para garantir casamento exato de tipos
    """
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "null", "-", ""]:
        return ""
    s = re.sub(r"\.0+$", "", s)
    s = s.replace(".0", "")
    digits = re.sub(r"\D", "", s)
    if digits:
        try:
            return str(int(digits))
        except Exception:
            return digits
    return s.strip()

def load_uploaded_file(file_obj):
    """
    Lê arquivos .xlsx, .xls ou .csv com verificação de extensão e fallback automático para HTML:
    1. Se .csv: pd.read_csv(file_obj, sep=None, engine='python')
    2. Se .xlsx: pd.read_excel(file_obj, engine='openpyxl')
    3. Se .xls: pd.read_excel(file_obj, engine='xlrd')
    4. Em caso de XLRDError, BadZipFile ou ValueError (ex.: arquivos HTML disfarçados de Excel),
       retorna ao início do arquivo (file_obj.seek(0)) e executa o fallback com pd.read_html.
    """
    if isinstance(file_obj, str):
        with open(file_obj, "rb") as f:
            content = f.read()
        buf = io.BytesIO(content)
        buf.name = file_obj
        file_obj = buf

    filename = getattr(file_obj, "name", "").strip().lower()
    file_obj.seek(0)
    
    # 1. Arquivos nativos CSV
    if filename.endswith(".csv"):
        for enc in ["utf-8", "latin1", "cp1252", "iso-8859-1"]:
            try:
                file_obj.seek(0)
                return pd.read_csv(file_obj, sep=None, engine="python", encoding=enc)
            except Exception:
                continue
        file_obj.seek(0)
        return pd.read_csv(file_obj, sep=None, engine="python")

    # 2. Arquivos Excel (.xlsx ou .xls)
    engine = "openpyxl" if filename.endswith(".xlsx") else "xlrd"
    try:
        file_obj.seek(0)
        return pd.read_excel(file_obj, engine=engine, sheet_name=0)
    except (xlrd.biffh.XLRDError, zipfile.BadZipFile, ValueError, Exception) as e:
        # Fallback 1: Ler como tabela HTML disfarçada de Excel (comum em ERPs)
        try:
            file_obj.seek(0)
            tables = pd.read_html(file_obj, decimal=",", thousands=".")
            if tables and len(tables) > 0:
                # Seleciona a tabela com maior número de células (para evitar pegar cabeçalhos secundários de 1 linha)
                df = max(tables, key=lambda t: t.shape[0] * t.shape[1])
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [" ".join([str(c) for c in col if str(c) != "" and not str(c).startswith("Unnamed")]).strip() for col in df.columns]
                return df
        except Exception:
            pass

        # Fallback 2: Tentar decodificar HTML com encodings latinos
        try:
            file_obj.seek(0)
            raw_bytes = file_obj.read()
            for enc in ["latin1", "utf-8", "cp1252", "iso-8859-1"]:
                try:
                    text = raw_bytes.decode(enc)
                    tables = pd.read_html(io.StringIO(text), decimal=",", thousands=".")
                    if tables and len(tables) > 0:
                        df = max(tables, key=lambda t: t.shape[0] * t.shape[1])
                        if isinstance(df.columns, pd.MultiIndex):
                            df.columns = [" ".join([str(c) for c in col if str(c) != "" and not str(c).startswith("Unnamed")]).strip() for col in df.columns]
                        return df
                except Exception:
                    continue
        except Exception:
            pass

        # Fallback 3: Tentar alternar engine de Excel (.xls salvo como .xlsx ou vice-versa)
        alt_engine = "xlrd" if engine == "openpyxl" else "openpyxl"
        try:
            file_obj.seek(0)
            return pd.read_excel(file_obj, engine=alt_engine, sheet_name=0)
        except Exception:
            pass

        # Fallback 4: Tentar como CSV delimitado
        try:
            file_obj.seek(0)
            return pd.read_csv(file_obj, sep=None, engine="python", encoding="latin1")
        except Exception:
            pass

        raise ValueError(f"Não foi possível processar o arquivo '{getattr(file_obj, 'name', '')}': {e}")


# ---------------------------------------------------------
# AUTENTICAÇÃO E CONTROLE DE ACESSO NATIVO COM PERFIS
# ---------------------------------------------------------
CREDENCIAIS = {
    "filipe": {
        "senha": "Filipe123@@",
        "perfil": "admin",
        "nome": "Filipe Marinho"
    },
    "admin": {
        "senha": "Filipe123@@",
        "perfil": "admin",
        "nome": "Administrador"
    },
    "cliente": {
        "senha": "inconfidencia2026",
        "perfil": "cliente",
        "nome": "Cliente Inconfidência"
    },
    "inconfidencia": {
        "senha": "inconfidencia2026",
        "perfil": "cliente",
        "nome": "Cliente Inconfidência"
    }
}

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_conectado" not in st.session_state:
    st.session_state["usuario_conectado"] = ""
if "perfil" not in st.session_state:
    st.session_state["perfil"] = "cliente"

def exibir_tela_login():
    # Oculta a barra lateral durante o login para uma interface limpa e focada
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_vazia1, col_login, col_vazia2 = st.columns([1, 1.2, 1])
    with col_login:
        col_l1, col_l2 = st.columns(2)
        try:
            path_fortpel = "logo_fortpel.png" if os.path.exists("logo_fortpel.png") else ("logo_fortpel.png.png" if os.path.exists("logo_fortpel.png.png") else None)
            if path_fortpel:
                col_l1.image(Image.open(path_fortpel), use_container_width=True)
            path_inconfidencia = "logo_inconfidencia.png" if os.path.exists("logo_inconfidencia.png") else ("logo_inconfidencia.png.jfif" if os.path.exists("logo_inconfidencia.png.jfif") else None)
            if path_inconfidencia:
                col_l2.image(Image.open(path_inconfidencia), use_container_width=True)
        except Exception:
            pass

        st.markdown("""
        <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 24px 20px 8px 20px; box-shadow: 0 4px 14px rgba(0,0,0,0.06); margin-top: 15px; margin-bottom: 20px; text-align: center;">
            <h2 style="font-size: 1.4rem; font-weight: 700; color: #1E293B; margin-bottom: 4px;">
                Portal Comercial
            </h2>
            <p style="font-size: 0.9rem; color: #64748B; margin-bottom: 12px;">
                Digite seu usuário e senha para acessar o sistema.
            </p>
        </div>
        """, unsafe_allow_html=True)

        with st.form("form_login", clear_on_submit=False):
            usuario_input = st.text_input("Usuário", placeholder="Ex: filipe ou cliente").strip().lower()
            senha_input = st.text_input("Senha", type="password", placeholder="Digite sua senha")
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            botao_entrar = st.form_submit_button("Entrar no Sistema", use_container_width=True)

            if botao_entrar:
                if usuario_input in CREDENCIAIS and CREDENCIAIS[usuario_input]["senha"] == senha_input:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_conectado"] = usuario_input
                    st.session_state["perfil"] = CREDENCIAIS[usuario_input]["perfil"]
                    st.success("✅ Acesso autorizado! Carregando dashboard...")
                    st.rerun()
                else:
                    st.error("❌ Usuário ou senha incorretos.")

if not st.session_state["autenticado"]:
    exibir_tela_login()
    st.stop()


# ---------------------------------------------------------
# 2. INTERFACE E UPLOADS NA SIDEBAR
# ---------------------------------------------------------
col1, col2 = st.sidebar.columns(2)

try:
    path_fortpel = "logo_fortpel.png" if os.path.exists("logo_fortpel.png") else ("logo_fortpel.png.png" if os.path.exists("logo_fortpel.png.png") else None)
    if path_fortpel:
        img_fortpel = Image.open(path_fortpel)
        col1.image(img_fortpel, use_container_width=True)
    else:
        col1.markdown("<div style='background:rgba(255,255,255,0.1); border-radius:6px; padding:6px; text-align:center;'><small style='color:#FFFFFF;'><b>[Falta logo_fortpel.png]</b></small></div>", unsafe_allow_html=True)

    path_inconfidencia = "logo_inconfidencia.png" if os.path.exists("logo_inconfidencia.png") else ("logo_inconfidencia.png.jfif" if os.path.exists("logo_inconfidencia.png.jfif") else None)
    if path_inconfidencia:
        img_inconfidencia = Image.open(path_inconfidencia)
        col2.image(img_inconfidencia, use_container_width=True)
    else:
        col2.markdown("<div style='background:rgba(255,255,255,0.1); border-radius:6px; padding:6px; text-align:center;'><small style='color:#FFFFFF;'><b>[Falta logo_inconfidencia.png]</b></small></div>", unsafe_allow_html=True)
except Exception as e:
    st.sidebar.warning(f"Erro ao carregar logos: {e}")

# Card de sessão e botão de Logout
usuario_chave = st.session_state.get('usuario_conectado', 'cliente')
info_usuario = CREDENCIAIS.get(usuario_chave, {"nome": usuario_chave.capitalize(), "perfil": "cliente"})
nome_exibicao = info_usuario["nome"]
perfil_atual = st.session_state.get('perfil', 'cliente')
badge_perfil = "Administrador" if perfil_atual == "admin" else "Cliente"

st.sidebar.markdown(f"""
<div style="background: rgba(255,255,255,0.12); padding: 6px 10px; border-radius: 6px; margin: 8px 0; text-align: center;">
    <small style="color: #FFFFFF;">👤 <b>{nome_exibicao}</b> ({badge_perfil})</small>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("🚪 Sair (Logout)", use_container_width=True):
    st.session_state["autenticado"] = False
    st.session_state["usuario_conectado"] = ""
    st.session_state["perfil"] = "cliente"
    st.rerun()

# ---------------------------------------------------------
# PAINEL EXCLUSIVO DO ADMINISTRADOR (ATUALIZAÇÃO DE BASES)
# ---------------------------------------------------------
if st.session_state.get("perfil") == "admin":
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Gestão de Bases (Admin)")
    st.sidebar.caption("Envie novas planilhas e clique em 'Salvar Bases na Nuvem' para atualizar as bases do sistema.")

    up_acomp = st.sidebar.file_uploader(
        "1. Planilha de Acompanhamento (.xlsx)",
        type=["xlsx"],
        key="admin_up_acomp"
    )
    up_bi = st.sidebar.file_uploader(
        "2. Planilha de Dados BI (.xlsx)",
        type=["xlsx"],
        key="admin_up_bi"
    )
    up_unidades = st.sidebar.file_uploader(
        "3. Relatório de Unidades (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="admin_up_unidades"
    )

    if st.sidebar.button("💾 Salvar Bases na Nuvem", use_container_width=True):
        arquivos_salvos = []
        if up_acomp is not None:
            with open("acompanhamento.xlsx", "wb") as f:
                f.write(up_acomp.getbuffer())
            arquivos_salvos.append("Acompanhamento")
        if up_bi is not None:
            with open("dados_bi.xlsx", "wb") as f:
                f.write(up_bi.getbuffer())
            arquivos_salvos.append("Dados BI")
        if up_unidades is not None:
            with open("unidades.xlsx", "wb") as f:
                f.write(up_unidades.getbuffer())
            arquivos_salvos.append("Unidades")

        if arquivos_salvos:
            st.cache_data.clear()
            st.sidebar.success(f"✅ Bases salvas e atualizadas na nuvem: {', '.join(arquivos_salvos)}!")
            st.rerun()
        else:
            st.sidebar.warning("⚠️ Selecione pelo menos uma planilha para salvar.")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Período do Relatório")

col_dt1, col_dt2 = st.sidebar.columns(2)
with col_dt1:
    data_inicio = st.date_input("Data Inicial", value=None, format="DD/MM/YYYY")
with col_dt2:
    data_fim = st.date_input("Data Final", value=None, format="DD/MM/YYYY")

# ---------------------------------------------------------
# CARREGAMENTO DOS DADOS (MODO SERVIDOR / NUVEM)
# ---------------------------------------------------------
file_acompanhamento = 'acompanhamento.xlsx'
file_bi = 'dados_bi.xlsx'
file_unidades = 'unidades.xlsx'


# ---------------------------------------------------------
# 3. PROCESSAMENTO E REGRAS DE NEGÓCIO
# ---------------------------------------------------------
@st.cache_data(show_spinner="Processando as três planilhas...")
def process_data(file_acomp, file_bi_data, file_unid_data, _cache_version="v7_perfis_acesso", _mtime=None):
    # Leitura dos arquivos principais
    df_acomp_raw = load_uploaded_file(file_acomp)
    df_bi_raw = load_uploaded_file(file_bi_data)

    # --- 3.1 PROCESSAR ACOMPANHAMENTO ---
    col_cliente_acomp = find_column(df_acomp_raw, ["nome_cliente", "cliente", "nome cliente"])
    if col_cliente_acomp:
        df_acomp = df_acomp_raw[
            df_acomp_raw[col_cliente_acomp].astype(str).str.contains("INCONFIDENCIA", case=False, na=False)
        ].copy()
    else:
        df_acomp = df_acomp_raw.copy()

    # Limpar a coluna Nota: converter para numérico, preencher nulos com 0, converter para int e string
    col_nota_acomp = find_column(df_acomp, ["Nota", "nota", "n_nota", "numero_nota"])
    if col_nota_acomp:
        nota_num = pd.to_numeric(df_acomp[col_nota_acomp], errors="coerce").fillna(0).astype(int)
        df_acomp["Nota"] = nota_num.astype(str)
    else:
        df_acomp["Nota"] = "0"

    # Identificar a coluna de Pedido no Acompanhamento
    col_pedido_acomp = find_column(df_acomp, ["Pedido", "pedido", "n_pedido", "num_pedido", "nr_pedido", "numero_pedido", "ordem", "num_ordem"])
    if col_pedido_acomp and col_pedido_acomp != "Pedido":
        df_acomp["Pedido"] = df_acomp[col_pedido_acomp]
    elif "Pedido" not in df_acomp.columns:
        df_acomp["Pedido"] = [f"PED-{i+1}" for i in range(len(df_acomp))]

    df_acomp["Pedido"] = df_acomp["Pedido"].astype(str).str.strip()

    # Garantir existência e limpeza da coluna 'cidade' para Fallback
    col_cidade = find_column(df_acomp, ["cidade", "municipio", "cid"])
    if col_cidade and col_cidade != "cidade":
        df_acomp["cidade"] = df_acomp[col_cidade]
    elif "cidade" not in df_acomp.columns:
        df_acomp["cidade"] = "Não Informada"
    df_acomp["cidade"] = df_acomp["cidade"].fillna("Não Informada").astype(str).str.strip()
    df_acomp["cidade"] = df_acomp["cidade"].replace({"": "Não Informada", "nan": "Não Informada", "None": "Não Informada"})

    # Mapear outras colunas esperadas do Acompanhamento
    col_dt_inser = find_column(df_acomp, ["Dt_inser", "dt_inser", "dt inser", "data_insercao", "data"])
    if col_dt_inser and col_dt_inser != "Dt_inser":
        df_acomp["Dt_inser"] = df_acomp[col_dt_inser]
    elif "Dt_inser" not in df_acomp.columns:
        df_acomp["Dt_inser"] = "-"

    col_status = find_column(df_acomp, ["Status", "status", "situacao"])
    if col_status and col_status != "Status":
        df_acomp["Status"] = df_acomp[col_status]
    elif "Status" not in df_acomp.columns:
        df_acomp["Status"] = "Indefinido"

    col_tipo = find_column(df_acomp, ["Tipo", "tipo"])
    if col_tipo and col_tipo != "Tipo":
        df_acomp["Tipo"] = df_acomp[col_tipo]
    elif "Tipo" not in df_acomp.columns:
        df_acomp["Tipo"] = "Padrão"

    df_acomp["Status"] = df_acomp["Status"].fillna("Não Informado").astype(str)
    df_acomp["Tipo"] = df_acomp["Tipo"].fillna("Não Informado").astype(str)

    # --- 3.2 PROCESSAR DADOS BI ---
    col_cliente_bi = find_column(df_bi_raw, ["Cliente", "cliente", "nome_cliente"])
    if col_cliente_bi:
        df_bi = df_bi_raw[
            df_bi_raw[col_cliente_bi].astype(str).str.contains("INCONFIDENCIA", case=False, na=False)
        ].copy()
    else:
        df_bi = df_bi_raw.copy()

    # Limpar a coluna n_nota_obs
    col_nota_bi = find_column(df_bi, ["n_nota_obs", "nota", "numero_nota", "n_nota"])
    if col_nota_bi:
        nota_bi_num = pd.to_numeric(df_bi[col_nota_bi], errors="coerce")
        df_bi = df_bi[nota_bi_num.notna()].copy()
        df_bi["n_nota_obs"] = nota_bi_num.dropna().astype(int).astype(str)
    else:
        df_bi["n_nota_obs"] = ""

    # Padronizar colunas de valores e itens no BI
    col_vendas = find_column(df_bi, ["Vendas Total", "vendas total", "vlr_total", "valor_total", "total"])
    if col_vendas:
        df_bi["Vendas Total"] = pd.to_numeric(df_bi[col_vendas], errors="coerce").fillna(0.0)
    else:
        df_bi["Vendas Total"] = 0.0

    col_qtd = find_column(df_bi, ["QTD", "qtd", "quantidade", "quant"])
    if col_qtd:
        df_bi["QTD"] = pd.to_numeric(df_bi[col_qtd], errors="coerce").fillna(0)
    else:
        df_bi["QTD"] = 0

    col_dt_hist = find_column(df_bi, ["dt_nota_hist", "data_nota", "dt_emissao", "data_faturamento"])
    if col_dt_hist and col_dt_hist != "dt_nota_hist":
        df_bi["dt_nota_hist"] = df_bi[col_dt_hist]
    elif "dt_nota_hist" not in df_bi.columns:
        df_bi["dt_nota_hist"] = "-"

    # Padronizar campos de itens para a Aba 2 (Cod, Descr. produto, Preço Venda)
    col_cod = find_column(df_bi, ["Cod", "cod", "codigo", "cod_produto", "cod_item"])
    if col_cod and col_cod != "Cod":
        df_bi["Cod"] = df_bi[col_cod]
    elif "Cod" not in df_bi.columns:
        df_bi["Cod"] = "-"

    col_descr = find_column(df_bi, ["Descr. produto", "descr. produto", "descricao", "descr_produto", "produto"])
    if col_descr and col_descr != "Descr. produto":
        df_bi["Descr. produto"] = df_bi[col_descr]
    elif "Descr. produto" not in df_bi.columns:
        df_bi["Descr. produto"] = "-"

    col_preco = find_column(df_bi, ["Preço Venda", "preço venda", "preco venda", "preco_venda", "unitario"])
    if col_preco and col_preco != "Preço Venda":
        df_bi["Preço Venda"] = pd.to_numeric(df_bi[col_preco], errors="coerce").fillna(0.0)
    elif "Preço Venda" not in df_bi.columns:
        df_bi["Preço Venda"] = np.where(df_bi["QTD"] > 0, df_bi["Vendas Total"] / df_bi["QTD"], 0.0)

    # Resumo agrupado do BI por n_nota_obs
    resumo_bi = df_bi.groupby("n_nota_obs", as_index=False).agg({
        "Vendas Total": "sum",
        "QTD": "sum",
        "dt_nota_hist": "first"
    }).rename(columns={
        "Vendas Total": "bi_vendas_total",
        "QTD": "bi_qtd_total",
        "dt_nota_hist": "bi_dt_nota_hist"
    })

    # --- 3.3 CRUZAMENTO (LEFT JOIN ACOMPANHAMENTO + BI VIA NF_CLEAN) ---
    # Chave 1: Nota Fiscal (BI <-> Acomp)
    df_bi['NF_Clean'] = df_bi['n_nota_obs'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
    df_acomp['NF_Clean'] = df_acomp['Nota'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')

    # Resumo agrupado do BI por NF_Clean
    resumo_bi = df_bi[df_bi['NF_Clean'] != ''].groupby("NF_Clean", as_index=False).agg({
        "Vendas Total": "sum",
        "QTD": "sum",
        "dt_nota_hist": "first"
    }).rename(columns={
        "Vendas Total": "bi_vendas_total",
        "QTD": "bi_qtd_total",
        "dt_nota_hist": "bi_dt_nota_hist"
    })

    # Merge 1 (Acomp + BI) -> Agora temos o faturamento atrelado aos pedidos
    df_merged = pd.merge(
        df_acomp,
        resumo_bi,
        on="NF_Clean",
        how="left"
    )

    is_faturado = (df_merged["NF_Clean"] != "") & (df_merged["bi_vendas_total"].notna())

    df_merged["Status_Faturamento"] = np.where(is_faturado, "Faturado", "Pendente / Não Faturado")
    df_merged["Vendas Total"] = np.where(is_faturado, df_merged["bi_vendas_total"], 0.0)
    df_merged["QTD Total"] = np.where(is_faturado, df_merged["bi_qtd_total"], 0)

    # Formatação prévia segura para texto evitando DTypePromotionError no NumPy
    dt_faturado_str = pd.to_datetime(df_merged["bi_dt_nota_hist"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("-")
    df_merged["dt_nota_hist"] = np.where(is_faturado, dt_faturado_str, "-")

    # --- 3.4 LEITURA DA 3ª PLANILHA E MERGE DEFINITIVO (Base <-> Relatório de Unidades) ---
    if isinstance(file_unid_data, str):
        with open(file_unid_data, "rb") as f:
            file_unidades_target = io.BytesIO(f.read())
        file_unidades_target.name = file_unid_data
    else:
        file_unidades_target = file_unid_data if 'file_unid_data' in locals() and file_unid_data is not None else file_unidades
        if hasattr(file_unidades_target, "seek"):
            file_unidades_target.seek(0)

    try:
        # Lê a 3ª planilha contornando o formato HTML/XLS exportado pelo ERP
        tables = pd.read_html(file_unidades_target, decimal=',', thousands='.')
        df_unidades = max(tables, key=lambda t: t.shape[0] * t.shape[1]) if tables else tables[0]
        # Achata o cabeçalho caso venha MultiIndex ("RELATÓRIO PORTAL", "Codigo pedido")
        if isinstance(df_unidades.columns, pd.MultiIndex):
            df_unidades.columns = df_unidades.columns.get_level_values(-1)
    except Exception:
        file_unidades_target.seek(0)
        try:
            df_unidades = pd.read_excel(file_unidades_target, skiprows=1) # Fallback padrão pulando título
        except Exception:
            file_unidades_target.seek(0)
            df_unidades = pd.read_excel(file_unidades_target)

    # Limpa espaços e nomes das colunas
    df_unidades.columns = [str(c).strip() for c in df_unidades.columns]

    # Localiza as colunas de 'Codigo pedido' e 'Unidade'
    ped_matches = [c for c in df_unidades.columns if 'codigo pedido' in str(c).lower()]
    if not ped_matches:
        ped_matches = [c for c in df_unidades.columns if any(p in str(c).lower() for p in ['pedido', 'doc']) and 'item' not in str(c).lower() and 'data' not in str(c).lower()]
    col_ped = ped_matches[0] if ped_matches else (df_unidades.columns[1] if len(df_unidades.columns) > 1 else df_unidades.columns[0])

    uni_matches = [c for c in df_unidades.columns if 'unidade' in str(c).lower() and 'data' not in str(c).lower()]
    col_uni = uni_matches[0] if uni_matches else (df_unidades.columns[3] if len(df_unidades.columns) > 3 else df_unidades.columns[-1])

    # Chave 2: Pedido (Base <-> Relatório de Unidades)
    ped_col_data = df_unidades[col_ped]
    if isinstance(ped_col_data, pd.DataFrame):
        ped_col_data = ped_col_data.iloc[:, 0]

    uni_col_data = df_unidades[col_uni]
    if isinstance(uni_col_data, pd.DataFrame):
        uni_col_data = uni_col_data.iloc[:, 0]

    df_merged['Pedido_Clean'] = df_merged['Pedido'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
    df_unidades['Pedido_Clean'] = ped_col_data.astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
    df_unidades['Unidade_Extraida'] = uni_col_data

    # Remove vazios e duplicatas, faz o merge
    df_unidades_unique = df_unidades[df_unidades['Pedido_Clean'] != ''].drop_duplicates(subset=['Pedido_Clean'])

    # Merge 2 (Base + Unidades)
    if 'Unidade_Extraida' in df_merged.columns:
        df_merged = df_merged.drop(columns=['Unidade_Extraida'])

    df_merged = pd.merge(df_merged, df_unidades_unique[['Pedido_Clean', 'Unidade_Extraida']], on='Pedido_Clean', how='left')

    # Define Unidade Final
    df_merged['Unidade_Final'] = df_merged['Unidade_Extraida'].fillna(df_merged['cidade'])
    df_merged['Unidade_Final'] = df_merged['Unidade_Final'].fillna('Não Informada').replace({'': 'Não Informada', 'nan': 'Não Informada', 'None': 'Não Informada'})

    # Sincronizar colunas de unidade para os filtros e visões
    df_merged["UNIDADE"] = df_merged["Unidade_Final"]
    df_merged["Unidade"] = df_merged["Unidade_Final"]
    df_merged["Unidade_Real"] = df_merged["Unidade_Final"]

    # Também enriquece df_bi com Pedido e Unidade_Final para os itens da Aba 2
    df_bi_merged = pd.merge(df_bi, df_merged[['NF_Clean', 'Pedido', 'Pedido_Clean', 'Unidade_Final']].drop_duplicates(subset=['NF_Clean']), on='NF_Clean', how='left')
    df_bi_merged['UNIDADE'] = df_bi_merged['Unidade_Final'].fillna('Não Informada')

    return df_merged, df_bi_merged, df_unidades_unique, col_ped, 'Unidade_Extraida'

# Verifica se os arquivos existem no servidor antes de tentar processar
if os.path.exists(file_acompanhamento) and os.path.exists(file_bi) and os.path.exists(file_unidades):
    mtimes = (
        os.path.getmtime(file_acompanhamento),
        os.path.getmtime(file_bi),
        os.path.getmtime(file_unidades)
    )
    df_merged, df_bi_raw, df_unidades_unique, col_ped, col_uni = process_data(file_acompanhamento, file_bi, file_unidades, _mtime=mtimes)
else:
    st.warning("⚠️ Os arquivos de banco de dados não foram encontrados no servidor. Por favor, contate o administrador.")
    st.stop()


# ---------------------------------------------------------
# 4. FILTROS NA SIDEBAR
# ---------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Filtros de Visualização")

# Filtro UNIDADE
lista_unidades = sorted(df_merged["UNIDADE"].unique().tolist())
selected_unidades = st.sidebar.multiselect(
    "Unidade de Entrega",
    options=lista_unidades,
    default=[]
)

# Filtro Status (da planilha acompanhamento)
lista_status = sorted(df_merged["Status"].unique().tolist())
selected_status = st.sidebar.multiselect(
    "Status do Acompanhamento",
    options=lista_status,
    default=[]
)

# Filtro Tipo
lista_tipo = sorted(df_merged["Tipo"].unique().tolist())
selected_tipo = st.sidebar.multiselect(
    "Tipo de Pedido",
    options=lista_tipo,
    default=[]
)

# Filtro Busca por Pedido ou Nota
busca_texto = st.sidebar.text_input("Buscar por Pedido ou Nota:", value="").strip()

# Aplicação dos Filtros
df_filtrado = df_merged.copy()

# Filtro de Período (Data Inicial e Data Final)
if data_inicio or data_fim:
    serie_dt = pd.to_datetime(df_filtrado['Dt_inser'], dayfirst=True, errors='coerce').dt.date
    if data_inicio and data_fim:
        if data_inicio <= data_fim:
            df_filtrado = df_filtrado[(serie_dt >= data_inicio) & (serie_dt <= data_fim)]
        else:
            st.sidebar.error("A Data Final deve ser maior que a Inicial.")
    elif data_inicio:
        df_filtrado = df_filtrado[serie_dt >= data_inicio]
    elif data_fim:
        df_filtrado = df_filtrado[serie_dt <= data_fim]

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
# 5. CABEÇALHO DO DASHBOARD E KPIS (st.metric)
# ---------------------------------------------------------
st.markdown("""
<div style="margin-bottom: 22px;">
    <h1 style="font-size: 2.2rem; font-weight: 800; margin-bottom: 4px; color: #1E293B;">
        Dashboard Comercial | <span style="color:#E6192B;">Grupo Fortpel</span> & <span style="color:#3B6077;">Inconfidência</span>
    </h1>
    <p style="font-size: 1.05rem; color: #64748B; margin-top: 0px;">
        Painel executivo de acompanhamento comercial, unidades de entrega e integração de faturamento com BI.
    </p>
</div>
""", unsafe_allow_html=True)

kpi_total_pedidos = len(df_filtrado)
kpi_faturados = int((df_filtrado["Status_Faturamento"] == "Faturado").sum())
kpi_pendentes = int((df_filtrado["Status_Faturamento"] != "Faturado").sum())
kpi_valor_total = float(df_filtrado["Vendas Total"].sum())

col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

with col_kpi1:
    st.metric(label="Total de Pedidos", value=f"{kpi_total_pedidos:,}".replace(",", "."))

with col_kpi2:
    st.metric(label="Pedidos Faturados", value=f"{kpi_faturados:,}".replace(",", "."), delta=f"{(kpi_faturados / kpi_total_pedidos * 100):.1f}% do total" if kpi_total_pedidos > 0 else None)

with col_kpi3:
    st.metric(label="Pedidos Pendentes / Sem Nota", value=f"{kpi_pendentes:,}".replace(",", "."), delta=f"-{(kpi_pendentes / kpi_total_pedidos * 100):.1f}%" if kpi_total_pedidos > 0 else None, delta_color="inverse")

with col_kpi4:
    st.metric(label="Valor Total Faturado", value=format_currency_brl(kpi_valor_total))

st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# PAINEL DE DIAGNÓSTICO (DEBUG VISUAL DE CRUZAMENTO)
# ---------------------------------------------------------
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
# 6. ABAS: VISÃO POR UNIDADE & PEDIDOS E ITENS
# ---------------------------------------------------------
tab_unidade, tab_pedidos = st.tabs(["🏢 Visão por Unidade", "📦 Pedidos e Itens"])

# ----------------- ABA 1: VISÃO POR UNIDADE -----------------
with tab_unidade:
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

    df_tabela['Data_Formatada'] = pd.to_datetime(df_tabela['Data de Inserção'], errors='coerce', dayfirst=True)

    if data_inicio and data_fim:
        # Garante que a data final seja sempre maior ou igual à inicial para não dar erro
        if data_inicio <= data_fim:
            df_tabela = df_tabela[(df_tabela['Data_Formatada'].dt.date >= data_inicio) & (df_tabela['Data_Formatada'].dt.date <= data_fim)]
        else:
            st.sidebar.error("A Data Final deve ser maior que a Inicial.")
    elif data_inicio:
        df_tabela = df_tabela[df_tabela['Data_Formatada'].dt.date >= data_inicio]
    elif data_fim:
        df_tabela = df_tabela[df_tabela['Data_Formatada'].dt.date <= data_fim]

    df_tabela = df_tabela.drop(columns=['Data_Formatada'], errors='ignore')

    # Formata Moeda
    df_tabela['Valor (R$)'] = pd.to_numeric(df_tabela['Valor (R$)'], errors='coerce').fillna(0)
    df_tabela['Valor (R$)'] = df_tabela['Valor (R$)'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

    df_tabela['Data de Inserção'] = df_tabela['Data de Inserção'].fillna('-')
    df_tabela['Pedido'] = df_tabela['Pedido'].fillna('-')

    colunas_exibicao = ['Pedido', 'Unidade', 'Data de Inserção', 'Valor (R$)']
    st.dataframe(df_tabela[colunas_exibicao], use_container_width=True, height=600, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Gera o Excel em memória
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_tabela[colunas_exibicao].to_excel(writer, index=False, sheet_name='Relatorio_Filtrado')

    # Exibe o botão de download
    st.download_button(
        label="📥 Baixar Relatório Filtrado (Excel)",
        data=buffer.getvalue(),
        file_name="Relatorio_Comercial_Inconfidencia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Alerta expansível de Pedidos sem correspondência (Fallback)
    pedidos_orfaos = df_merged[df_merged['Unidade_Final'] == df_merged['cidade']][['Pedido', 'cidade', 'Status', 'Dt_inser']]

    if not pedidos_orfaos.empty:
        with st.expander("⚠️ Atenção: Pedidos sem correspondência no Relatório de Unidades"):
            st.write("Os pedidos abaixo não constam na planilha de Unidades (ou possuem numeração divergente). A cidade foi usada como Unidade provisória para não perder o registro do faturamento:")
            st.dataframe(pedidos_orfaos, use_container_width=True, hide_index=True)


# ----------------- ABA 2: PEDIDOS E ITENS -----------------
with tab_pedidos:
    st.subheader("Relação Detalhada de Pedidos")

    if df_filtrado.empty:
        st.warning("Nenhum pedido encontrado para os filtros selecionados.")
    else:
        colunas_pedidos = ["Pedido", "Nota", "UNIDADE", "Dt_inser", "dt_nota_hist", "Status", "Tipo", "Vendas Total", "Status_Faturamento"]
        df_pedidos_view = df_filtrado[colunas_pedidos].copy()

        df_pedidos_exibicao = df_pedidos_view.copy()
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

        st.dataframe(
            df_pedidos_exibicao,
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")
        st.subheader("🔍 Detalhamento e Drill-Down do Pedido")

        pedidos_disponiveis = df_filtrado["Pedido"].astype(str).unique().tolist()
        
        pedido_selecionado = st.selectbox(
            "Selecione um Pedido para visualizar os detalhes e itens:",
            options=pedidos_disponiveis,
            index=0
        )

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

            # Drill-down de itens da nota fiscal
            if nota_pedido == "0" or nota_pedido == "" or nota_pedido.lower() == "nan":
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

                    st.dataframe(
                        itens_exibicao,
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning(f"⚠️ A Nota Fiscal `{nota_pedido}` não foi encontrada na base de dados do BI.")
