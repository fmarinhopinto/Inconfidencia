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
import datetime
import codecs
from unidecode import unidecode
try:
    import ftfy
except ImportError:
    ftfy = None

# ---------------------------------------------------------
# 1. CONFIGURAÇÃO DA PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    layout="wide",
    page_title="Dashboard Comercial - Grupo Fortpel & Inconfidência",
    page_icon="🏢",
    initial_sidebar_state="collapsed"
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
    /* Força o container principal a ocupar 100% da largura útil sem margens restritivas */
    .block-container {
        max-width: 100% !important;
        width: 100% !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
    }

    /* Anula qualquer centralização ou max-width padrão do app */
    .main .block-container {
        max-width: 100% !important;
    }

    /* Estica todos os blocos horizontais de métricas e KPIs para 100% */
    [data-testid="stHorizontalBlock"] {
        width: 100% !important;
        gap: 1.5rem !important;
    }
    
    /* Estica as colunas internas e dataframes de ponta a ponta */
    [data-testid="column"] {
        width: 100% !important;
        flex: 1 1 0% !important;
    }

    div[data-testid="stDataFrame"],
    div[data-testid="stDataFrame"] > div {
        width: 100% !important;
    }

    [data-testid="stAppViewContainer"], .main {
        background-color: #F8F9FA !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    [data-testid="stHeader"] {
        background-color: rgba(248, 249, 250, 0.85) !important;
    }

    /* 2. BARRA LATERAL (SIDEBAR) - AZUL PETRÓLEO COM LARGURA CONFORTÁVEL */
    [data-testid="stSidebar"],
    section[data-testid="stSidebar"] {
        background-color: #3B6077 !important;
        color: #FFFFFF !important;
        min-width: 340px !important;
        max-width: 380px !important;
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

    /* MANTÉM O BOTÃO DE RECOLHER/ABRIR A BARRA LATERAL SEMPRE VISÍVEL E GRANDE */
    [data-testid="collapsedControl"] {
        display: block !important;
        opacity: 1 !important;
        visibility: visible !important;
        color: #FFFFFF !important;
        background-color: #3B6077 !important; /* Cor de fundo sólida para destacar */
        border-radius: 8px !important;
        padding: 8px !important;
        transform: scale(1.2);
        margin-top: 10px !important;
        margin-left: 10px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
        z-index: 999999 !important;
    }
    
    /* Ícone interno branco */
    [data-testid="collapsedControl"] svg {
        fill: #FFFFFF !important;
        stroke: #FFFFFF !important;
        color: #FFFFFF !important;
        width: 22px !important;
        height: 22px !important;
    }

    [data-testid="collapsedControl"]:hover {
        background-color: #E6192B !important;
    }

    /* Botão interno de recolher (quando a barra lateral está aberta) */
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarHeader"] button,
    section[data-testid="stSidebar"] button[kind="header"] {
        display: block !important;
        opacity: 1 !important;
        visibility: visible !important;
        color: #FFFFFF !important;
        background-color: rgba(255, 255, 255, 0.2) !important;
        border-radius: 8px !important;
        padding: 6px !important;
        transform: scale(1.2);
        transition: all 0.2s ease !important;
    }

    [data-testid="stSidebarCollapseButton"] svg,
    [data-testid="stSidebarHeader"] button svg,
    section[data-testid="stSidebar"] button[kind="header"] svg {
        fill: #FFFFFF !important;
        stroke: #FFFFFF !important;
        color: #FFFFFF !important;
        width: 22px !important;
        height: 22px !important;
    }

    [data-testid="stSidebarCollapseButton"]:hover,
    [data-testid="stSidebarHeader"] button:hover,
    section[data-testid="stSidebar"] button[kind="header"]:hover {
        background-color: rgba(255, 255, 255, 0.35) !important;
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
    /* Garante que o input de data ocupe 100% da largura da barra lateral com fonte preta */
    div[data-testid="stDateInput"] {
        width: 100% !important;
    }
    div[data-testid="stDateInput"] input,
    div[data-testid="stDateInput"] span,
    section[data-testid="stSidebar"] div[data-testid="stDateInput"] input,
    section[data-testid="stSidebar"] div[data-testid="stDateInput"] span {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        font-size: 14px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stDateInput"] label,
    section[data-testid="stSidebar"] div[data-testid="stDateInput"] label p,
    section[data-testid="stSidebar"] div[data-testid="stDateInput"] label span {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stDateInput"] input {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        width: 100% !important;
    }

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

def limpar_texto_corrompido(texto):
    """
    Função avançada de sanitização e normalização de texto:
    - Trata problemas de codificação, mojibake e double-encoding (ex.: ASSOCIAÃ±Ã£O -> ASSOCIAÇÃO)
    - Remove acentos, caracteres corrompidos e padroniza para maiúsculas
    - Remove pontuações e símbolos preservando letras e números essenciais para uniformizar o match
    """
    if pd.isna(texto) or texto is None:
        return ""
    
    txt = str(texto).strip()
    if txt.lower() in ["nan", "none", "null", "-", ""]:
        return ""
    
    # 1. Tratamento preliminar para padrões conhecidos de mojibake corrompido
    padroes_mojibake = {
        "Ã±Ã£O": "CAO", "Ã±Ã£o": "cao", "Ã§Ã£o": "cao", "Ã§Ã£O": "CAO",
        "Ã°": "", "Â°": "", "Âº": "", "º": "", "°": "", "ª": "", 
        "\ufffd": "", "ð": "", "Ð": "", "?": ""
    }
    for err, fix in padroes_mojibake.items():
        txt = txt.replace(err, fix)

    # 2. Corrige com ftfy se disponível para reparar mojibakes complexos
    if ftfy is not None:
        try:
            txt = ftfy.fix_text(txt)
        except Exception:
            pass

    # 3. Tenta corrigir erros comuns de encoding do Excel/HTML exportado (double-encoding)
    try:
        txt = txt.encode('latin1').decode('utf-8')
    except Exception:
        try:
            txt = txt.encode('cp1252').decode('utf-8')
        except Exception:
            pass
        
    # Tratamento secundário de caracteres ordinais remanescentes
    for err, fix in padroes_mojibake.items():
        txt = txt.replace(err, fix)

    # 4. Remove acentos, caracteres especiais corrompidos e normaliza para maiúsculas
    txt = unidecode(txt).upper()
    
    # 5. Remove tudo o que não for letra ou número (mantendo espaços simples) para uniformizar o match
    txt = re.sub(r'[^A-Z0-9\s]', ' ', txt)
    txt = re.sub(r'\s+', ' ', txt).strip()
    
    return txt

# Aliases para compatibilidade retroativa
normalizar_texto = limpar_texto_corrompido
normalizar_nome_unidade = limpar_texto_corrompido

def limpar_id(val):
    """Garante string sem ponto flutuante e sem espaços para chave de cruzamento."""
    if pd.isna(val):
        return ""
    texto = str(val).strip()
    texto = re.sub(r'\.0$', '', texto)
    match = re.search(r'(\d+)', texto)
    return match.group(1) if match else texto

def sanitizar_chave_pedido(serie):
    """
    Padroniza qualquer coluna de Pedido para string numérica limpa:
    - Converte '2406179.0' -> '2406179'
    - Remove espaços e caracteres não numéricos
    """
    return (
        serie.astype(str)
        .str.replace(r'\.0$', '', regex=True) # remove decimal se veio como float
        .str.extract(r'(\d+)', expand=False)   # extrai estritamente os dígitos
        .fillna('')
        .str.strip()
    )

def cruzar_unidades_blindado(df_alvo, df_unidades):
    """
    Cruza o DataFrame (seja ele a base normal ou emergencial)
    com a planilha de Unidades de forma infalível.
    """
    if df_unidades is None or df_unidades.empty or df_alvo is None or df_alvo.empty:
        return df_alvo

    df_u = df_unidades.copy()

    # 1. Identificar as colunas de Pedido e Unidade na base de Unidades
    # Baseado na planilha real: Coluna B = Pedido, Coluna D = Unidade
    if 'Pedido' in df_u.columns:
        col_ped = 'Pedido'
    elif any('codigo pedido' in str(c).lower() for c in df_u.columns):
        col_ped = [c for c in df_u.columns if 'codigo pedido' in str(c).lower()][0]
    elif any(any(p in str(c).lower() for p in ['pedido', 'doc']) for c in df_u.columns):
        col_ped = [c for c in df_u.columns if any(p in str(c).lower() for p in ['pedido', 'doc'])][0]
    else:
        # Se veio sem cabeçalho nomeado, pega pela posição (Coluna B = índice 1)
        col_ped = df_u.columns[1] if len(df_u.columns) > 1 else df_u.columns[0]

    if 'Unidade' in df_u.columns:
        col_unid = 'Unidade'
    elif any('unidade' in str(c).lower() for c in df_u.columns):
        col_unid = [c for c in df_u.columns if 'unidade' in str(c).lower()][0]
    else:
        # Se veio sem cabeçalho nomeado, pega pela posição (Coluna D = índice 3)
        col_unid = df_u.columns[3] if len(df_u.columns) > 3 else df_u.columns[-1]

    # Identifica coluna de Pedido em df_alvo
    col_alvo_ped = 'Pedido'
    if 'Pedido' not in df_alvo.columns:
        cands = [c for c in df_alvo.columns if any(p in str(c).lower() for p in ['pedido', 'ped', 'ordem'])]
        if cands:
            col_alvo_ped = cands[0]
        else:
            return df_alvo

    # 2. Padroniza as chaves em ambos os DataFrames para string pura
    df_u['chave_join'] = sanitizar_chave_pedido(df_u[col_ped])
    df_alvo['chave_join'] = sanitizar_chave_pedido(df_alvo[col_alvo_ped])

    # 3. Cria o mapa de pedidos únicos para unidades (ignora linhas duplicadas de itens)
    mapa_unidades = (
        df_u[df_u['chave_join'] != '']
        .drop_duplicates(subset=['chave_join'])
        .set_index('chave_join')[col_unid]
        .to_dict()
    )

    # 4. Aplica o mapeamento apenas onde encontrou correspondência
    unidade_mapeada = df_alvo['chave_join'].map(mapa_unidades)
    
    # Atualiza a coluna Unidade sem sobrescrever com nulo se já existia algo válido
    if 'Unidade' in df_alvo.columns:
        df_alvo['Unidade'] = unidade_mapeada.combine_first(df_alvo['Unidade'])
    else:
        df_alvo['Unidade'] = unidade_mapeada

    # Se df_alvo tiver colunas de unidade derivadas, sincroniza onde cabível
    for col_sync in ['Unidade_Final', 'UNIDADE', 'Unidade_Real']:
        if col_sync in df_alvo.columns:
            df_alvo[col_sync] = unidade_mapeada.combine_first(df_alvo[col_sync])

    # Remove a coluna temporária de ligação
    df_alvo.drop(columns=['chave_join'], inplace=True, errors='ignore')

    return df_alvo

def extrair_unidade_excecao(texto):
    """
    Extrai cirurgicamente a Unidade do campo 'Local de entrega'
    ignorando prefixos, códigos de cliente e cortando antes do endereço.
    """
    if pd.isna(texto):
        return None
    
    t = str(texto).strip()
    
    # 1. Corta tudo antes do PED.: XXXXXXX
    match_ped = re.search(r'PED\.?:\s*\d+\s*[-–—]+', t, re.IGNORECASE)
    if not match_ped:
        # Se não achou 'PED.:', tenta pegar depois do início
        miolo = t
    else:
        miolo = t[match_ped.end():].strip()
    
    # 2. Remove código numérico de cliente inicial (ex: '33466 - ', '20195 - ', '- ')
    miolo = re.sub(r'^(?:\d+|-)\s*[-–—]+\s*', '', miolo).strip()
    
    # 3. Corta tudo a partir do endereço de rua/rodovia/av/estrada
    # Trunca onde começar o logradouro
    padroes_corte = [
        r'\s*[-–—]+\s*RUA\b',
        r'\s*[-–—]+\s*RODOVIA\b',
        r'\s*[-–—]+\s*ROD\b',
        r'\s*[-–—]+\s*AV\b',
        r'\s*[-–—]+\s*AVENIDA\b',
        r'\s*[-–—]+\s*ESTRADA\b',
        r'\s*[-–—]+\s*TRAVESSA\b',
        r'\s*[-–—]+\s*ALAMEDA\b'
    ]
    
    for padrao in padroes_corte:
        partes = re.split(padrao, miolo, flags=re.IGNORECASE)
        if len(partes) > 1:
            miolo = partes[0]
            break
            
    # Se ainda tiver traços duplos no final, limpa
    miolo = re.sub(r'[-–—\s]+$', '', miolo).strip()
    
    return miolo if miolo else None

def aplicar_regra_excecao(df_emergencial):
    """
    Aplica a extração de Pedido e Unidade específica para a planilha de exceções.
    """
    if df_emergencial is None or df_emergencial.empty:
        return df_emergencial

    df = df_emergencial.copy()
    
    # Localiza a coluna 'Local de entrega'
    col_local = next((c for c in df.columns if "local de entrega" in str(c).lower().strip()), None)
    
    if col_local:
        # Extrai número do Pedido (somente dígitos)
        df['Pedido'] = df[col_local].astype(str).str.extract(r'PED\.?:\s*(\d+)', expand=False)
        
        # Extrai o nome da Unidade direto do texto
        df['Unidade_Excecao'] = df[col_local].apply(extrair_unidade_excecao)
        
        # Força o nome da unidade na coluna oficial 'Unidade'
        df['Unidade'] = df['Unidade_Excecao']
        
        # Remove a coluna temporária
        df.drop(columns=['Unidade_Excecao'], inplace=True, errors='ignore')
        
    return df

def extrair_dados_entrega(texto):
    """
    Extrai tanto o número do pedido quanto o nome da unidade
    direto da string de 'Local de entrega' da Inconfidência.
    """
    if pd.isna(texto):
        return pd.Series([None, None])
    
    texto_str = str(texto).strip()
    
    # 1. Extrai o número do Pedido (apenas dígitos)
    match_ped = re.search(r'PED\.?:\s*(\d+)', texto_str, re.IGNORECASE)
    num_pedido = match_ped.group(1) if match_ped else None
    
    # 2. Extrai cirurgicamente a unidade
    nome_unidade_fallback = extrair_unidade_excecao(texto_str)
    if not nome_unidade_fallback:
        partes = [p.strip() for p in re.split(r'\s*-\s*', texto_str) if p.strip()]
        for p in partes:
            p_upper = p.upper()
            if any(chave in p_upper for chave in ['BANRISUL', 'PRF', 'DIPESUL', 'BRDE', 'MAGNO', 'BRASLASER', 'SEDE', 'AJURIS', 'DNV']):
                nome_unidade_fallback = p
                break
            
    return pd.Series([num_pedido, nome_unidade_fallback])

def processar_planilha_emergencial(df_emergencial, df_unidades=None):
    """
    Extrai o número do pedido real e a unidade a partir do texto de 'Local de entrega'
    e ajusta as colunas para o padrão do relatório.
    """
    if df_emergencial is None or df_emergencial.empty:
        return df_emergencial

    # Aplica regra cirúrgica de exceção
    df_emergencial = aplicar_regra_excecao(df_emergencial)

    # 2. Se a base de unidades oficial estiver disponível, preenche os casos faltantes
    if df_unidades is not None and not df_unidades.empty and 'Pedido' in df_emergencial.columns:
        df_u = df_unidades.copy()
        col_ped_u = df_u.columns[1] if 'Pedido' not in df_u.columns else 'Pedido'
        col_nome_u = df_u.columns[3] if len(df_u.columns) > 3 and 'Unidade' not in df_u.columns else ('Unidade' if 'Unidade' in df_u.columns else df_u.columns[-1])
        
        df_u['chave_limpa'] = (
            df_u[col_ped_u]
            .astype(str)
            .str.replace(r'\.0$', '', regex=True)
            .str.strip()
        )
        
        mapa_unidades = (
            df_u[df_u['chave_limpa'] != '']
            .drop_duplicates(subset=['chave_limpa'])
            .set_index('chave_limpa')[col_nome_u]
            .to_dict()
        )
        
        chave_emerg = (
            df_emergencial['Pedido']
            .astype(str)
            .str.replace(r'\.0$', '', regex=True)
            .str.strip()
        )
        
        unidade_cruzada = chave_emerg.map(mapa_unidades)
        
        df_emergencial['Unidade'] = (
            df_emergencial['Unidade']
            .fillna(unidade_cruzada)
            .fillna(df_emergencial.get('cidade', 'INCONFIDENCIA LOCADORA'))
        )

    return df_emergencial

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
upload_unidades = None
upload_pedidos = None
upload_faturamento = None
upload_controle = None
upload_emergencial = None

if st.session_state.get("perfil") == "admin":
    with st.sidebar:
        st.markdown("---")
        st.header("📂 Carregamento de Bases")

        # Uploads padrão
        upload_unidades = st.file_uploader(
            "1. Listagem de Unidades (.xlsx/.csv)", 
            type=["xlsx", "xls", "csv"], 
            key="up_unid"
        )
        upload_pedidos = st.file_uploader(
            "2. Relatório de Pedidos / Itens (.xlsx/.csv)", 
            type=["xlsx", "xls", "csv"], 
            key="up_ped"
        )
        upload_faturamento = st.file_uploader(
            "3. Base de Faturamento / BI (.xlsx/.csv)", 
            type=["xlsx", "xls", "csv"], 
            key="up_fat"
        )
        upload_controle = st.file_uploader(
            "4. Controle de Emissão (.xlsx/.csv)", 
            type=["xlsx", "xls", "csv"], 
            key="up_ctrl"
        )

        # 5ª Base Oficial: Emergencial / NFs Pré-Corte (Ativa até o fim do mês)
        upload_emergencial = st.file_uploader(
            "5. Base Emergencial (NFs Pré-Corte)", 
            type=["xlsx", "xls", "csv"], 
            key="up_emerg",
            help="Suba aqui a planilha com os números de NF emitidas antes do corte."
        )

        if st.button("💾 Salvar Bases na Nuvem", use_container_width=True):
            arquivos_salvos = []
            if upload_unidades is not None:
                ext_unid = os.path.splitext(upload_unidades.name)[1].lower()
                if ext_unid not in [".xlsx", ".xls", ".csv"]:
                    ext_unid = ".xlsx"
                nome_listagem = f"listagem_unidades{ext_unid}"
                with open(nome_listagem, "wb") as f:
                    f.write(upload_unidades.getbuffer())
                arquivos_salvos.append("1. Listagem de Unidades")

            if upload_pedidos is not None:
                with open("acompanhamento.xlsx", "wb") as f:
                    f.write(upload_pedidos.getbuffer())
                arquivos_salvos.append("2. Relatório de Pedidos / Itens")

            if upload_faturamento is not None:
                with open("dados_bi.xlsx", "wb") as f:
                    f.write(upload_faturamento.getbuffer())
                arquivos_salvos.append("3. Base de Faturamento / BI")

            if upload_controle is not None:
                with open("unidades.xlsx", "wb") as f:
                    f.write(upload_controle.getbuffer())
                arquivos_salvos.append("4. Controle de Emissão")

            if upload_emergencial is not None:
                ext_nf = os.path.splitext(upload_emergencial.name)[1].lower()
                if ext_nf not in [".xlsx", ".xls", ".csv"]:
                    ext_nf = ".xlsx"
                nome_nf_emerg = f"nf_emergencial{ext_nf}"
                with open(nome_nf_emerg, "wb") as f:
                    f.write(upload_emergencial.getbuffer())
                arquivos_salvos.append("5. Base Emergencial (NFs Pré-Corte)")

            if arquivos_salvos:
                st.cache_data.clear()
                st.sidebar.success(f"✅ Bases salvas e atualizadas na nuvem: {', '.join(arquivos_salvos)}!")
                st.rerun()
            else:
                st.sidebar.warning("⚠️ Selecione pelo menos uma planilha para salvar.")

        st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
        if st.button("🗑️ Limpar / Resetar Base de Dados", use_container_width=True, type="secondary"):
            arquivos_para_remover = [
                'acompanhamento.xlsx',
                'dados_bi.xlsx',
                'unidades.xlsx',
                'listagem_unidades.xlsx',
                'listagem_unidades.xls',
                'listagem_unidades.csv',
                'nf_emergencial.xlsx',
                'nf_emergencial.xls',
                'nf_emergencial.csv'
            ]
            removidos = []
            for fpath in arquivos_para_remover:
                if os.path.exists(fpath):
                    try:
                        os.remove(fpath)
                        removidos.append(fpath)
                    except Exception:
                        pass
            st.cache_data.clear()
            st.sidebar.success("✅ Base de dados limpa com sucesso!")
            st.rerun()

# Aliases de compatibilidade interna
file_listagem_unidades = upload_unidades
uploaded_nf_emergencial = upload_emergencial
up_acomp = upload_pedidos
up_bi = upload_faturamento
up_unidades = upload_controle

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Período do Relatório")

# Data de corte fixa para produção (a partir de 21/10/2026)
DATA_CORTE = datetime.date(2026, 10, 21)

# Filtro de datas restrito à data de corte (impede seleção anterior a 21/10/2026)
data_inicio = st.sidebar.date_input("Data Inicial", value=DATA_CORTE, min_value=DATA_CORTE, format="DD/MM/YYYY")
data_fim = st.sidebar.date_input("Data Final", value=None, min_value=DATA_CORTE, format="DD/MM/YYYY")

# Mensagem de atualização logo abaixo das datas, com fonte maior
data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
st.sidebar.markdown(
    f"<div style='margin-top: 8px; margin-bottom: 8px; color: #FFFFFF; font-size: 14px; font-weight: 500; text-align: center;'>"
    f"🔄 Atualizado {data_atual} por Filipe"
    f"</div>", 
    unsafe_allow_html=True
)

st.sidebar.markdown("---")

# ---------------------------------------------------------
# CARREGAMENTO DOS DADOS (MODO SERVIDOR / NUVEM)
# ---------------------------------------------------------
file_acompanhamento = upload_pedidos if upload_pedidos is not None else 'acompanhamento.xlsx'
file_bi = upload_faturamento if upload_faturamento is not None else 'dados_bi.xlsx'
file_unidades = upload_controle if upload_controle is not None else 'unidades.xlsx'
file_listagem_unidades_fixo = 'listagem_unidades.xlsx' if os.path.exists('listagem_unidades.xlsx') else ('listagem_unidades.csv' if os.path.exists('listagem_unidades.csv') else ('listagem_unidades.xls' if os.path.exists('listagem_unidades.xls') else None))


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

    df_merged['Pedido_Clean'] = sanitizar_chave_pedido(df_merged['Pedido'])
    df_unidades['Pedido_Clean'] = sanitizar_chave_pedido(ped_col_data)
    df_unidades['Unidade'] = uni_col_data
    df_unidades['Unidade_Clean'] = df_unidades['Unidade'].apply(limpar_texto_corrompido)
    df_unidades['Unidade_Extraida'] = uni_col_data

    # Remove vazios e duplicatas, faz o merge
    df_unidades_unique = df_unidades[df_unidades['Pedido_Clean'] != ''].drop_duplicates(subset=['Pedido_Clean'])

    # Merge 2 (Base + Unidades)
    if 'Unidade_Extraida' in df_merged.columns:
        df_merged = df_merged.drop(columns=['Unidade_Extraida'])

    df_merged = pd.merge(df_merged, df_unidades_unique[['Pedido_Clean', 'Unidade_Extraida']], on='Pedido_Clean', how='left')

    # Blindagem adicional com cruzar_unidades_blindado para garantia total
    df_merged = cruzar_unidades_blindado(df_merged, df_unidades_unique)

    # Define Unidade Final
    df_merged['Unidade_Final'] = df_merged.get('Unidade_Extraida', pd.Series(dtype='object')).fillna(df_merged.get('Unidade', pd.Series(dtype='object')))
    df_merged['Unidade_Final'] = df_merged['Unidade_Final'].fillna(df_merged['cidade'])
    df_merged['Unidade_Final'] = df_merged['Unidade_Final'].fillna('Não Informada').replace({'': 'Não Informada', 'nan': 'Não Informada', 'None': 'Não Informada'})

    # Sincronizar colunas de unidade para os filtros e visões
    df_merged["UNIDADE"] = df_merged["Unidade_Final"]
    df_merged["Unidade"] = df_merged["Unidade_Final"]
    df_merged["Unidade_Real"] = df_merged["Unidade_Final"]

    # Também enriquece df_bi com Pedido e Unidade_Final para os itens da Aba 2
    df_bi_merged = pd.merge(df_bi, df_merged[['NF_Clean', 'Pedido', 'Pedido_Clean', 'Unidade_Final']].drop_duplicates(subset=['NF_Clean']), on='NF_Clean', how='left')
    df_bi_merged['UNIDADE'] = df_bi_merged['Unidade_Final'].fillna('Não Informada')

    return df_merged, df_bi_merged, df_unidades_unique, col_ped, 'Unidade_Extraida'

def base_valida(b):
    if b is None:
        return False
    if isinstance(b, str):
        return os.path.exists(b)
    return hasattr(b, "read") or hasattr(b, "seek")

# Verifica se os arquivos existem no servidor ou foram enviados via upload
if base_valida(file_acompanhamento) and base_valida(file_bi) and base_valida(file_unidades):
    def get_file_mtime(f):
        if isinstance(f, str) and os.path.exists(f):
            return os.path.getmtime(f)
        return getattr(f, "size", getattr(f, "name", 0))

    mtimes = (
        get_file_mtime(file_acompanhamento),
        get_file_mtime(file_bi),
        get_file_mtime(file_unidades)
    )
    df_merged, df_bi_raw, df_unidades_unique, col_ped, col_uni = process_data(file_acompanhamento, file_bi, file_unidades, _mtime=mtimes)
    df_unidades = df_unidades_unique

    # ---------------------------------------------------------
    # PROCESSAMENTO DE NFs EMERGENCIAIS PRÉ-CORTE
    # ---------------------------------------------------------
    df_nf_emergencial = None
    file_nf_emergencial_fixo = 'nf_emergencial.xlsx' if os.path.exists('nf_emergencial.xlsx') else ('nf_emergencial.csv' if os.path.exists('nf_emergencial.csv') else ('nf_emergencial.xls' if os.path.exists('nf_emergencial.xls') else None))
    arquivo_nf_emerg_alvo = uploaded_nf_emergencial if uploaded_nf_emergencial is not None else file_nf_emergencial_fixo

    nfs_extras = set()
    pedidos_extras = set()

    if arquivo_nf_emerg_alvo is not None:
        try:
            # Leitura da planilha emergencial
            if hasattr(arquivo_nf_emerg_alvo, "name") and arquivo_nf_emerg_alvo.name.endswith(".csv"):
                df_emergencial = pd.read_csv(arquivo_nf_emerg_alvo, sep=None, engine='python')
            else:
                df_emergencial = load_uploaded_file(arquivo_nf_emerg_alvo)
            
            if df_emergencial is not None and not df_emergencial.empty:
                if isinstance(df_emergencial.columns, pd.MultiIndex):
                    df_emergencial.columns = df_emergencial.columns.get_level_values(-1)

                # APLICA A REGRA ESPECÍFICA DE EXCEÇÃO
                df_emergencial = aplicar_regra_excecao(df_emergencial)
                
                # Converte o Pedido para string sem decimais
                if 'Pedido' in df_emergencial.columns:
                    df_emergencial['Pedido'] = df_emergencial['Pedido'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                # Processa coluna 'Local de entrega' e extrai Pedido e Unidade se aplicável
                df_emergencial = processar_planilha_emergencial(df_emergencial, df_unidades=df_unidades_unique)
                df_nf_emergencial = df_emergencial
                
                # Identifica a coluna que contém o número da NF na planilha emergencial
                cols_nf_encontradas = [c for c in df_nf_emergencial.columns if str(c).strip().upper() == 'NF']
                if not cols_nf_encontradas:
                    cols_nf_encontradas = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['nota', 'nf', 'n_nota', 'numero_nf', 'numero_nota', 'doc', 'fatura'])]
                
                col_nf_emerg = cols_nf_encontradas[0] if cols_nf_encontradas else df_nf_emergencial.columns[0]
                col_nf = col_nf_emerg
                
                # Limpa possíveis prefixos como 'NF-' ou espaços para bater certinho
                nfs_extras_clean = (
                    df_nf_emergencial[col_nf_emerg]
                    .dropna()
                    .astype(str)
                    .str.replace('NF-', '', case=False)
                    .str.replace(r'\D', '', regex=True)
                    .str.lstrip('0')
                    .str.strip()
                )
                nfs_extras_raw = (
                    df_nf_emergencial[col_nf_emerg]
                    .dropna()
                    .astype(str)
                    .str.replace('NF-', '', case=False)
                    .str.strip()
                )
                nfs_extras = set(nfs_extras_clean[nfs_extras_clean != '']).union(set(nfs_extras_raw[nfs_extras_raw != '']))
                
                # Busca por coluna de Pedido (se existir na planilha)
                cols_ped = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['pedido', 'ped', 'ordem'])]
                if cols_ped:
                    vals_ped_raw = df_nf_emergencial[cols_ped[0]].dropna().astype(str).str.strip()
                    vals_ped_limpos = vals_ped_raw.str.replace(r'\D', '', regex=True).str.lstrip('0')
                    pedidos_extras = set(vals_ped_raw[vals_ped_raw != '']).union(set(vals_ped_limpos[vals_ped_limpos != '']))

                st.sidebar.success(f"Notas extras carregadas: {len(df_nf_emergencial)}")
        except Exception as e:
            st.sidebar.error(f"Erro ao processar exceções: {e}")

    # Aliases de compatibilidade interna
    nfs_emergenciais = nfs_extras
    pedidos_emergenciais = pedidos_extras

    # --- UNIFICAÇÃO DO FATURAMENTO (Oficial + Emergencial) ---
    df_fat_oficial = df_bi_raw.copy() if 'df_bi_raw' in locals() and df_bi_raw is not None else None
    df_fat_emerg = df_nf_emergencial.copy() if df_nf_emergencial is not None else None

    df_faturamento_total = pd.DataFrame()

    # Se tem base oficial, usa ela
    if df_fat_oficial is not None and not df_fat_oficial.empty:
        df_faturamento_total = df_fat_oficial.copy()

    # Se tem base emergencial, trata a coluna 'Local de entrega' e junta à oficial
    if df_fat_emerg is not None and not df_fat_emerg.empty:
        df_emerg = aplicar_regra_excecao(df_fat_emerg)

        # 1. GARANTIR QUE df_unidades NÃO SOFRA FILTRO DE DATA
        # Padroniza a chave de pedidos em ambas as bases para string pura sem ponto decimal
        if df_unidades is not None and not df_unidades.empty and 'Pedido' in df_emerg.columns:
            df_u = df_unidades.copy()
            
            # Coluna B = Pedido (índice 1), Coluna D = Unidade (índice 3)
            col_ped_u = df_u.columns[1] if 'Pedido' not in df_u.columns else 'Pedido'
            col_nome_u = df_u.columns[3] if len(df_u.columns) > 3 and 'Unidade' not in df_u.columns else ('Unidade' if 'Unidade' in df_u.columns else df_u.columns[-1])
            
            # Limpa para string pura (remove '.0' e espaços)
            df_u['chave_limpa'] = (
                df_u[col_ped_u]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.strip()
            )
            
            mapa_unidades = (
                df_u[df_u['chave_limpa'] != '']
                .drop_duplicates(subset=['chave_limpa'])
                .set_index('chave_limpa')[col_nome_u]
                .to_dict()
            )
            
            # Chave limpa da emergencial
            chave_emerg = (
                df_emerg['Pedido']
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.strip()
            )
            
            # Busca na planilha de Unidades oficial se faltou alguma
            unidade_cruzada = chave_emerg.map(mapa_unidades)
            
            df_emerg['Unidade'] = (
                df_emerg['Unidade']
                .fillna(unidade_cruzada)
                .fillna(df_emerg.get('cidade', 'INCONFIDENCIA LOCADORA'))
            )

        # Atualiza a referência de df_nf_emergencial já com as Unidades preenchidas
        df_nf_emergencial = df_emerg.copy()

        # Junta as duas bases de faturamento em uma só
        df_faturamento_total = pd.concat([df_faturamento_total, df_emerg], ignore_index=True)

    # --- CRUZAMENTO DAS UNIDADES COM A BASE UNIFICADA ---
    if df_unidades is not None and not df_faturamento_total.empty:
        # Identifica colunas de Pedido e Unidade na Listagem de Unidades
        # Baseado na sua planilha: Coluna B = Pedido (índice 1), Coluna D = Unidade (índice 3)
        col_ped_unid = next((c for c in df_unidades.columns if 'pedido' in str(c).lower()), df_unidades.columns[1] if len(df_unidades.columns) > 1 else df_unidades.columns[0])
        col_nome_unid = next((c for c in df_unidades.columns if 'unidade' in str(c).lower()), df_unidades.columns[3] if len(df_unidades.columns) > 3 else df_unidades.columns[-1])
        
        # Limpa as chaves para bater string com string
        df_unidades['chave_cruzamento'] = df_unidades[col_ped_unid].apply(limpar_id)
        
        # Se na tabela de faturamento a coluna Pedido existir, padroniza
        col_ped_fat = next((c for c in df_faturamento_total.columns if 'pedido' in str(c).lower()), 'Pedido')
        if col_ped_fat in df_faturamento_total.columns:
            df_faturamento_total['chave_cruzamento'] = df_faturamento_total[col_ped_fat].apply(limpar_id)
        else:
            df_faturamento_total['chave_cruzamento'] = ""
        
        # Cria o dicionário { '2406179': 'BANRISUL SUREG ALTO URUGUAI AG UPF CAMPUS I 21547', ... }
        mapa_unidades = (
            df_unidades[df_unidades['chave_cruzamento'] != '']
            .drop_duplicates(subset=['chave_cruzamento'])
            .set_index('chave_cruzamento')[col_nome_unid]
            .to_dict()
        )
        
        # Mapeia a Unidade correta para todos os pedidos
        unidades_encontradas = df_faturamento_total['chave_cruzamento'].map(mapa_unidades)
        
        # Se encontrou no mapa de unidades, usa o nome real (Banrisul, PRF, etc.)
        # Se não encontrou, mantém o valor original ou razão social
        df_faturamento_total['Unidade'] = unidades_encontradas.fillna(df_faturamento_total.get('Unidade', '-'))
        
        # Limpa a coluna auxiliar
        df_faturamento_total.drop(columns=['chave_cruzamento'], inplace=True, errors='ignore')

        # Atualiza a base de itens de faturamento (df_bi_raw) com a visão unificada
        if not df_faturamento_total.empty:
            df_bi_raw = df_faturamento_total

    # --- INTEGRAÇÃO COM OS SEUS PEDIDOS / FATURAMENTO ---
    # 1. Pega os números de NF da planilha normal de faturamento (BI) e da base unificada
    nfs_validas = set()
    if 'NF_Clean' in df_bi_raw.columns:
        nfs_validas.update(df_bi_raw['NF_Clean'].dropna().astype(str).str.strip().unique())
    if 'n_nota_obs' in df_bi_raw.columns:
        nfs_validas.update(df_bi_raw['n_nota_obs'].dropna().astype(str).str.replace('NF-', '', case=False).str.strip().unique())
        nfs_validas.update(df_bi_raw['n_nota_obs'].dropna().astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0').unique())

    if not df_faturamento_total.empty:
        for c_nf in [c for c in df_faturamento_total.columns if any(p in str(c).lower() for p in ['nf', 'nota'])]:
            nfs_validas.update(df_faturamento_total[c_nf].dropna().astype(str).str.strip().unique())
            nfs_validas.update(df_faturamento_total[c_nf].dropna().astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0').unique())

    # 2. Se a planilha emergencial foi anexada, adiciona as NFs dela ao mesmo conjunto combinado
    if nfs_extras:
        nfs_validas.update(nfs_extras)

    # Opção A: Mesclar via LEFT JOIN mantendo todos os pedidos originais intactos
    if df_nf_emergencial is not None and not df_nf_emergencial.empty:
        try:
            cols_chave_cand = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['pedido', 'ped', 'ordem'])]
            chave = cols_chave_cand[0] if cols_chave_cand else None

            if chave and chave in df_nf_emergencial.columns and col_nf and col_nf in df_nf_emergencial.columns:
                df_merged['chave_ped_join'] = sanitizar_chave_pedido(df_merged['Pedido'])
                df_emerg_sub = df_nf_emergencial.dropna(subset=[chave]).copy()
                df_emerg_sub['chave_ped_join'] = sanitizar_chave_pedido(df_emerg_sub[chave])
                df_emerg_sub = df_emerg_sub[df_emerg_sub['chave_ped_join'] != ''].drop_duplicates(subset=['chave_ped_join'])

                cols_merge_emerg = ['chave_ped_join', col_nf]
                tem_unid_emerg = 'Unidade' in df_emerg_sub.columns
                if tem_unid_emerg:
                    cols_merge_emerg.append('Unidade')

                df_emerg_sub_slice = df_emerg_sub[cols_merge_emerg].copy()
                df_emerg_sub_slice.rename(columns={col_nf: 'NF_Emergencial'}, inplace=True)
                if tem_unid_emerg:
                    df_emerg_sub_slice.rename(columns={'Unidade': 'Unidade_Emergencial'}, inplace=True)

                df_merged = df_merged.merge(
                    df_emerg_sub_slice,
                    on='chave_ped_join',
                    how='left'
                )

                if 'NF_Emergencial' in df_merged.columns:
                    mask_tem_emerg = df_merged['NF_Emergencial'].notna() & (df_merged['NF_Emergencial'].astype(str).str.strip() != '')
                    mask_vazia = df_merged['Nota'].isna() | df_merged['Nota'].astype(str).str.strip().isin(['', '0', 'nan', 'None', '-'])
                    df_merged.loc[mask_vazia & mask_tem_emerg, 'Nota'] = df_merged['NF_Emergencial']
                    df_merged['NF_Clean'] = df_merged['Nota'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
                    df_merged.drop(columns=['NF_Emergencial'], inplace=True)

                if 'Unidade_Emergencial' in df_merged.columns:
                    mask_unid_emerg = df_merged['Unidade_Emergencial'].notna() & (df_merged['Unidade_Emergencial'].astype(str).str.strip() != '')
                    mask_unid_vazia = df_merged['Unidade_Final'].isna() | df_merged['Unidade_Final'].astype(str).str.strip().isin(['', 'Não Informada', 'nan', 'None', '-'])
                    df_merged.loc[mask_unid_vazia & mask_unid_emerg, 'Unidade_Final'] = df_merged['Unidade_Emergencial']
                    df_merged['UNIDADE'] = df_merged['Unidade_Final']
                    df_merged['Unidade'] = df_merged['Unidade_Final']
                    df_merged['Unidade_Real'] = df_merged['Unidade_Final']
                    df_merged.drop(columns=['Unidade_Emergencial'], inplace=True)

                df_merged.drop(columns=['chave_ped_join'], inplace=True, errors='ignore')
        except Exception as e_merge:
            st.sidebar.warning(f"Aviso no cruzamento emergencial por chave: {e_merge}")

    # 3. Cruzamento padronizado: marca como Faturado tudo que está no conjunto combinado (Normal + Emergencial)
    df_merged['NF_limpa'] = df_merged['Nota'].astype(str).str.replace('NF-', '', case=False).str.strip()
    df_merged['NF_limpa_dig'] = df_merged['Nota'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
    
    mask_faturados_combinado = (
        df_merged['NF_limpa'].isin(nfs_validas) | 
        df_merged['NF_limpa_dig'].isin(nfs_validas) |
        df_merged['NF_Clean'].isin(nfs_validas)
    )

    df_merged.loc[mask_faturados_combinado, 'Status_Faturamento'] = 'Faturado'
    df_merged.loc[mask_faturados_combinado, 'Status'] = 'Faturado'

    df_merged.drop(columns=['NF_limpa', 'NF_limpa_dig'], inplace=True, errors='ignore')

    # 4. Liberação da data de corte para registros do período ou faturados no conjunto combinado
    serie_dt_corte = pd.to_datetime(df_merged['Dt_inser'], dayfirst=True, errors='coerce').dt.date
    cond_corte = (serie_dt_corte >= DATA_CORTE) | mask_faturados_combinado

    if pedidos_extras:
        mask_ped = df_merged['Pedido_Clean'].isin(pedidos_extras) | df_merged['Pedido'].astype(str).str.strip().isin(pedidos_extras)
        cond_corte = cond_corte | mask_ped

    df_merged = df_merged[cond_corte].copy()

    # 2. Concatenação de eventuais linhas adicionais da planilha emergencial (pd.concat)
    if df_nf_emergencial is not None and not df_nf_emergencial.empty:
        try:
            df_novas_linhas = pd.DataFrame()
            
            # Pedido
            cols_ped = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['pedido', 'ped', 'ordem'])]
            if cols_ped:
                df_novas_linhas['Pedido'] = df_nf_emergencial[cols_ped[0]].astype(str).str.strip()
            elif col_nf and col_nf in df_nf_emergencial.columns:
                df_novas_linhas['Pedido'] = "NF-" + df_nf_emergencial[col_nf].astype(str).str.strip()
            else:
                df_novas_linhas['Pedido'] = [f"EMERG-{i+1}" for i in range(len(df_nf_emergencial))]
            
            df_novas_linhas['Pedido_Clean'] = df_novas_linhas['Pedido'].str.replace(r'\D', '', regex=True).str.lstrip('0')
            
            # Nota Fiscal
            if col_nf and col_nf in df_nf_emergencial.columns:
                df_novas_linhas['Nota'] = df_nf_emergencial[col_nf].astype(str).str.strip()
            else:
                df_novas_linhas['Nota'] = df_novas_linhas['Pedido']
            df_novas_linhas['NF_Clean'] = df_novas_linhas['Nota'].str.replace(r'\D', '', regex=True).str.lstrip('0')
            
            # Unidade
            cols_uni = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['unidade', 'cliente', 'razao social', 'nome'])]
            if cols_uni:
                df_novas_linhas['UNIDADE'] = df_nf_emergencial[cols_uni[0]].astype(str).str.strip()
            else:
                df_novas_linhas['UNIDADE'] = "Ajuste Emergencial"
            
            df_novas_linhas['Unidade_Final'] = df_novas_linhas['UNIDADE']
            df_novas_linhas['Unidade'] = df_novas_linhas['UNIDADE']
            df_novas_linhas['Unidade_Real'] = df_novas_linhas['UNIDADE']
            df_novas_linhas['cidade'] = df_novas_linhas['UNIDADE']
            df_novas_linhas['Unidade_Clean'] = df_novas_linhas['UNIDADE'].apply(limpar_texto_corrompido)
            
            # Data de Inserção
            cols_dt = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['data', 'dt', 'emissao'])]
            if cols_dt:
                df_novas_linhas['Dt_inser'] = df_nf_emergencial[cols_dt[0]].astype(str).str.strip()
            else:
                df_novas_linhas['Dt_inser'] = DATA_CORTE.strftime("%d/%m/%Y")
            df_novas_linhas['dt_nota_hist'] = df_novas_linhas['Dt_inser']
            
            # Status e Situação
            df_novas_linhas['Status'] = 'Faturado'
            df_novas_linhas['Status_Faturamento'] = 'Faturado'
            df_novas_linhas['Tipo'] = 'Emergencial'
            
            # Valor / Faturamento
            cols_vlr = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['valor', 'total', 'vendas', 'preco'])]
            if cols_vlr:
                df_novas_linhas['Vendas Total'] = pd.to_numeric(df_nf_emergencial[cols_vlr[0]], errors='coerce').fillna(0.0)
            else:
                df_novas_linhas['Vendas Total'] = 0.0
            
            cols_qtd = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['qtd', 'quantidade', 'quant'])]
            if cols_qtd:
                df_novas_linhas['QTD Total'] = pd.to_numeric(df_nf_emergencial[cols_qtd[0]], errors='coerce').fillna(0)
            else:
                df_novas_linhas['QTD Total'] = 1

            # Mantém as colunas padronizadas do seu relatório
            colunas_padrao = ['Pedido', 'Unidade', 'Data de Inserção', 'Valor (R$)']
            df_novas_linhas['Data de Inserção'] = df_novas_linhas['Dt_inser']
            df_novas_linhas['Valor (R$)'] = df_novas_linhas['Vendas Total']

            # Se for concatenar diretamente com a tabela principal de faturados:
            df_merged = pd.concat([df_merged, df_novas_linhas], ignore_index=True)

            # Remove duplicatas mantendo a versão com dados completos
            if 'Pedido' in df_merged.columns:
                df_merged = df_merged.drop_duplicates(subset=['Pedido'], keep='last')
        except Exception as err_concat:
            st.sidebar.warning(f"Aviso ao conciliar novas linhas da planilha emergencial: {err_concat}")
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

# Filtro de Período (com garantia da data de corte >= 21/10/2026, exceto NFs/Pedidos emergenciais)
data_ini_efetiva = max(data_inicio, DATA_CORTE) if data_inicio else DATA_CORTE
serie_dt = pd.to_datetime(df_filtrado['Dt_inser'], dayfirst=True, errors='coerce').dt.date

cond_inicio = serie_dt >= data_ini_efetiva
if nfs_emergenciais:
    cond_inicio = cond_inicio | df_filtrado['NF_Clean'].isin(nfs_emergenciais)
if pedidos_emergenciais:
    cond_inicio = cond_inicio | df_filtrado['Pedido_Clean'].isin(pedidos_emergenciais)

if data_fim:
    cond_fim = serie_dt <= data_fim
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
# 6. ABAS: VISÃO POR UNIDADE, PEDIDOS E CONTROLE DE EMISSÃO
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 Visão por Unidade", "📦 Pedidos e Itens", "Controle de Emissão de pedidos"])

# ----------------- ABA 1: VISÃO POR UNIDADE -----------------
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

    df_tabela['Data_Formatada'] = pd.to_datetime(df_tabela['Data de Inserção'], errors='coerce', dayfirst=True)

    data_ini_efetiva = max(data_inicio, DATA_CORTE) if data_inicio else DATA_CORTE
    cond_tab = df_tabela['Data_Formatada'].dt.date >= data_ini_efetiva
    if nfs_emergenciais or pedidos_emergenciais:
        peds_liberados = df_merged['Pedido'].unique()
        cond_tab = cond_tab | df_tabela['Pedido'].isin(peds_liberados)

    if data_fim:
        cond_tab_fim = df_tabela['Data_Formatada'].dt.date <= data_fim
        if nfs_emergenciais or pedidos_emergenciais:
            cond_tab_fim = cond_tab_fim | df_tabela['Pedido'].isin(peds_liberados)
        df_tabela = df_tabela[cond_tab & cond_tab_fim]
    else:
        df_tabela = df_tabela[cond_tab]

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
with tab2:
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


# ----------------- ABA 3: CONTROLE DE EMISSÃO DE PEDIDOS -----------------
with tab3:
    st.markdown("### Controle de Emissão de Pedidos - Visão Trimestral")
    st.caption("Consolidação trimestral automática (últimos 3 meses até a data presente) cruzando a listagem mestre de unidades com os relatórios de pedidos do portal.")

    arquivo_listagem_alvo = file_listagem_unidades if file_listagem_unidades is not None else file_listagem_unidades_fixo
    arquivo_pedidos_alvo = file_unidades if file_unidades is not None else ('unidades.xlsx' if os.path.exists('unidades.xlsx') else None)

    if not arquivo_listagem_alvo:
        st.warning("⚠️ Por favor, faça o upload da 'Listagem de Unidades' na área administrativa da barra lateral para visualizar o controle trimestral.")
    elif not arquivo_pedidos_alvo:
        st.warning("⚠️ O relatório de pedidos do portal ('unidades.xlsx') não foi encontrado no servidor.")
    else:
        try:
            # 1. Carrega Cadastro Mestre de Unidades (4ª Planilha)
            df_cad_raw = load_uploaded_file(arquivo_listagem_alvo)
            col_nome_cad = [c for c in df_cad_raw.columns if str(c).strip().lower() in ['nome', 'unidade', 'razao social']][0] if any(str(c).strip().lower() in ['nome', 'unidade', 'razao social'] for c in df_cad_raw.columns) else df_cad_raw.columns[0]
            col_ativo_cad = [c for c in df_cad_raw.columns if str(c).strip().lower() in ['ativo', 'situacao', 'status']][0] if any(str(c).strip().lower() in ['ativo', 'situacao', 'status'] for c in df_cad_raw.columns) else None

            cols_mestre = [col_nome_cad]
            if col_ativo_cad:
                cols_mestre.append(col_ativo_cad)
            
            df_unidades = df_cad_raw[cols_mestre].drop_duplicates(subset=[col_nome_cad]).copy()
            df_unidades.rename(columns={col_nome_cad: 'Unidade'}, inplace=True)
            if col_ativo_cad:
                df_unidades.rename(columns={col_ativo_cad: 'ATIVO'}, inplace=True)
                df_unidades['ATIVO'] = df_unidades['ATIVO'].astype(str).str.strip().str.upper()
            else:
                df_unidades['ATIVO'] = 'SIM'

            # Aplicação nas colunas de unidades com sanitização avançada
            df_unidades['Unidade_Clean'] = df_unidades['Unidade'].apply(limpar_texto_corrompido)

            # 2. Carrega Relatório de Pedidos do Portal (3ª Planilha)
            df_ped_raw = load_uploaded_file(arquivo_pedidos_alvo)
            if isinstance(df_ped_raw.columns, pd.MultiIndex):
                df_ped_raw.columns = df_ped_raw.columns.get_level_values(-1)

            col_data_ped = [c for c in df_ped_raw.columns if 'data' in str(c).lower()][0]
            col_uni_ped = [c for c in df_ped_raw.columns if 'unidade' in str(c).lower()][0]
            col_vlr_ped = [c for c in df_ped_raw.columns if any(p in str(c).lower() for p in ['valor total', 'vendas total', 'total'])][0]

            df_pedidos = df_ped_raw[[col_data_ped, col_uni_ped, col_vlr_ped]].copy()
            df_pedidos.rename(columns={col_uni_ped: 'Unidade'}, inplace=True)
            df_pedidos['Unidade_Clean'] = df_pedidos['Unidade'].apply(limpar_texto_corrompido)
            df_pedidos['Valor'] = pd.to_numeric(df_pedidos[col_vlr_ped], errors='coerce').fillna(0.0)

            # Conversão robusta de data para datetime
            dt_parsed = pd.to_datetime(df_pedidos[col_data_ped], errors='coerce', dayfirst=True)
            df_pedidos['Data_Obj'] = dt_parsed.dt.date

            # Janela fixa trimestral dinâmica (últimos 3 meses até a data presente)
            hoje = datetime.date.today()
            ano_inicio = hoje.year
            mes_inicio = hoje.month - 2
            while mes_inicio <= 0:
                mes_inicio += 12
                ano_inicio -= 1
            data_inicio_trimestre = datetime.date(ano_inicio, mes_inicio, 1)
            data_fim_trimestre = hoje

            # Aplicação do período fixo trimestral (ignora filtro de data da barra lateral)
            df_pedidos = df_pedidos[(df_pedidos['Data_Obj'] >= data_inicio_trimestre) & (df_pedidos['Data_Obj'] <= data_fim_trimestre)]

            # Mapeamento do mês em português
            dt_filtrado = pd.to_datetime(df_pedidos[col_data_ped], errors='coerce', dayfirst=True)
            meses_pt = {
                1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
                5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
                9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
            }
            df_pedidos['AnoMes'] = dt_filtrado.dt.to_period('M')
            df_pedidos['Mes'] = dt_filtrado.dt.month.map(meses_pt)

            meses_ordenados = df_pedidos.dropna(subset=['AnoMes', 'Mes']).sort_values('AnoMes')['Mes'].unique().tolist()

            if not meses_ordenados:
                st.warning(f"⚠️ Nenhum pedido registrado no período trimestral ({data_inicio_trimestre.strftime('%d/%m/%Y')} a {data_fim_trimestre.strftime('%d/%m/%Y')}).")
                meses_ordenados = ["Período Selecionado"]
                pivot = pd.DataFrame(columns=['Unidade_Clean', 'Período Selecionado'])
            else:
                # 3. Pivot por Unidade e Mês
                pivot = df_pedidos.groupby(['Unidade_Clean', 'Mes'])['Valor'].sum().unstack(fill_value=0.0)
                for m in meses_ordenados:
                    if m not in pivot.columns:
                        pivot[m] = 0.0
                pivot = pivot[meses_ordenados].reset_index()

            # 4. Cruzamento Mestre <-> Pedidos
            matriz = pd.merge(df_unidades, pivot, on='Unidade_Clean', how='left')
            for m in meses_ordenados:
                matriz[m] = matriz[m].fillna(0.0)

            # 5. Formatação Exata ("Sem Compra" quando não houver emissão)
            matriz_exibicao = matriz[['Unidade', 'ATIVO'] + meses_ordenados].copy()
            for m in meses_ordenados:
                matriz_exibicao[m] = matriz_exibicao[m].apply(
                    lambda v: f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if isinstance(v, (int, float)) and v > 0 else "Sem Compra"
                )

            # Métricas e KPIs no topo da Aba
            total_cad = len(matriz_exibicao)
            total_ativas = len(matriz_exibicao[matriz_exibicao['ATIVO'] == 'SIM'])
            unidades_com_compra = (matriz[meses_ordenados].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) > 0).sum()
            unidades_sem_compra = total_cad - unidades_com_compra

            c_kp1, c_kp2, c_kp3, c_kp4 = st.columns(4)
            c_kp1.metric("Unidades Cadastradas", f"{total_cad:,}".replace(",", "."))
            c_kp2.metric("Unidades Ativas", f"{total_ativas:,}".replace(",", "."))
            c_kp3.metric("Com Compra no Período", f"{unidades_com_compra:,}".replace(",", "."))
            c_kp4.metric("Sem Compra no Período", f"{unidades_sem_compra:,}".replace(",", "."))

            st.markdown("---")
            st.markdown("#### 🔍 Filtros de Consulta")
            col_f1, col_f2 = st.columns(2)
            
            df_controle = matriz_exibicao.copy()

            with col_f1:
                # Filtro por Status ATIVO
                opcoes_ativo = ["TODOS"] + list(df_controle['ATIVO'].dropna().unique()) if 'ATIVO' in df_controle.columns else ["TODOS"]
                filtro_ativo = st.selectbox("Filtrar por Status (ATIVO)", opcoes_ativo)
                
            with col_f2:
                # Busca rápida por texto na Unidade
                busca_unidade = st.text_input("🔍 Buscar Unidade por Nome")

            # Aplicação dos filtros solicitados
            df_filtrado = df_controle.copy()

            if filtro_ativo != "TODOS" and 'ATIVO' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['ATIVO'] == filtro_ativo]

            if busca_unidade and 'Unidade' in df_filtrado.columns:
                busca_norm = limpar_texto_corrompido(busca_unidade)
                df_filtrado = df_filtrado[
                    df_filtrado['Unidade'].apply(limpar_texto_corrompido).str.contains(busca_norm, case=False, na=False) |
                    df_filtrado['Unidade'].astype(str).str.contains(busca_unidade, case=False, na=False)
                ]

            st.markdown(f"Exibindo {len(df_filtrado)} unidades encontradas:")
            st.dataframe(df_filtrado, use_container_width=True, height=600, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Botão de Download em Excel
            buf_matriz = io.BytesIO()
            with pd.ExcelWriter(buf_matriz, engine='xlsxwriter') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Controle_Emissao')

            st.download_button(
                label="📥 Baixar Controle de Emissão Trimestral (Excel)",
                data=buf_matriz.getvalue(),
                file_name="Controle_Emissao_Pedidos_Trimestral.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as err:
            st.error(f"Erro ao processar visão trimestral de emissão de pedidos: {err}")
