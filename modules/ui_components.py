import os
import streamlit as st
import pandas as pd
from PIL import Image

def inject_custom_css():
    """
    Injeta a identidade visual corporativa do Grupo Fortpel & Inconfidência.
    Garante layout 100% responsivo, largura total sem margens restritivas,
    e estética moderna de nível executivo.
    """
    st.markdown("""
    <style>
        /* Container principal em largura total */
        .block-container {
            max-width: 100% !important;
            width: 100% !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            padding-top: 1rem !important;
            padding-bottom: 2rem !important;
        }

        .main .block-container {
            max-width: 100% !important;
        }

        [data-testid="stHorizontalBlock"] {
            width: 100% !important;
            gap: 1.5rem !important;
        }
        
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

        /* Barra Lateral (Sidebar) em Azul Petróleo */
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

        /* Botão recolher/abrir da Sidebar sempre visível */
        [data-testid="collapsedControl"] {
            display: block !important;
            opacity: 1 !important;
            visibility: visible !important;
            color: #FFFFFF !important;
            background-color: #3B6077 !important;
            border-radius: 8px !important;
            padding: 8px !important;
            transform: scale(1.2);
            margin-top: 10px !important;
            margin-left: 10px !important;
            box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
            z-index: 999999 !important;
        }
        
        [data-testid="collapsedControl"] svg {
            fill: #FFFFFF !important;
            stroke: #FFFFFF !important;
            color: #FFFFFF !important;
            width: 22px !important;
            height: 22px !important;
        }

        /* Cards de KPIs com sombra e borda arredondada */
        div[data-testid="stMetric"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 16px 20px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
            transition: transform 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
        }
        div[data-testid="stMetric"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 12px -2px rgba(0, 0, 0, 0.08);
        }

        div[data-testid="stMetric"] label {
            color: #64748B !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            color: #1E293B !important;
            font-size: 1.8rem !important;
            font-weight: 700 !important;
        }

        /* Card 4 (Valor Total Faturado): Destaque Corporativo Inconfidência */
        div[data-testid="stMetric"]:nth-of-type(4) {
            border-left: 5px solid #3B6077 !important;
        }
        div[data-testid="stMetric"]:nth-of-type(4) div[data-testid="stMetricValue"] {
            color: #3B6077 !important;
        }

        /* Abas com design executivo */
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            border-bottom: 2px solid #E2E8F0;
            padding-bottom: 4px;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: transparent;
            border: none;
            color: #64748B;
            font-weight: 600;
            padding: 10px 18px;
            border-radius: 8px;
            transition: all 0.2s ease;
        }
        .stTabs [data-baseweb="tab"]:hover {
            background-color: #EDEAE1;
            color: #1E293B;
        }
        .stTabs [aria-selected="true"] {
            background-color: #3B6077 !important;
            color: #FFFFFF !important;
            box-shadow: 0 2px 4px rgba(59, 96, 119, 0.25);
        }

        /* Botões corporativos */
        button[kind="primary"] {
            background-color: #E6192B !important;
            border-color: #E6192B !important;
            color: #FFFFFF !important;
            font-weight: 600;
            border-radius: 8px;
        }
        button[kind="primary"]:hover {
            background-color: #B51220 !important;
            border-color: #B51220 !important;
        }

        /* Card de detalhamento de pedido */
        .order-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 20px;
            margin-top: 15px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.04);
        }
    </style>
    """, unsafe_allow_html=True)


def format_currency_brl(val):
    """Formata valor float em moeda BRL (R$ 1.234,56)."""
    if pd.isna(val) or val is None:
        return "R$ 0,00"
    try:
        val_float = float(val)
        return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return f"R$ {val}"


def format_date_br(val):
    """Formata datas em padrão brasileiro (DD/MM/AAAA)."""
    if pd.isna(val) or val is None or str(val).strip() in ["", "-", "nan", "None"]:
        return "-"
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
        return str(val).split("T")[0]
    except Exception:
        return str(val)


def render_header():
    """Renderiza o cabeçalho executivo do dashboard."""
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


def render_kpi_cards(total_pedidos, faturados, pendentes, valor_total):
    """Renderiza os 4 cards executivos de métricas e KPIs."""
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(label="Total de Pedidos", value=f"{total_pedidos:,}".replace(",", "."))

    with col2:
        delta_fat = f"{(faturados / total_pedidos * 100):.1f}% do total" if total_pedidos > 0 else None
        st.metric(label="Pedidos Faturados", value=f"{faturados:,}".replace(",", "."), delta=delta_fat)

    with col3:
        delta_pend = f"-{(pendentes / total_pedidos * 100):.1f}%" if total_pedidos > 0 else None
        st.metric(
            label="Pedidos Pendentes / Sem Nota",
            value=f"{pendentes:,}".replace(",", "."),
            delta=delta_pend,
            delta_color="inverse"
        )

    with col4:
        st.metric(label="Valor Total Faturado", value=format_currency_brl(valor_total))

    st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)


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


def exibir_tela_login(credenciais=None):
    """
    Renderiza a interface de login segura e corporativa.
    Oculta a barra lateral até que a autenticação seja concluída.
    """
    if credenciais is None:
        credenciais = CREDENCIAIS
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    _, col_login, _ = st.columns([1, 1.2, 1])
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
                if usuario_input in credenciais and credenciais[usuario_input]["senha"] == senha_input:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_conectado"] = usuario_input
                    st.session_state["perfil"] = credenciais[usuario_input].get("perfil", "cliente")
                    st.rerun()
                else:
                    st.error("❌ Usuário ou senha incorretos. Verifique suas credenciais.")
