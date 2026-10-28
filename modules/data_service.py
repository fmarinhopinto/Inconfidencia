import os
import io
import re
import datetime
import pandas as pd
import numpy as np
import streamlit as st
import xlrd
from unidecode import unidecode

from modules.db_manager import (
    DB_PATH,
    table_exists,
    run_query,
    limpar_id,
    limpar_texto_corrompido,
    extrair_unidade_cirurgica,
    get_mapa_global_unidades,
    EXCECOES_MANUAIS_UNIDADES
)

DATA_CORTE = datetime.date(2026, 10, 21)


def find_column(df, target_names, partial=False):
    """Localiza uma coluna de forma insensível a maiúsculas/minúsculas e acentos."""
    if df is None or not hasattr(df, "columns"):
        return None
        
    cols = list(df.columns)
    
    def norm(s):
        return unidecode(str(s).lower().strip().replace("_", " ").replace(".", " "))

    norm_targets = [norm(t) for t in target_names]

    for col in cols:
        c_norm = norm(col)
        for target in norm_targets:
            if c_norm == target:
                return col

    if partial:
        for col in cols:
            c_norm = norm(col)
            for target in norm_targets:
                if target in c_norm:
                    return col

    return None


def sanitize_order_id(val):
    """Sanitiza chaves de pedido removendo pontuações decimais indesejadas."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.strip()


def sanitizar_chave_pedido(serie):
    """Aplica sanitização vetorizada em série de pedidos."""
    return serie.astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.lstrip('0').str.strip()


def cruzar_unidades_blindado(df_alvo, df_unidades):
    """Cruza o DataFrame alvo com o cadastro de unidades com blindagem de tipo e fallback."""
    if df_alvo is None or df_alvo.empty or df_unidades is None or df_unidades.empty:
        return df_alvo
        
    df_res = df_alvo.copy()
    
    col_ped_alvo = find_column(df_res, ["Pedido", "pedido", "n_pedido", "num_pedido", "nr_pedido"])
    if not col_ped_alvo:
        return df_res

    col_ped_u = find_column(df_unidades, ["Codigo pedido", "Pedido", "pedido", "n_pedido", "codigo_pedido"])
    col_nome_u = find_column(df_unidades, ["Unidade", "unidade", "nome_unidade", "nome"])
    
    if not col_ped_u or not col_nome_u:
        return df_res

    df_u_limpo = df_unidades.copy()
    df_u_limpo['chave_lookup'] = sanitizar_chave_pedido(df_u_limpo[col_ped_u])
    df_u_limpo = df_u_limpo[df_u_limpo['chave_lookup'] != ''].drop_duplicates(subset=['chave_lookup'])
    
    mapa = df_u_limpo.set_index('chave_lookup')[col_nome_u].to_dict()
    
    df_res['chave_lookup'] = sanitizar_chave_pedido(df_res[col_ped_alvo])
    unidades_encontradas = df_res['chave_lookup'].map(mapa)
    
    if 'Unidade' not in df_res.columns:
        df_res['Unidade'] = unidades_encontradas
    else:
        mask_vazia = df_res['Unidade'].isna() | df_res['Unidade'].astype(str).str.strip().isin(['', 'Não Informada', 'nan', 'None', '-'])
        df_res.loc[mask_vazia, 'Unidade'] = unidades_encontradas
        
    df_res.drop(columns=['chave_lookup'], inplace=True, errors='ignore')
    return df_res


def extrair_unidade_excecao(texto):
    """Extrai o nome da unidade a partir do campo textual Local de Entrega."""
    if pd.isna(texto) or not str(texto).strip():
        return None
    res = extrair_unidade_cirurgica(texto)
    if res and res.lower() not in ["nan", "none", "não informada", ""]:
        return res
    return None


def aplicar_regra_excecao(df_emergencial):
    """Aplica regra de extração da unidade e pedido a partir de Local de entrega."""
    if df_emergencial is None or df_emergencial.empty:
        return df_emergencial
    
    df = df_emergencial.copy()
    col_local = next((c for c in df.columns if any(p in str(c).lower() for p in ['entrega', 'local'])), None)
    
    if col_local:
        if 'Pedido' not in df.columns:
            df['Pedido'] = df[col_local].astype(str).str.extract(r'PED\.?:\s*(\d+)', expand=False)
        else:
            df['Pedido'] = df['Pedido'].fillna(df[col_local].astype(str).str.extract(r'PED\.?:\s*(\d+)', expand=False))
        
        df['Unidade_Extraida'] = df[col_local].apply(extrair_unidade_excecao)
        if 'Unidade' not in df.columns:
            df['Unidade'] = df['Unidade_Extraida']
        else:
            mask_unid = df['Unidade'].isna() | df['Unidade'].astype(str).str.strip().isin(['', 'nan', 'None', 'Não Informada'])
            df.loc[mask_unid & df['Unidade_Extraida'].notna(), 'Unidade'] = df['Unidade_Extraida']
            
    return df


def extrair_dados_entrega(texto):
    """Extrai tupla (pedido, unidade) de strings de entrega."""
    if pd.isna(texto) or not str(texto).strip():
        return None, None
    s = str(texto).strip()
    ped = None
    m_ped = re.search(r'PED\.?:\s*(\d+)', s, re.IGNORECASE)
    if m_ped:
        ped = m_ped.group(1).strip()
    unid = extrair_unidade_excecao(s)
    return ped, unid


def processar_planilha_emergencial(df_emergencial, df_unidades=None):
    """Processa a planilha complementar de pedidos e notas emergenciais."""
    if df_emergencial is None or df_emergencial.empty:
        return df_emergencial
        
    df = df_emergencial.copy()
    col_loc = next((c for c in df.columns if any(p in str(c).lower() for p in ['local', 'entrega'])), None)
    
    if col_loc:
        df['ped_extraido'] = df[col_loc].apply(lambda x: extrair_dados_entrega(x)[0])
        df['unid_extraida'] = df[col_loc].apply(lambda x: extrair_dados_entrega(x)[1])
        
        if 'Pedido' not in df.columns:
            df['Pedido'] = df['ped_extraido']
        else:
            df['Pedido'] = df['Pedido'].fillna(df['ped_extraido'])
            
        if 'Unidade' not in df.columns:
            df['Unidade'] = df['unid_extraida']
        else:
            mask = df['Unidade'].isna() | df['Unidade'].astype(str).str.strip().isin(['', 'nan', 'None', 'Não Informada'])
            df.loc[mask & df['unid_extraida'].notna(), 'Unidade'] = df['unid_extraida']
            
        df.drop(columns=['ped_extraido', 'unid_extraida'], inplace=True, errors='ignore')
        
    if df_unidades is not None and not df_unidades.empty:
        df = cruzar_unidades_blindado(df, df_unidades)
        
    return df


def load_uploaded_file(file_obj):
    """Lê arquivo enviado via upload com suporte a múltiplos formatos e engines."""
    if file_obj is None:
        return None
    try:
        if isinstance(file_obj, str):
            if not os.path.exists(file_obj):
                return None
            filename = file_obj
            with open(file_obj, "rb") as f:
                content = f.read()
        elif hasattr(file_obj, "name"):
            filename = file_obj.name
            content = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
        else:
            filename = "dados.xlsx"
            content = file_obj.read()

        buf = io.BytesIO(content)

        # 1. Tenta Openpyxl
        try:
            buf.seek(0)
            return pd.read_excel(buf, engine="openpyxl")
        except Exception:
            pass

        # 2. Tenta Xlrd (formato .xls antigo)
        try:
            buf.seek(0)
            return pd.read_excel(buf, engine="xlrd")
        except Exception:
            pass

        # 3. Tenta HTML / XML Spreadsheet
        try:
            buf.seek(0)
            tabelas = pd.read_html(buf, encoding="utf-8")
            if tabelas:
                return tabelas[0]
        except Exception:
            pass

        # 4. Tenta CSV
        try:
            buf.seek(0)
            return pd.read_csv(buf, sep=None, engine="python", encoding="utf-8")
        except Exception:
            pass

        return None
    except Exception as e:
        print(f"Erro ao carregar arquivo: {e}")
        return None


def base_valida(b):
    """Verifica se o caminho ou objeto de arquivo é válido para leitura."""
    if b is None:
        return False
    if isinstance(b, str):
        return os.path.exists(b)
    return hasattr(b, "read") or hasattr(b, "seek")


def get_file_mtime(f):
    """Obtém carimbo de tempo para invalidação de cache."""
    if isinstance(f, str) and os.path.exists(f):
        return os.path.getmtime(f)
    return getattr(f, "size", getattr(f, "name", 0))


@st.cache_data(show_spinner="Processando base com motor DuckDB de alta performance...", ttl=3600)
def process_data(file_acomp, file_bi_data, file_unid_data, _cache_version="v9_duckdb_elite", _mtime=None):
    """
    Pipeline analítico de alta performance para processamento dos dados do dashboard.
    Lê os dados diretamente do DuckDB persistido e executa deduplicação estrita.
    """
    # 1. Leitura de Alta Performance via DuckDB
    if not (isinstance(file_acomp, io.BytesIO) or hasattr(file_acomp, "read")) and table_exists("acompanhamento"):
        df_acomp_raw = run_query("SELECT * FROM acompanhamento")
    else:
        df_acomp_raw = load_uploaded_file(file_acomp)

    if not (isinstance(file_bi_data, io.BytesIO) or hasattr(file_bi_data, "read")) and table_exists("dados_bi"):
        df_bi_raw = run_query("SELECT * FROM dados_bi")
    else:
        df_bi_raw = load_uploaded_file(file_bi_data)

    # 2. Processamento do Acompanhamento
    col_cliente_acomp = find_column(df_acomp_raw, ["nome_cliente", "cliente", "nome cliente"])
    if col_cliente_acomp:
        df_acomp = df_acomp_raw[
            df_acomp_raw[col_cliente_acomp].astype(str).str.contains("INCONFIDENCIA", case=False, na=False)
        ].copy()
    else:
        df_acomp = df_acomp_raw.copy()

    # Nota Fiscal
    col_nota_acomp = find_column(df_acomp, ["Nota", "nota", "n_nota", "numero_nota"])
    if col_nota_acomp:
        nota_num = pd.to_numeric(df_acomp[col_nota_acomp], errors="coerce").fillna(0).astype(int)
        df_acomp["Nota"] = nota_num.astype(str)
    else:
        df_acomp["Nota"] = "0"

    # Pedido
    col_pedido_acomp = find_column(df_acomp, ["Pedido", "pedido", "n_pedido", "num_pedido", "nr_pedido", "numero_pedido", "ordem", "num_ordem"])
    if col_pedido_acomp and col_pedido_acomp != "Pedido":
        df_acomp["Pedido"] = df_acomp[col_pedido_acomp]
    elif "Pedido" not in df_acomp.columns:
        df_acomp["Pedido"] = [f"PED-{i+1}" for i in range(len(df_acomp))]

    df_acomp["Pedido"] = df_acomp["Pedido"].astype(str).str.strip()
    df_acomp["Pedido_Clean"] = df_acomp["Pedido"].apply(limpar_id)

    # Blindagem analítica determinística contra duplicidade de pedidos no portal
    col_nota_temp = find_column(df_acomp, ["Nota", "nota", "n_nota", "numero_nota"])
    if col_nota_temp:
        df_acomp = df_acomp.sort_values(by=[col_nota_temp], ascending=False).drop_duplicates(subset=["Pedido_Clean"], keep="first")
    else:
        df_acomp = df_acomp.drop_duplicates(subset=["Pedido_Clean"], keep="first")

    # Cidade (Fallback)
    col_cidade = find_column(df_acomp, ["cidade", "municipio", "cid"])
    if col_cidade and col_cidade != "cidade":
        df_acomp["cidade"] = df_acomp[col_cidade]
    elif "cidade" not in df_acomp.columns:
        df_acomp["cidade"] = "Não Informada"
    df_acomp["cidade"] = df_acomp["cidade"].fillna("Não Informada").astype(str).str.strip()
    df_acomp["cidade"] = df_acomp["cidade"].replace({"": "Não Informada", "nan": "Não Informada", "None": "Não Informada"})

    # Data e Status
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

    # 3. Processamento de Dados BI (Faturamento)
    col_cliente_bi = find_column(df_bi_raw, ["Cliente", "cliente", "nome_cliente"])
    if col_cliente_bi:
        df_bi = df_bi_raw[
            df_bi_raw[col_cliente_bi].astype(str).str.contains("INCONFIDENCIA", case=False, na=False)
        ].copy()
    else:
        df_bi = df_bi_raw.copy()

    col_nota_bi = find_column(df_bi, ["n_nota_obs", "nota", "numero_nota", "n_nota"])
    if col_nota_bi:
        nota_bi_num = pd.to_numeric(df_bi[col_nota_bi], errors="coerce")
        df_bi = df_bi[nota_bi_num.notna()].copy()
        df_bi["n_nota_obs"] = nota_bi_num.dropna().astype(int).astype(str)
    else:
        df_bi["n_nota_obs"] = ""

    col_vendas = find_column(df_bi, ["Vendas Total", "vendas total", "vlr_total", "valor_total", "total"])
    df_bi["Vendas Total"] = pd.to_numeric(df_bi[col_vendas], errors="coerce").fillna(0.0) if col_vendas else 0.0

    col_qtd = find_column(df_bi, ["QTD", "qtd", "quantidade", "quant"])
    df_bi["QTD"] = pd.to_numeric(df_bi[col_qtd], errors="coerce").fillna(0) if col_qtd else 0

    col_dt_hist = find_column(df_bi, ["dt_nota_hist", "data_nota", "dt_emissao", "data_faturamento"])
    df_bi["dt_nota_hist"] = df_bi[col_dt_hist] if col_dt_hist else "-"

    col_cod = find_column(df_bi, ["Cod", "cod", "codigo", "cod_produto", "cod_item"])
    df_bi["Cod"] = df_bi[col_cod] if col_cod else "-"

    col_descr = find_column(df_bi, ["Descr. produto", "descr. produto", "descricao", "descr_produto", "produto"])
    df_bi["Descr. produto"] = df_bi[col_descr] if col_descr else "-"

    col_preco = find_column(df_bi, ["Preço Venda", "preço venda", "preco venda", "preco_venda", "unitario"])
    if col_preco:
        df_bi["Preço Venda"] = pd.to_numeric(df_bi[col_preco], errors="coerce").fillna(0.0)
    else:
        df_bi["Preço Venda"] = np.where(df_bi["QTD"] > 0, df_bi["Vendas Total"] / df_bi["QTD"], 0.0)

    # Chaves limpas e deduplicação de itens no BI
    df_bi['NF_Clean'] = df_bi['n_nota_obs'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
    df_acomp['NF_Clean'] = df_acomp['Nota'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')

    cols_dup_bi = [c for c in ['NF_Clean', 'Cod', 'QTD', 'Vendas Total'] if c in df_bi.columns]
    if len(cols_dup_bi) >= 3:
        df_bi = df_bi.drop_duplicates(subset=cols_dup_bi, keep="last")

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

    # Cruzamento Acompanhamento + BI
    df_merged = pd.merge(df_acomp, resumo_bi, on="NF_Clean", how="left")
    is_faturado = (df_merged["NF_Clean"] != "") & (df_merged["bi_vendas_total"].notna())

    df_merged["Status_Faturamento"] = np.where(is_faturado, "Faturado", "Pendente / Não Faturado")
    df_merged["Vendas Total"] = np.where(is_faturado, df_merged["bi_vendas_total"], 0.0)
    df_merged["QTD Total"] = np.where(is_faturado, df_merged["bi_qtd_total"], 0)

    dt_faturado_str = pd.to_datetime(df_merged["bi_dt_nota_hist"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("-")
    df_merged["dt_nota_hist"] = np.where(is_faturado, dt_faturado_str, "-")

    # 4. Leitura da 3ª Planilha (Unidades)
    if not (isinstance(file_unid_data, io.BytesIO) or hasattr(file_unid_data, "read")) and table_exists("unidades"):
        df_unidades = run_query("SELECT * FROM unidades")
    else:
        df_unidades = load_uploaded_file(file_unid_data)

    if df_unidades is not None and not df_unidades.empty:
        col_ped = find_column(df_unidades, ["Codigo pedido", "pedido", "n_pedido", "codigo_pedido"])
        col_uni = find_column(df_unidades, ["Unidade", "unidade", "nome_unidade", "nome"])
        if not col_ped:
            col_ped = df_unidades.columns[1] if len(df_unidades.columns) > 1 else df_unidades.columns[0]
        if not col_uni:
            col_uni = df_unidades.columns[3] if len(df_unidades.columns) > 3 else df_unidades.columns[-1]

        df_unidades_unique = df_unidades[[col_ped, col_uni]].copy()
        df_unidades_unique.rename(columns={col_ped: "Pedido", col_uni: "Unidade_Extraida"}, inplace=True)
        df_unidades_unique["Pedido"] = df_unidades_unique["Pedido"].astype(str).str.strip()
        df_unidades_unique["Pedido_Clean"] = df_unidades_unique["Pedido"].apply(limpar_id)
        df_unidades_unique["Unidade_Extraida"] = df_unidades_unique["Unidade_Extraida"].apply(limpar_texto_corrompido)
        df_unidades_unique = df_unidades_unique.drop_duplicates(subset=["Pedido_Clean"], keep="first")

        df_merged = pd.merge(df_merged, df_unidades_unique[["Pedido_Clean", "Unidade_Extraida"]], on="Pedido_Clean", how="left")
    else:
        df_unidades_unique = pd.DataFrame(columns=["Pedido", "Pedido_Clean", "Unidade_Extraida"])
        df_merged["Unidade_Extraida"] = None
        col_ped = "Não Informado"
        col_uni = "Não Informado"

    df_merged = cruzar_unidades_blindado(df_merged, df_unidades_unique)

    # Unidade Final consolidada
    df_merged['Unidade_Final'] = df_merged.get('Unidade_Extraida', pd.Series(dtype='object')).fillna(df_merged.get('Unidade', pd.Series(dtype='object')))
    df_merged['Unidade_Final'] = df_merged['Unidade_Final'].fillna(df_merged['cidade'])
    df_merged['Unidade_Final'] = df_merged['Unidade_Final'].fillna('Não Informada').replace({'': 'Não Informada', 'nan': 'Não Informada', 'None': 'Não Informada'})

    df_merged["UNIDADE"] = df_merged["Unidade_Final"]
    df_merged["Unidade"] = df_merged["Unidade_Final"]
    df_merged["Unidade_Real"] = df_merged["Unidade_Final"]

    # Enriquecimento de df_bi com Pedido e Unidade para a Aba 2
    df_bi_merged = pd.merge(df_bi, df_merged[['NF_Clean', 'Pedido', 'Pedido_Clean', 'Unidade_Final']].drop_duplicates(subset=['NF_Clean']), on='NF_Clean', how='left')
    df_bi_merged['UNIDADE'] = df_bi_merged['Unidade_Final'].fillna('Não Informada')

    return df_merged, df_bi_merged, df_unidades_unique, col_ped, 'Unidade_Extraida'


def obter_base_consolidada(file_acomp, file_bi, file_unid, uploaded_nf_emergencial=None):
    """
    Consolida todas as fontes analíticas (Acompanhamento + BI + Unidades + Emergencial)
    com garantia determinística anti-duplicação, mapeamento cirúrgico de unidades
    e aplicação das exceções prioritárias manuais.
    """
    db_mtime = os.path.getmtime(DB_PATH) if os.path.exists(DB_PATH) else 0.0
    mtimes = (
        get_file_mtime(file_acomp),
        get_file_mtime(file_bi),
        get_file_mtime(file_unid),
        db_mtime
    )
    df_merged, df_bi_raw, df_unidades_unique, col_ped, col_uni = process_data(
        file_acomp, file_bi, file_unid, _mtime=mtimes
    )

    # Identificação da planilha emergencial / monitoramento
    candidatos_emerg = [
        'monitoramento.xlsx',
        'monitoramento.xls',
        'monitoramento.csv',
        'nf_emergencial.xlsx',
        'nf_emergencial.csv',
        'nf_emergencial.xls',
        os.path.join('INC', 'monitoramento', 'monitoramento.xlsx')
    ]
    file_nf_emergencial_fixo = None
    for cand in candidatos_emerg:
        if os.path.exists(cand):
            file_nf_emergencial_fixo = cand
            break

    arquivo_nf_emerg_alvo = uploaded_nf_emergencial if uploaded_nf_emergencial is not None else file_nf_emergencial_fixo

    nfs_extras = set()
    pedidos_extras = set()
    df_nf_emergencial = None
    col_nf = None

    if arquivo_nf_emerg_alvo is not None:
        try:
            df_emerg = load_uploaded_file(arquivo_nf_emerg_alvo)
            if df_emerg is not None and not df_emerg.empty:
                if isinstance(df_emerg.columns, pd.MultiIndex):
                    df_emerg.columns = df_emerg.columns.get_level_values(-1)

                df_emerg = aplicar_regra_excecao(df_emerg)
                if 'Pedido' in df_emerg.columns:
                    df_emerg['Pedido'] = df_emerg['Pedido'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_emerg = processar_planilha_emergencial(df_emerg, df_unidades=df_unidades_unique)
                df_nf_emergencial = df_emerg

                cols_nf_encontradas = [c for c in df_nf_emergencial.columns if str(c).strip().upper() == 'NF']
                if not cols_nf_encontradas:
                    cols_nf_encontradas = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['nota', 'nf', 'n_nota', 'numero_nf', 'numero_nota', 'doc', 'fatura'])]
                col_nf = cols_nf_encontradas[0] if cols_nf_encontradas else df_nf_emergencial.columns[0]

                nfs_extras_clean = (
                    df_nf_emergencial[col_nf]
                    .dropna().astype(str)
                    .str.replace('NF-', '', case=False)
                    .str.replace(r'\D', '', regex=True)
                    .str.lstrip('0').str.strip()
                )
                nfs_extras_raw = (
                    df_nf_emergencial[col_nf]
                    .dropna().astype(str)
                    .str.replace('NF-', '', case=False)
                    .str.strip()
                )
                nfs_extras = set(nfs_extras_clean[nfs_extras_clean != '']).union(set(nfs_extras_raw[nfs_extras_raw != '']))

                cols_ped = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['pedido', 'ped', 'ordem'])]
                if cols_ped:
                    vals_ped_raw = df_nf_emergencial[cols_ped[0]].dropna().astype(str).str.strip()
                    vals_ped_limpos = vals_ped_raw.str.replace(r'\D', '', regex=True).str.lstrip('0')
                    pedidos_extras = set(vals_ped_raw[vals_ped_raw != '']).union(set(vals_ped_limpos[vals_ped_limpos != '']))
        except Exception as e:
            print(f"Aviso leitura emergencial: {e}")

    # Unificação de Faturamento Oficial + Emergencial
    df_fat_oficial = df_bi_raw.copy() if 'df_bi_raw' in locals() and df_bi_raw is not None else None
    df_fat_emerg = df_nf_emergencial.copy() if df_nf_emergencial is not None else None
    df_faturamento_total = pd.DataFrame()

    if df_fat_oficial is not None and not df_fat_oficial.empty:
        df_faturamento_total = df_fat_oficial.copy()

    if df_fat_emerg is not None and not df_fat_emerg.empty:
        df_emerg = aplicar_regra_excecao(df_fat_emerg)
        if df_unidades_unique is not None and not df_unidades_unique.empty and 'Pedido' in df_emerg.columns:
            df_u = df_unidades_unique.copy()
            col_ped_u = 'Pedido'
            col_nome_u = 'Unidade_Extraida'
            df_u['chave_limpa'] = df_u[col_ped_u].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            mapa_unidades = df_u[df_u['chave_limpa'] != ''].drop_duplicates(subset=['chave_limpa']).set_index('chave_limpa')[col_nome_u].to_dict()
            chave_emerg = df_emerg['Pedido'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            unidade_cruzada = chave_emerg.map(mapa_unidades)
            df_emerg['Unidade'] = df_emerg['Unidade'].fillna(unidade_cruzada).fillna(df_emerg.get('cidade', 'INCONFIDENCIA LOCADORA'))
        df_nf_emergencial = df_emerg.copy()
        df_faturamento_total = pd.concat([df_faturamento_total, df_emerg], ignore_index=True)

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

    if nfs_extras:
        nfs_validas.update(nfs_extras)

    # Cruzamento emergencial por chave de pedido
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

                df_merged = df_merged.merge(df_emerg_sub_slice, on='chave_ped_join', how='left')

                if 'NF_Emergencial' in df_merged.columns:
                    mask_tem_emerg = df_merged['NF_Emergencial'].notna() & (df_merged['NF_Emergencial'].astype(str).str.strip() != '')
                    mask_vazia = df_merged['Nota'].isna() | df_merged['Nota'].astype(str).str.strip().isin(['', '0', 'nan', 'None', '-'])
                    alvo_nota = mask_vazia & mask_tem_emerg
                    df_merged.loc[alvo_nota, 'Nota'] = df_merged.loc[alvo_nota, 'NF_Emergencial'].astype(str)
                    df_merged['NF_Clean'] = df_merged['Nota'].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')
                    df_merged.drop(columns=['NF_Emergencial'], inplace=True)

                if 'Unidade_Emergencial' in df_merged.columns:
                    mask_unid_emerg = df_merged['Unidade_Emergencial'].notna() & (df_merged['Unidade_Emergencial'].astype(str).str.strip() != '')
                    mask_unid_vazia = df_merged['Unidade_Final'].isna() | df_merged['Unidade_Final'].astype(str).str.strip().isin(['', 'Não Informada', 'nan', 'None', '-'])
                    alvo_unid = mask_unid_vazia & mask_unid_emerg
                    df_merged.loc[alvo_unid, 'Unidade_Final'] = df_merged.loc[alvo_unid, 'Unidade_Emergencial'].astype(str)
                    df_merged['UNIDADE'] = df_merged['Unidade_Final']
                    df_merged['Unidade'] = df_merged['Unidade_Final']
                    df_merged['Unidade_Real'] = df_merged['Unidade_Final']
                    df_merged.drop(columns=['Unidade_Emergencial'], inplace=True)

                df_merged.drop(columns=['chave_ped_join'], inplace=True, errors='ignore')
        except Exception as e_merge:
            print(f"Aviso no cruzamento emergencial: {e_merge}")

    # Cruzamento padronizado de status faturado
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

    # Liberação da data de corte para registros do período ou faturados no conjunto combinado
    serie_dt_corte = pd.to_datetime(df_merged['Dt_inser'], dayfirst=True, errors='coerce').dt.normalize()
    ts_corte = pd.Timestamp(DATA_CORTE).normalize()
    cond_corte = (serie_dt_corte >= ts_corte) | mask_faturados_combinado

    if pedidos_extras:
        mask_ped = df_merged['Pedido_Clean'].isin(pedidos_extras) | df_merged['Pedido'].astype(str).str.strip().isin(pedidos_extras)
        cond_corte = cond_corte | mask_ped

    df_merged = df_merged[cond_corte].copy()

    # Adiciona linhas complementares da emergencial
    if df_nf_emergencial is not None and not df_nf_emergencial.empty:
        try:
            df_novas = pd.DataFrame()
            cols_ped = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['pedido', 'ped', 'ordem'])]
            if cols_ped:
                df_novas['Pedido'] = df_nf_emergencial[cols_ped[0]].astype(str).str.strip()
            elif col_nf and col_nf in df_nf_emergencial.columns:
                df_novas['Pedido'] = "NF-" + df_nf_emergencial[col_nf].astype(str).str.strip()
            else:
                df_novas['Pedido'] = [f"EMERG-{i+1}" for i in range(len(df_nf_emergencial))]

            df_novas['Pedido_Clean'] = df_novas['Pedido'].str.replace(r'\D', '', regex=True).str.lstrip('0')
            df_novas['Nota'] = df_nf_emergencial[col_nf].astype(str).str.strip() if col_nf else df_novas['Pedido']
            df_novas['NF_Clean'] = df_novas['Nota'].str.replace(r'\D', '', regex=True).str.lstrip('0')

            col_loc_cand = next((c for c in df_nf_emergencial.columns if 'entrega' in str(c).lower() or 'local' in str(c).lower()), None)
            col_unid_cand = next((c for c in df_nf_emergencial.columns if str(c).strip().lower() in ['unidade', 'unidade_excecao', 'unidade_extraida', 'unidade_final']), None)

            if col_loc_cand:
                df_novas['UNIDADE'] = df_nf_emergencial[col_loc_cand].apply(extrair_unidade_excecao)
            elif col_unid_cand:
                df_novas['UNIDADE'] = df_nf_emergencial[col_unid_cand].astype(str).str.strip()
            else:
                df_novas['UNIDADE'] = ""

            mapa_global = get_mapa_global_unidades()
            mask_unid_gen = (
                df_novas['UNIDADE'].isna() |
                df_novas['UNIDADE'].astype(str).str.strip().isin(['', 'nan', 'None', 'Não Informada']) |
                df_novas['UNIDADE'].astype(str).str.contains(r'INCONFID[EÊ]NCIA\s+LOCADORA', case=False, na=False)
            )
            unid_map = df_novas['Pedido_Clean'].map(mapa_global)
            df_novas.loc[mask_unid_gen & unid_map.notna() & (unid_map != ''), 'UNIDADE'] = unid_map

            df_novas['Unidade_Final'] = df_novas['UNIDADE']
            df_novas['Unidade'] = df_novas['UNIDADE']
            df_novas['Unidade_Real'] = df_novas['UNIDADE']
            df_novas['cidade'] = df_novas['UNIDADE']
            df_novas['Unidade_Clean'] = df_novas['UNIDADE'].apply(limpar_texto_corrompido)

            cols_dt = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['data', 'dt', 'emissao'])]
            df_novas['Dt_inser'] = df_nf_emergencial[cols_dt[0]].astype(str).str.strip() if cols_dt else DATA_CORTE.strftime("%d/%m/%Y")
            df_novas['dt_nota_hist'] = df_novas['Dt_inser']
            df_novas['Status'] = 'Faturado'
            df_novas['Status_Faturamento'] = 'Faturado'
            df_novas['Tipo'] = 'Emergencial'

            cols_vlr = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['valor', 'total', 'vendas', 'preco'])]
            df_novas['Vendas Total'] = pd.to_numeric(df_nf_emergencial[cols_vlr[0]], errors='coerce').fillna(0.0) if cols_vlr else 0.0

            cols_qtd = [c for c in df_nf_emergencial.columns if any(p in str(c).lower() for p in ['qtd', 'quantidade', 'quant'])]
            df_novas['QTD Total'] = pd.to_numeric(df_nf_emergencial[cols_qtd[0]], errors='coerce').fillna(0) if cols_qtd else 1

            df_novas['Data de Inserção'] = df_novas['Dt_inser']
            df_novas['Valor (R$)'] = df_novas['Vendas Total']

            df_merged = pd.concat([df_merged, df_novas], ignore_index=True)

            if 'Pedido_Clean' in df_merged.columns:
                df_merged = df_merged.drop_duplicates(subset=['Pedido_Clean'], keep='last')
            elif 'Pedido' in df_merged.columns:
                df_merged = df_merged.drop_duplicates(subset=['Pedido'], keep='last')
        except Exception as err:
            print(f"Aviso concat emergencial: {err}")

    # Blindagem de Unidades
    mapa_global = get_mapa_global_unidades()
    mask_precisa_unidade = (
        df_merged['Unidade_Final'].isna() |
        df_merged['Unidade_Final'].astype(str).str.strip().isin(['', 'nan', 'None', 'Não Informada', '-']) |
        df_merged['Unidade_Final'].astype(str).str.contains(r'INCONFID[EÊ]NCIA\s+LOCADORA', case=False, na=False)
    )
    if mask_precisa_unidade.any():
        unid_resolvida = df_merged['Pedido_Clean'].map(mapa_global)
        df_merged.loc[mask_precisa_unidade & unid_resolvida.notna() & (unid_resolvida != ''), 'Unidade_Final'] = unid_resolvida
        col_local_merged = next((c for c in df_merged.columns if 'entrega' in str(c).lower() or 'local' in str(c).lower()), None)
        if col_local_merged:
            unid_rec = df_merged[col_local_merged].apply(extrair_unidade_excecao)
            df_merged.loc[mask_precisa_unidade & unid_rec.notna() & (unid_rec != ''), 'Unidade_Final'] = unid_rec
        mask_ainda_gen = df_merged['Unidade_Final'].astype(str).str.contains(r'INCONFID[EÊ]NCIA\s+LOCADORA', case=False, na=False)
        if 'cidade' in df_merged.columns:
            df_merged.loc[mask_ainda_gen & df_merged['cidade'].notna() & (df_merged['cidade'] != 'Não Informada'), 'Unidade_Final'] = df_merged['cidade']

    # Exceções Manuais Prioritárias
    for ped_exc, unid_exc in EXCECOES_MANUAIS_UNIDADES.items():
        mask_exc = (df_merged['Pedido_Clean'] == ped_exc) | (df_merged['Pedido'].astype(str).str.strip() == ped_exc)
        if mask_exc.any():
            df_merged.loc[mask_exc, 'Unidade_Final'] = unid_exc

    df_merged['UNIDADE'] = df_merged['Unidade_Final']
    df_merged['Unidade'] = df_merged['Unidade_Final']
    df_merged['Unidade_Real'] = df_merged['Unidade_Final']
    df_merged['Unidade_Clean'] = df_merged['Unidade_Final'].apply(limpar_texto_corrompido)

    return df_merged, df_bi_raw, df_unidades_unique, col_ped, col_uni, nfs_extras, pedidos_extras


def processar_aba_trimestral(df_cad_raw, df_ped_raw):
    """
    Processa a consolidação trimestral automática (últimos 3 meses até a data presente)
    cruzando o cadastro mestre de unidades com os relatórios de pedidos.
    """
    if df_cad_raw is None or df_cad_raw.empty or df_ped_raw is None or df_ped_raw.empty:
        return None, None, None, None, None, None

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

    df_unidades['Unidade_Clean'] = df_unidades['Unidade'].apply(limpar_texto_corrompido)

    if isinstance(df_ped_raw.columns, pd.MultiIndex):
        df_ped_raw.columns = df_ped_raw.columns.get_level_values(-1)

    col_data_ped = [c for c in df_ped_raw.columns if 'data' in str(c).lower()][0]
    col_uni_ped = [c for c in df_ped_raw.columns if 'unidade' in str(c).lower()][0]
    col_vlr_ped = [c for c in df_ped_raw.columns if any(p in str(c).lower() for p in ['valor total', 'vendas total', 'total'])][0]

    df_pedidos = df_ped_raw[[col_data_ped, col_uni_ped, col_vlr_ped]].copy()
    df_pedidos.rename(columns={col_uni_ped: 'Unidade'}, inplace=True)
    df_pedidos['Unidade_Clean'] = df_pedidos['Unidade'].apply(limpar_texto_corrompido)
    df_pedidos['Valor'] = pd.to_numeric(df_pedidos[col_vlr_ped], errors='coerce').fillna(0.0)

    dt_parsed = pd.to_datetime(df_pedidos[col_data_ped], errors='coerce', dayfirst=True).dt.normalize()

    hoje = datetime.date.today()
    ano_inicio = hoje.year
    mes_inicio = hoje.month - 2
    while mes_inicio <= 0:
        mes_inicio += 12
        ano_inicio -= 1
    data_inicio_trimestre = datetime.date(ano_inicio, mes_inicio, 1)
    data_fim_trimestre = hoje
    ts_ini_tri = pd.Timestamp(data_inicio_trimestre).normalize()
    ts_fim_tri = pd.Timestamp(data_fim_trimestre).normalize()

    df_pedidos = df_pedidos[(dt_parsed >= ts_ini_tri) & (dt_parsed <= ts_fim_tri)]

    meses_pt = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    dt_filtrado = pd.to_datetime(df_pedidos[col_data_ped], errors='coerce', dayfirst=True)
    df_pedidos['AnoMes'] = dt_filtrado.dt.to_period('M')
    df_pedidos['Mes'] = dt_filtrado.dt.month.map(meses_pt)

    meses_ordenados = df_pedidos.dropna(subset=['AnoMes', 'Mes']).sort_values('AnoMes')['Mes'].unique().tolist()

    if not meses_ordenados:
        meses_ordenados = ["Período Selecionado"]
        pivot = pd.DataFrame(columns=['Unidade_Clean', 'Período Selecionado'])
    else:
        pivot = df_pedidos.groupby(['Unidade_Clean', 'Mes'])['Valor'].sum().unstack(fill_value=0.0)
        for m in meses_ordenados:
            if m not in pivot.columns:
                pivot[m] = 0.0
        pivot = pivot[meses_ordenados].reset_index()

    matriz = pd.merge(df_unidades, pivot, on='Unidade_Clean', how='left')
    for m in meses_ordenados:
        matriz[m] = matriz[m].fillna(0.0)

    matriz_exibicao = matriz[['Unidade', 'ATIVO'] + meses_ordenados].copy()
    for m in meses_ordenados:
        matriz_exibicao[m] = matriz_exibicao[m].apply(
            lambda v: f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if isinstance(v, (int, float)) and v > 0 else "Sem Compra"
        )

    total_cad = len(matriz_exibicao)
    total_ativas = len(matriz_exibicao[matriz_exibicao['ATIVO'] == 'SIM'])
    unidades_com_compra = int((matriz[meses_ordenados].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) > 0).sum())
    unidades_sem_compra = total_cad - unidades_com_compra

    return matriz_exibicao, meses_ordenados, total_cad, total_ativas, unidades_com_compra, unidades_sem_compra

