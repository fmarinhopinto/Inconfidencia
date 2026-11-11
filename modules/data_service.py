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

DATA_CORTE = datetime.date(2026, 11, 1)


def find_column(df, target_names, partial=False):
    """Localiza uma coluna de forma insensível a maiúsculas/minúsculas, acentos e tuplas (MultiIndex)."""
    if df is None or not hasattr(df, "columns"):
        return None
        
    cols = list(df.columns)
    
    def norm(s):
        if isinstance(s, tuple):
            s = " ".join(str(p) for p in s if str(p) and not str(p).startswith("Unnamed"))
        return unidecode(str(s).lower().strip().replace("_", " ").replace(".", " "))

    norm_targets = [norm(t) for t in target_names]

    for col in cols:
        c_norm = norm(col)
        for target in norm_targets:
            if c_norm == target or (isinstance(col, tuple) and any(norm(elem) == target for elem in col)):
                return col

    if partial:
        for col in cols:
            c_norm = norm(col)
            for target in norm_targets:
                if target in c_norm or (isinstance(col, tuple) and any(target in norm(elem) for elem in col)):
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





def load_uploaded_file(file_obj):
    """Lê arquivo enviado via upload com suporte resiliente a múltiplos formatos, encodings e engines."""
    if file_obj is None:
        return None
    try:
        filename = ""
        if isinstance(file_obj, str):
            if not os.path.exists(file_obj):
                return None
            filename = os.path.basename(file_obj).lower()
            with open(file_obj, "rb") as f:
                content = f.read()
        elif hasattr(file_obj, "name"):
            filename = str(file_obj.name).strip().lower()
            content = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
        else:
            filename = "dados.xlsx"
            content = file_obj.read()

        buf = io.BytesIO(content)
        df_result = None

        # 1. Tenta Openpyxl (.xlsx)
        if filename.endswith(".xlsx") or not filename.endswith(".csv"):
            try:
                buf.seek(0)
                df_result = pd.read_excel(buf, engine="openpyxl")
            except Exception:
                pass

        # 2. Tenta Xlrd (.xls antigo)
        if df_result is None and (filename.endswith(".xls") or not filename.endswith(".csv")):
            try:
                buf.seek(0)
                df_result = pd.read_excel(buf, engine="xlrd")
            except Exception:
                pass

        # 3. Fallback para HTML / XML Spreadsheet (muito comum em ERPs como unidades.xls e listagem_unidades.xls)
        if df_result is None:
            for enc in ["latin1", "utf-8", "cp1252", "iso-8859-1"]:
                try:
                    buf.seek(0)
                    tables = pd.read_html(buf, encoding=enc, decimal=",", thousands=".")
                    if tables:
                        df_result = max(tables, key=lambda t: t.shape[0] * t.shape[1])
                        break
                except Exception:
                    pass

        # 4. Fallback para CSV delimitado
        if df_result is None:
            for enc in ["utf-8", "latin1", "cp1252", "iso-8859-1"]:
                try:
                    buf.seek(0)
                    df_result = pd.read_csv(buf, sep=None, engine="python", encoding=enc)
                    break
                except Exception:
                    pass

        if df_result is None:
            return None

        # Padronização e Blindagem dos Cabeçalhos
        if isinstance(df_result.columns, pd.MultiIndex):
            # Se for MultiIndex, pega o último nível
            df_result.columns = df_result.columns.get_level_values(-1)

        # Higieniza espaços e desduplica colunas de mesmo nome
        cols = []
        counts = {}
        for col in df_result.columns:
            c_str = str(col).strip()
            if c_str in counts:
                counts[c_str] += 1
                cols.append(f"{c_str}_{counts[c_str]}")
            else:
                counts[c_str] = 0
                cols.append(c_str)
        df_result.columns = cols

        return df_result
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
    # 1. Leitura de Alta Performance via DuckDB ou Objetos em Memória
    if isinstance(file_acomp, pd.DataFrame):
        df_acomp_raw = file_acomp
    elif not (isinstance(file_acomp, io.BytesIO) or hasattr(file_acomp, "read")) and table_exists("acompanhamento"):
        df_acomp_raw = run_query("SELECT * FROM acompanhamento")
    else:
        df_acomp_raw = load_uploaded_file(file_acomp)

    if isinstance(file_bi_data, pd.DataFrame):
        df_bi_raw = file_bi_data
    elif not (isinstance(file_bi_data, io.BytesIO) or hasattr(file_bi_data, "read")) and table_exists("dados_bi"):
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

    # Garantir que não existam colunas com nomes duplicados em df_bi
    df_bi = df_bi.loc[:, ~df_bi.columns.duplicated()].copy()

    cols_dup_bi = [c for c in ['NF_Clean', 'Cod', 'QTD', 'Vendas Total'] if c in df_bi.columns]
    if len(cols_dup_bi) >= 3:
        df_bi = df_bi.drop_duplicates(subset=cols_dup_bi, keep="last")

    # Resumo agrupado do BI por NF_Clean com proteção total contra falhas de índice
    df_bi_valid = df_bi[df_bi['NF_Clean'] != ''].copy()
    if not df_bi_valid.empty:
        agg_dict = {}
        if "Vendas Total" in df_bi_valid.columns:
            agg_dict["Vendas Total"] = "sum"
        if "QTD" in df_bi_valid.columns:
            agg_dict["QTD"] = "sum"
        if "dt_nota_hist" in df_bi_valid.columns:
            agg_dict["dt_nota_hist"] = "first"

        resumo_bi = df_bi_valid.groupby("NF_Clean", as_index=False).agg(agg_dict).rename(columns={
            "Vendas Total": "bi_vendas_total",
            "QTD": "bi_qtd_total",
            "dt_nota_hist": "bi_dt_nota_hist"
        })
        for c in ["bi_vendas_total", "bi_qtd_total", "bi_dt_nota_hist"]:
            if c not in resumo_bi.columns:
                resumo_bi[c] = None
    else:
        resumo_bi = pd.DataFrame(columns=["NF_Clean", "bi_vendas_total", "bi_qtd_total", "bi_dt_nota_hist"])

    # Cruzamento Acompanhamento + BI
    df_merged = pd.merge(df_acomp, resumo_bi, on="NF_Clean", how="left")
    is_faturado = (df_merged["NF_Clean"] != "") & (df_merged["bi_vendas_total"].notna())

    df_merged["Status_Faturamento"] = np.where(is_faturado, "Faturado", "Pendente / Não Faturado")
    df_merged["Vendas Total"] = np.where(is_faturado, df_merged["bi_vendas_total"], 0.0)
    df_merged["QTD Total"] = np.where(is_faturado, df_merged["bi_qtd_total"], 0)

    dt_faturado_str = pd.to_datetime(df_merged["bi_dt_nota_hist"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("-")
    df_merged["dt_nota_hist"] = np.where(is_faturado, dt_faturado_str, "-")

    # 4. Leitura da 3ª Planilha (Unidades)
    if isinstance(file_unid_data, pd.DataFrame):
        df_unidades = file_unid_data
    elif not (isinstance(file_unid_data, io.BytesIO) or hasattr(file_unid_data, "read")) and table_exists("unidades"):
        df_unidades = run_query("SELECT * FROM unidades")
    else:
        df_unidades = load_uploaded_file(file_unid_data)

    if df_unidades is not None and not df_unidades.empty:
        if isinstance(df_unidades.columns, pd.MultiIndex):
            df_unidades.columns = df_unidades.columns.get_level_values(-1)
        df_unidades.columns = [str(c).strip() for c in df_unidades.columns]

        col_ped = find_column(df_unidades, ["Codigo pedido", "pedido", "n_pedido", "codigo_pedido"])
        col_uni = find_column(df_unidades, ["Unidade", "unidade", "nome_unidade", "nome"])
        if not col_ped:
            col_ped = df_unidades.columns[1] if len(df_unidades.columns) > 1 else df_unidades.columns[0]
        if not col_uni:
            col_uni = df_unidades.columns[3] if len(df_unidades.columns) > 3 else df_unidades.columns[-1]

        ped_series = df_unidades[col_ped]
        if isinstance(ped_series, pd.DataFrame):
            ped_series = ped_series.iloc[:, 0]

        uni_series = df_unidades[col_uni]
        if isinstance(uni_series, pd.DataFrame):
            uni_series = uni_series.iloc[:, 0]

        df_unidades_unique = pd.DataFrame({
            "Pedido": ped_series.astype(str).str.strip(),
            "Unidade_Extraida": uni_series.apply(limpar_texto_corrompido)
        })
        df_unidades_unique["Pedido_Clean"] = df_unidades_unique["Pedido"].apply(limpar_id)
        df_unidades_unique = df_unidades_unique[df_unidades_unique["Pedido_Clean"] != ""].drop_duplicates(subset=["Pedido_Clean"], keep="first")

        if "Unidade_Extraida" in df_merged.columns:
            df_merged = df_merged.drop(columns=["Unidade_Extraida"])

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


def obter_base_consolidada(file_acomp, file_bi, file_unid):
    """
    Consolida as fontes analíticas oficiais (Acompanhamento + BI + Unidades)
    com garantia determinística anti-duplicação e controle dos pedidos aprovados a partir da data de corte (01/11/2026),
    sem dependência de tabelas emergenciais nem exceções manuais.
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

    # Filtro da data de corte (01/11/2026): quando existirem pedidos pós-corte, prioriza estritamente
    serie_dt_corte = pd.to_datetime(df_merged['Dt_inser'], dayfirst=True, errors='coerce').dt.normalize()
    ts_corte = pd.Timestamp(DATA_CORTE).normalize()
    pedidos_pos_corte = df_merged[serie_dt_corte >= ts_corte]
    if not pedidos_pos_corte.empty:
        df_merged = pedidos_pos_corte.copy()

    # Blindagem e sincronização das colunas de unidade
    df_merged['UNIDADE'] = df_merged['Unidade_Final']
    df_merged['Unidade'] = df_merged['Unidade_Final']
    df_merged['Unidade_Real'] = df_merged['Unidade_Final']
    df_merged['Unidade_Clean'] = df_merged['Unidade_Final'].apply(limpar_texto_corrompido)

    return df_merged, df_bi_raw, df_unidades_unique, col_ped, col_uni


def processar_aba_trimestral(df_cad_raw, df_ped_raw):
    """
    Processa a consolidação trimestral automática (últimos 3 meses até a data presente)
    cruzando o cadastro mestre de unidades com os relatórios de pedidos.
    """
    if df_cad_raw is None or df_cad_raw.empty or df_ped_raw is None or df_ped_raw.empty:
        return None, None, None, None, None, None

    if isinstance(df_cad_raw.columns, pd.MultiIndex):
        df_cad_raw.columns = df_cad_raw.columns.get_level_values(-1)
    df_cad_raw.columns = [str(c).strip() for c in df_cad_raw.columns]

    if isinstance(df_ped_raw.columns, pd.MultiIndex):
        df_ped_raw.columns = df_ped_raw.columns.get_level_values(-1)
    df_ped_raw.columns = [str(c).strip() for c in df_ped_raw.columns]

    col_nome_cad = find_column(df_cad_raw, ["Nome", "unidade", "razao social", "cliente"]) or df_cad_raw.columns[0]
    col_ativo_cad = find_column(df_cad_raw, ["Ativo", "situacao", "status"])

    cols_mestre = [col_nome_cad]
    if col_ativo_cad and col_ativo_cad != col_nome_cad:
        cols_mestre.append(col_ativo_cad)

    df_unidades = df_cad_raw[cols_mestre].drop_duplicates(subset=[col_nome_cad]).copy()
    df_unidades.rename(columns={col_nome_cad: 'Unidade'}, inplace=True)
    if col_ativo_cad:
        df_unidades.rename(columns={col_ativo_cad: 'ATIVO'}, inplace=True)
        df_unidades['ATIVO'] = df_unidades['ATIVO'].astype(str).str.strip().str.upper()
    else:
        df_unidades['ATIVO'] = 'SIM'

    df_unidades['Unidade_Clean'] = df_unidades['Unidade'].apply(limpar_texto_corrompido)

    col_data_ped = find_column(df_ped_raw, ["Data", "dt", "data_pedido", "data_emissao"]) or df_ped_raw.columns[0]
    col_uni_ped = find_column(df_ped_raw, ["Unidade", "nome_unidade", "unidade_entrega"]) or (df_ped_raw.columns[3] if len(df_ped_raw.columns) > 3 else df_ped_raw.columns[-1])
    col_vlr_ped = find_column(df_ped_raw, ["Valor total", "vendas total", "total", "valor", "vlr_total"]) or (df_ped_raw.columns[9] if len(df_ped_raw.columns) > 9 else df_ped_raw.columns[-1])

    ped_data_s = df_ped_raw[col_data_ped]
    if isinstance(ped_data_s, pd.DataFrame): ped_data_s = ped_data_s.iloc[:, 0]
    
    ped_uni_s = df_ped_raw[col_uni_ped]
    if isinstance(ped_uni_s, pd.DataFrame): ped_uni_s = ped_uni_s.iloc[:, 0]
    
    ped_vlr_s = df_ped_raw[col_vlr_ped]
    if isinstance(ped_vlr_s, pd.DataFrame): ped_vlr_s = ped_vlr_s.iloc[:, 0]

    df_pedidos = pd.DataFrame({
        "Data": ped_data_s,
        "Unidade": ped_uni_s,
        "Valor": pd.to_numeric(ped_vlr_s, errors="coerce").fillna(0.0)
    })
    df_pedidos['Unidade_Clean'] = df_pedidos['Unidade'].apply(limpar_texto_corrompido)

    dt_parsed = pd.to_datetime(df_pedidos['Data'], errors='coerce', dayfirst=True).dt.normalize()

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
    dt_filtrado = pd.to_datetime(df_pedidos['Data'], errors='coerce', dayfirst=True)
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

