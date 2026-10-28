"""
Módulo de Gestão de Banco de Dados Analítico Local (DuckDB)
Responsável por conexão, inicialização, ingestão padronizada e consultas analíticas de alta performance.
"""

import os
import io
import re
import datetime
import zipfile
import duckdb
import pandas as pd
import numpy as np
from unidecode import unidecode
import xlrd

try:
    import ftfy
except ImportError:
    ftfy = None

# Diretórios base do projeto
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "dashboard_data.duckdb")

# Garante que a pasta 'data' exista
os.makedirs(DATA_DIR, exist_ok=True)


# -------------------------------------------------------------------------
# FUNÇÕES DE HIGIENIZAÇÃO E SANITIZAÇÃO
# -------------------------------------------------------------------------

def limpar_texto_corrompido(texto):
    """
    Higieniza e normaliza textos corrigindo problemas de encoding/mojibake.
    """
    if pd.isna(texto) or texto is None:
        return ""
    
    txt = str(texto).strip()
    if txt.lower() in ["nan", "none", "null", "-", ""]:
        return ""
    
    padroes_mojibake = {
        "Ã±Ã£O": "CAO", "Ã±Ã£o": "cao", "Ã§Ã£o": "cao", "Ã§Ã£O": "CAO",
        "Ã°": "", "Â°": "", "Âº": "", "º": "", "°": "", "ª": "", 
        "\ufffd": "", "ð": "", "Ð": "", "?": ""
    }
    for err, fix in padroes_mojibake.items():
        txt = txt.replace(err, fix)

    if ftfy is not None:
        try:
            txt = ftfy.fix_text(txt)
        except Exception:
            pass

    try:
        txt = txt.encode('latin1').decode('utf-8')
    except Exception:
        try:
            txt = txt.encode('cp1252').decode('utf-8')
        except Exception:
            pass
        
    for err, fix in padroes_mojibake.items():
        txt = txt.replace(err, fix)

    txt = unidecode(txt).upper()
    txt = re.sub(r'[^A-Z0-9\s]', ' ', txt)
    txt = re.sub(r'\s+', ' ', txt).strip()
    return txt


def limpar_id(val):
    """
    Padroniza identificadores (Pedido, NF) para string numérica pura:
    - Remove sufixo decimal '.0' gerado pelo Excel
    - Extrai estritamente os dígitos numéricos
    """
    if pd.isna(val) or val is None:
        return ""
    texto = str(val).strip()
    texto = re.sub(r'\.0$', '', texto)
    match = re.search(r'(\d+)', texto)
    return match.group(1) if match else texto


def extrair_unidade_cirurgica(texto):
    """
    Extrai com alta precisão o nome da unidade a partir do campo 'Local de entrega':
    1. Se houver '---- NOME DA UNIDADE' no final da string, extrai diretamente.
    2. Caso contrário, recorta cirurgicamente logo após 'PED.: XXXXX - CÓDIGO - ' e antes do logradouro.
    """
    if pd.isna(texto) or texto is None:
        return None
    t = str(texto).strip()
    
    # 1. Estratégia de ouro: Se termina com ---- UNIDADE
    m_fim = re.search(r'[-–—]{3,}\s*([A-Za-z0-9À-ÿ\s\.\-\(\)\/]+)$', t)
    if m_fim:
        res = m_fim.group(1).strip()
        if len(res) > 3 and not re.match(r'^(?:RS|SC|PR|SP|RJ|MG|BRASIL)$', res, re.I):
            return res

    # 2. Estratégia de corte cirúrgico no miolo (após PED.: XXXXX - CÓDIGO - )
    m = re.search(r'PED\.?:\s*\d+\s*[-–—]+\s*(?:\d+|-)\s*[-–—]+', t, re.IGNORECASE)
    if m:
        miolo = t[m.end():].strip()
    else:
        m2 = re.search(r'PED\.?:\s*\d+\s*[-–—]+', t, re.IGNORECASE)
        miolo = t[m2.end():].strip() if m2 else t
    
    # Corta antes do logradouro
    padroes_corte = [
        r'\s*[-–—]+\s*RUA\b', r'\s*[-–—]+\s*RODOVIA\b', r'\s*[-–—]+\s*ROD\b',
        r'\s*[-–—]+\s*AV\b', r'\s*[-–—]+\s*AVENIDA\b', r'\s*[-–—]+\s*PRACA\b',
        r'\s*[-–—]+\s*PRAÇA\b', r'\s*[-–—]+\s*ESTRADA\b', r'\s*[-–—]+\s*TRAVESSA\b',
        r'\s*[-–—]+\s*ALAMEDA\b', r'\s*[-–—]+\s*ACESSO\b', r'\s*[-–—]+\s*BR-\d+',
        r'\s*[-–—]+\s*RS-\d+'
    ]
    for p in padroes_corte:
        partes = re.split(p, miolo, flags=re.IGNORECASE)
        if len(partes) > 1:
            miolo = partes[0]
            break
            
    miolo = re.sub(r'[-–—\s]+$', '', miolo).strip()
    return miolo if miolo else None


def normalizar_nome_coluna(col_name):
    """
    Gera um identificador de coluna limpo e compatível com SQL.
    Ex: 'Preço Venda (R$)' -> 'preco_venda_r'
    """
    if isinstance(col_name, tuple):
        col_name = " ".join([str(c) for c in col_name if str(c) != "" and not str(c).startswith("Unnamed")])
    c = str(col_name).strip()
    c = unidecode(c).lower()
    c = re.sub(r'[^a-z0-9_]+', '_', c)
    c = re.sub(r'_+', '_', c).strip('_')
    return c if c else "coluna"


# -------------------------------------------------------------------------
# LEITURA RESILIENTE DE PLANILHAS (EXCEL, CSV E HTML DISFARÇADO)
# -------------------------------------------------------------------------

def load_sheet(file_obj):
    """
    Lê arquivos .xlsx, .xls ou .csv com fallbacks automáticos para planilhas
    geradas por ERPs exportadas em formato HTML disfarçado de Excel.
    """
    filename = ""
    if isinstance(file_obj, str):
        filename = os.path.basename(file_obj).lower()
        with open(file_obj, "rb") as f:
            content = f.read()
        buf = io.BytesIO(content)
        buf.name = file_obj
        file_obj = buf
    else:
        filename = getattr(file_obj, "name", "").strip().lower()

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    
    # 1. Arquivos CSV
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
    except (xlrd.biffh.XLRDError, zipfile.BadZipFile, ValueError, Exception):
        # Fallback 1: Ler como tabela HTML disfarçada de Excel
        try:
            file_obj.seek(0)
            tables = pd.read_html(file_obj, decimal=",", thousands=".")
            if tables and len(tables) > 0:
                df = max(tables, key=lambda t: t.shape[0] * t.shape[1])
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = df.columns.get_level_values(-1)
                return df
        except Exception:
            pass

        # Fallback 2: Decodificar HTML com encodings latinos
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
                            df.columns = df.columns.get_level_values(-1)
                        return df
                except Exception:
                    continue
        except Exception:
            pass

        # Fallback 3: Tentar alternar engine Excel
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

        raise ValueError(f"Não foi possível ler o arquivo '{filename}'. Formato desconhecido ou corrompido.")


# -------------------------------------------------------------------------
# CONEXÃO E INICIALIZAÇÃO DO BANCO DUCKDB
# -------------------------------------------------------------------------

def get_connection(read_only=True):
    """
    Retorna uma conexão com o banco DuckDB persistido em data/dashboard_data.duckdb.
    Para consultas no Streamlit, use preferencialmente read_only=True para alta concorrência.
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DB_PATH):
        try:
            c = duckdb.connect(DB_PATH, read_only=False)
            c.close()
        except Exception:
            pass

    if read_only:
        try:
            return duckdb.connect(DB_PATH, read_only=True)
        except Exception:
            try:
                return duckdb.connect(DB_PATH, read_only=False)
            except Exception as e_ro:
                raise e_ro
            
    try:
        return duckdb.connect(DB_PATH, read_only=False)
    except Exception as e:
        if os.path.exists(DB_PATH):
            try:
                return duckdb.connect(DB_PATH, read_only=True)
            except Exception:
                pass
        raise e


def init_db():
    """
    Cria as tabelas base do sistema e a tabela de controle de metadados de ingestão.
    Se o arquivo de banco já existir e contiver dados, não tenta travar o arquivo com escrita.
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(DB_PATH):
        try:
            if table_exists("ingestion_metadata"):
                return True
        except Exception:
            pass

    try:
        with get_connection(read_only=False) as conn:
            conn.execute("""
                CREATE SEQUENCE IF NOT EXISTS seq_ingestion_id START 1;
                
                CREATE TABLE IF NOT EXISTS ingestion_metadata (
                    id INTEGER PRIMARY KEY DEFAULT nextval('seq_ingestion_id'),
                    table_name VARCHAR NOT NULL,
                    file_name VARCHAR,
                    row_count INTEGER,
                    mode VARCHAR,
                    ingested_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    status VARCHAR,
                    details VARCHAR
                );
            """)
        return True
    except Exception:
        return False


def table_exists(table_name):
    """
    Verifica se uma tabela ou view existe no DuckDB.
    """
    if not os.path.exists(DB_PATH):
        return False
    try:
        with get_connection(read_only=True) as conn:
            res = conn.execute("""
                SELECT COUNT(*) 
                FROM information_schema.tables 
                WHERE table_name = ?
            """, [table_name.lower()]).fetchone()
            return res[0] > 0 if res else False
    except Exception:
        return False


def get_table_row_count(table_name):
    """
    Retorna o número total de linhas de uma tabela.
    """
    if not table_exists(table_name):
        return 0
    try:
        with get_connection(read_only=True) as conn:
            res = conn.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()
            return res[0] if res else 0
    except Exception:
        return 0


# -------------------------------------------------------------------------
# INGESTÃO PADRONIZADA DE ALTA PERFORMANCE
# -------------------------------------------------------------------------

def ingest_sheet(file_path_or_buffer, table_name, mode="replace", original_filename=None):
    """
    Lê a planilha, higieniza os dados e grava diretamente no DuckDB com alta performance.
    
    Parâmetros:
    - file_path_or_buffer: caminho do arquivo ou objeto BytesIO / UploadedFile
    - table_name: nome da tabela no DuckDB (ex: 'acompanhamento', 'dados_bi', 'unidades')
    - mode: 'replace' (substitui a tabela) ou 'append' (adiciona registros)
    - original_filename: nome amigável do arquivo para registro de log
    
    Retorna:
    - dict com métricas e status da carga
    """
    init_db()
    
    # Determina o nome do arquivo
    if original_filename:
        file_name = original_filename
    elif isinstance(file_path_or_buffer, str):
        file_name = os.path.basename(file_path_or_buffer)
    else:
        file_name = getattr(file_path_or_buffer, "name", "planilha_upload")

    # 1. Leitura com fallbacks
    df = load_sheet(file_path_or_buffer)
    if df is None or df.empty:
        raise ValueError(f"O arquivo '{file_name}' está vazio ou não pôde ser lido.")

    # Achata MultiIndex se houver
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(-1)

    # 2. Remoção de linhas de rodapé típicas de ERP (ex: 'Total', 'Totais', 'Filtros aplicados') e linhas vazias
    for col in df.columns[:5]:
        mask_rodape = df[col].astype(str).str.contains(r'^(?:Total|Totais|Filtros\s*aplicados)\b', case=False, na=False)
        if mask_rodape.any():
            df = df[~mask_rodape]
    df = df.dropna(how='all')

    # 3. Padronização e sanitização das colunas
    # Mapeia colunas para nomes legíveis e originais
    colunas_limpas = []
    nomes_vistos = {}
    for c in df.columns:
        c_str = str(c).strip()
        # Se vier com cabeçalho sujo tipo 'RELATÓRIO DE PEDIDOS PORTAL.Data', pega a parte final
        if any(p in c_str.upper() for p in ["RELATÓRIO", "RELATORIO", "PORTAL."]) and "." in c_str:
            c_str = c_str.split(".")[-1].strip()
        
        # Garante unicidade dos nomes das colunas
        base = c_str if c_str else "coluna"
        if base in nomes_vistos:
            nomes_vistos[base] += 1
            c_str = f"{base}_{nomes_vistos[base]}"
        else:
            nomes_vistos[base] = 0
            
        colunas_limpas.append(c_str)

    df.columns = colunas_limpas

    # 4. Tipagem e limpeza de dados específica para colunas conhecidas
    for col in df.columns:
        col_lower = str(col).lower()
        
        # Limpeza de chaves de Pedido e Nota Fiscal
        if any(p in col_lower for p in ["pedido", "codigo pedido", "n_pedido", "nr_pedido"]):
            df[col] = df[col].apply(limpar_id)
        elif any(p in col_lower for p in ["nota", "n_nota", "numero_nota", "n_nota_obs"]):
            nota_num = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
            df[col] = nota_num.astype(str).replace("0", "")
        elif col_lower in ["cod", "codigo", "cod_produto", "cod_item"]:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
        # Limpeza de valores monetários e numéricos (R$, vírgulas)
        elif any(p in col_lower for p in ["vendas total", "preco", "preço", "custo", "vlr_total", "valor"]):
            if df[col].dtype == object:
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.replace("R$", "", case=False)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .str.strip()
                )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        # Limpeza de quantidades
        elif any(p in col_lower for p in ["qtd", "quantidade", "quant"]):
            if df[col].dtype == object:
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .str.strip()
                )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Conversão de datas
        elif any(p in col_lower for p in ["dt_inser", "data", "dt_nota_hist", "data_cadastro", "dt_separ"]):
            try:
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
            except Exception:
                pass

        # Strings gerais: remoção de mojibake se for texto de unidade ou cliente
        elif any(p in col_lower for p in ["unidade", "nome", "cliente", "razao"]):
            if df[col].dtype == object:
                df[col] = df[col].astype(str).apply(limpar_texto_corrompido)

    # Tratamento automático de 'Local de entrega' (extração cirúrgica de Pedido e Unidade)
    col_local_cand = next((c for c in df.columns if "entrega" in str(c).lower() or "local" in str(c).lower()), None)
    if col_local_cand:
        if "Pedido" not in df.columns and "Codigo pedido" not in df.columns:
            df["Pedido"] = df[col_local_cand].astype(str).str.extract(r'PED\.?:\s*(\d+)', expand=False)
            df["Pedido_Clean"] = df["Pedido"].apply(limpar_id)
        
        unid_rec = df[col_local_cand].apply(extrair_unidade_cirurgica)
        if "Unidade" not in df.columns:
            df["Unidade"] = unid_rec
        else:
            mask_subst = df["Unidade"].isna() | df["Unidade"].astype(str).str.strip().isin(["", "nan", "None"]) | df["Unidade"].astype(str).str.contains(r'INCONFID[EÊ]NCIA\s+LOCADORA', case=False, na=False)
            df.loc[mask_subst & unid_rec.notna(), "Unidade"] = unid_rec

    # Criação de chaves auxiliares pré-computadas
    col_nota_cand = [c for c in df.columns if any(p in str(c).lower() for p in ["nota", "n_nota_obs", "n_nota"]) and "dt" not in str(c).lower()]
    if col_nota_cand and "NF_Clean" not in df.columns:
        df["NF_Clean"] = df[col_nota_cand[0]].astype(str).str.replace(r'\D', '', regex=True).str.lstrip('0')

    col_ped_cand = [c for c in df.columns if any(p in str(c).lower() for p in ["pedido", "codigo pedido"]) and "item" not in str(c).lower()]
    if col_ped_cand and "Pedido_Clean" not in df.columns:
        df["Pedido_Clean"] = df[col_ped_cand[0]].astype(str).apply(limpar_id)

    # 5. Ingestão de Alta Performance no DuckDB
    table_clean = re.sub(r'[^a-zA-Z0-9_]', '_', table_name).lower().strip('_')
    row_count = len(df)
    
    with get_connection(read_only=False) as conn:
        conn.register("df_ingest_buffer", df)
        
        t_exists = table_exists(table_clean)
        
        if mode == "append" and t_exists:
            # Garante colunas compatíveis ou insere por posição
            conn.execute(f"INSERT INTO {table_clean} SELECT * FROM df_ingest_buffer")
        else:
            conn.execute(f"CREATE OR REPLACE TABLE {table_clean} AS SELECT * FROM df_ingest_buffer")
            
        conn.unregister("df_ingest_buffer")
        
        # Registra metadados
        conn.execute("""
            INSERT INTO ingestion_metadata (table_name, file_name, row_count, mode, status, details)
            VALUES (?, ?, ?, ?, 'SUCCESS', ?)
        """, [table_clean, file_name, row_count, mode, f"Colunas: {len(df.columns)}"])

    return {
        "status": "success",
        "table_name": table_clean,
        "file_name": file_name,
        "row_count": row_count,
        "columns_count": len(df.columns),
        "columns": list(df.columns),
        "mode": mode,
        "timestamp": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    }


# -------------------------------------------------------------------------
# EXECUÇÃO DE QUERIES COM RETORNO EM DATAFRAME
# -------------------------------------------------------------------------

def run_query(query, params=None):
    """
    Executa uma consulta SQL agregada no DuckDB e retorna diretamente
    um DataFrame do Pandas pronto para visualização.
    
    Abre e fecha a conexão com segurança para evitar travamentos de arquivo no Windows.
    """
    if not os.path.exists(DB_PATH):
        init_db()

    # Se a query for apenas de leitura (SELECT / WITH), usa read_only=True
    query_stripped = query.strip().upper()
    is_read_only = query_stripped.startswith("SELECT") or query_stripped.startswith("WITH") or query_stripped.startswith("DESCRIBE")

    with get_connection(read_only=is_read_only) as conn:
        if params is not None:
            rel = conn.execute(query, params)
        else:
            rel = conn.execute(query)
            
        if is_read_only:
            return rel.df()
        return pd.DataFrame()


# -------------------------------------------------------------------------
# CONSULTAS DE METADADOS E HISTÓRICO DE CARGA
# -------------------------------------------------------------------------

def get_last_ingestion_info():
    """
    Retorna o histórico de cargas registradas no sistema.
    """
    if not table_exists("ingestion_metadata"):
        return pd.DataFrame(columns=["id", "table_name", "file_name", "row_count", "mode", "ingested_at", "status"])
    
    query = """
        SELECT 
            id,
            table_name AS Tabela,
            file_name AS Arquivo,
            row_count AS Linhas,
            mode AS Modo,
            strftime(ingested_at, '%d/%m/%Y %H:%M:%S') AS Data_Carga,
            status AS Status
        FROM ingestion_metadata
        ORDER BY id DESC
    """
    return run_query(query)


def get_tables_summary():
    """
    Retorna um resumo de todas as tabelas criadas no DuckDB com total de linhas.
    """
    if not os.path.exists(DB_PATH):
        return pd.DataFrame(columns=["Tabela", "Linhas"])
    
    query = """
        SELECT table_name AS Tabela, -1 AS Linhas
        FROM information_schema.tables
        WHERE table_schema = 'main' AND table_name != 'ingestion_metadata'
    """
    try:
        df_tables = run_query(query)
        linhas = []
        for tbl in df_tables["Tabela"]:
            cnt = get_table_row_count(tbl)
            linhas.append(cnt)
        df_tables["Linhas"] = linhas
        return df_tables
    except Exception:
        return pd.DataFrame(columns=["Tabela", "Linhas"])


def reset_database():
    """
    Exclui o arquivo do banco DuckDB para resetar completamente a base de dados.
    """
    if os.path.exists(DB_PATH):
        try:
            os.remove(DB_PATH)
            init_db()
            return True
        except Exception as e:
            return False
    return True


# -------------------------------------------------------------------------
# INGESTÃO AUTOMÁTICA DOS ARQUIVOS PADRÃO EXISTENTES (BOOTSTRAP)
# -------------------------------------------------------------------------

def ingest_default_files_if_needed():
    """
    Verifica se as tabelas padrão existem no DuckDB.
    Se o banco ainda não foi populado e os arquivos .xlsx/.xls existem no projeto,
    executa a ingestão inicial automaticamente.
    """
    init_db()
    
    mapeamento_arquivos = {
        "acompanhamento": os.path.join(BASE_DIR, "acompanhamento.xlsx"),
        "dados_bi": os.path.join(BASE_DIR, "dados_bi.xlsx"),
        "unidades": os.path.join(BASE_DIR, "unidades.xlsx"),
        "listagem_unidades": os.path.join(BASE_DIR, "listagem_unidades.xls") if os.path.exists(os.path.join(BASE_DIR, "listagem_unidades.xls")) else os.path.join(BASE_DIR, "listagem_unidades.xlsx"),
        "nf_emergencial": os.path.join(BASE_DIR, "nf_emergencial.xlsx") if os.path.exists(os.path.join(BASE_DIR, "nf_emergencial.xlsx")) else (os.path.join(BASE_DIR, "nf_emergencial.csv") if os.path.exists(os.path.join(BASE_DIR, "nf_emergencial.csv")) else None)
    }

    tabelas_ingestao = []
    for tabela, caminho in mapeamento_arquivos.items():
        if caminho and os.path.exists(caminho):
            # Se a tabela não existe ou está vazia, ingere
            if not table_exists(tabela) or get_table_row_count(tabela) == 0:
                try:
                    ingest_sheet(caminho, table_name=tabela, mode="replace", original_filename=os.path.basename(caminho))
                    tabelas_ingestao.append(tabela)
                except Exception as e:
                    print(f"Aviso ao ingerir {tabela}: {e}")

    return tabelas_ingestao


def sync_files_if_modified(force=False):
    """
    Sincroniza os arquivos de planilha da pasta raiz ou da pasta INC com as tabelas do DuckDB.
    Se o arquivo físico existir e tiver data de modificação mais recente que
    a última ingestão registrada para a tabela (ou se force=True), executa a ingestão.
    """
    init_db()
    
    # 1. Verifica se há arquivos mais novos na subpasta INC e traz para a raiz
    inc_dir = os.path.join(BASE_DIR, "INC")
    if os.path.exists(inc_dir):
        # Acompanhamento em INC
        inc_acomp = os.path.join(inc_dir, "acompanhamento.xlsx")
        raiz_acomp = os.path.join(BASE_DIR, "acompanhamento.xlsx")
        if os.path.exists(inc_acomp):
            if not os.path.exists(raiz_acomp) or (os.path.getmtime(inc_acomp) > os.path.getmtime(raiz_acomp) + 1.0):
                try:
                    import shutil
                    shutil.copy2(inc_acomp, raiz_acomp)
                    print("[*] Base 'acompanhamento.xlsx' atualizada a partir da pasta INC.")
                except Exception:
                    pass

        # Unidades em INC
        for ext in [".xls", ".xlsx"]:
            inc_u = os.path.join(inc_dir, f"unidades{ext}")
            raiz_u = os.path.join(BASE_DIR, f"unidades{ext}")
            if os.path.exists(inc_u):
                if not os.path.exists(raiz_u) or (os.path.getmtime(inc_u) > os.path.getmtime(raiz_u) + 1.0):
                    try:
                        import shutil
                        shutil.copy2(inc_u, raiz_u)
                        print(f"[*] Base 'unidades{ext}' atualizada a partir da pasta INC.")
                    except Exception:
                        pass

        # Dados BI em INC (se houver novo, faz merge inteligente com histórico)
        inc_bi = os.path.join(inc_dir, "dados_bi.xlsx")
        raiz_bi = os.path.join(BASE_DIR, "dados_bi.xlsx")
        if os.path.exists(inc_bi) and os.path.exists(raiz_bi):
            if os.path.getmtime(inc_bi) > os.path.getmtime(raiz_bi) + 1.0:
                try:
                    df_old = pd.read_excel(raiz_bi)
                    df_new = pd.read_excel(inc_bi)
                    df_total = pd.concat([df_old, df_new], ignore_index=True).drop_duplicates()
                    df_total.to_excel(raiz_bi, index=False)
                    print("[*] Base 'dados_bi.xlsx' unificada com os novos dados de INC.")
                except Exception as e:
                    print(f"Aviso no merge de dados_bi de INC: {e}")

    mapeamento_arquivos = {
        "acompanhamento": os.path.join(BASE_DIR, "acompanhamento.xlsx"),
        "dados_bi": os.path.join(BASE_DIR, "dados_bi.xlsx"),
        "unidades": os.path.join(BASE_DIR, "unidades.xls") if os.path.exists(os.path.join(BASE_DIR, "unidades.xls")) else os.path.join(BASE_DIR, "unidades.xlsx"),
        "listagem_unidades": os.path.join(BASE_DIR, "listagem_unidades.xls") if os.path.exists(os.path.join(BASE_DIR, "listagem_unidades.xls")) else os.path.join(BASE_DIR, "listagem_unidades.xlsx"),
        "nf_emergencial": os.path.join(BASE_DIR, "nf_emergencial.xlsx") if os.path.exists(os.path.join(BASE_DIR, "nf_emergencial.xlsx")) else (os.path.join(BASE_DIR, "nf_emergencial.csv") if os.path.exists(os.path.join(BASE_DIR, "nf_emergencial.csv")) else None),
        "monitoramento": os.path.join(BASE_DIR, "monitoramento.xlsx") if os.path.exists(os.path.join(BASE_DIR, "monitoramento.xlsx")) else (os.path.join(BASE_DIR, "INC", "monitoramento", "monitoramento.xlsx") if os.path.exists(os.path.join(BASE_DIR, "INC", "monitoramento", "monitoramento.xlsx")) else None)
    }
    
    last_ingestion_map = {}
    if table_exists("ingestion_metadata"):
        try:
            df_epoch = run_query("SELECT table_name, max(epoch(ingested_at)) as last_epoch FROM ingestion_metadata GROUP BY table_name")
            if not df_epoch.empty:
                for _, row in df_epoch.iterrows():
                    last_ingestion_map[str(row["table_name"])] = float(row["last_epoch"]) if pd.notna(row["last_epoch"]) else 0.0
        except Exception:
            pass

    tabelas_sincronizadas = []
    for tabela, caminho in mapeamento_arquivos.items():
        if caminho and os.path.exists(caminho):
            mtime = os.path.getmtime(caminho)
            last_ingested = last_ingestion_map.get(tabela, 0.0)
            
            # Se for forçado, se a tabela não existir, estiver vazia, ou se o arquivo for mais novo que a última ingestão
            deve_ingerir = force or (not table_exists(tabela)) or (get_table_row_count(tabela) == 0) or (mtime > last_ingested + 2.0)
            
            if deve_ingerir:
                try:
                    print(f"[*] Ingerindo {os.path.basename(caminho)} na tabela '{tabela}'...")
                    ingest_sheet(caminho, table_name=tabela, mode="replace", original_filename=os.path.basename(caminho))
                    tabelas_sincronizadas.append(tabela)
                    print(f"[OK] Tabela '{tabela}' atualizada com sucesso.")
                except Exception as e:
                    print(f"[ERRO] Falha ao sincronizar {tabela}: {e}")
                    
    return tabelas_sincronizadas


def get_mapa_global_unidades():
    """
    Retorna um dicionário consolidado { 'Pedido': 'Nome_Real_Unidade' }
    combinando todas as fontes do DuckDB (unidades, monitoramento, listagem_unidades).
    Garante que nenhum pedido fique associado à razão social genérica da Inconfidência.
    """
    mapa = {}
    
    # 1. Da tabela unidades (relatório do portal)
    if table_exists("unidades"):
        try:
            df_u = run_query("SELECT DISTINCT Pedido_Clean, Unidade FROM unidades WHERE Unidade IS NOT NULL")
            for _, r in df_u.iterrows():
                p = str(r['Pedido_Clean']).strip()
                u = str(r['Unidade']).strip()
                if p and u and u.lower() not in ['nan', 'none', 'não informada', ''] and not re.search(r'INCONFID[EÊ]NCIA\s+LOCADORA', u, re.I):
                    mapa[p] = u
        except Exception:
            pass

    # 2. Da tabela monitoramento (recorte cirúrgico de 'Local de entrega')
    if table_exists("monitoramento"):
        try:
            df_m = run_query("SELECT * FROM monitoramento")
            col_loc = next((c for c in df_m.columns if 'entrega' in str(c).lower()), None)
            if col_loc:
                df_m['p_ext'] = df_m[col_loc].astype(str).str.extract(r'PED\.?:\s*(\d+)', expand=False)
                df_m['u_ext'] = df_m[col_loc].apply(extrair_unidade_cirurgica)
                for _, r in df_m.dropna(subset=['p_ext', 'u_ext']).iterrows():
                    p = str(r['p_ext']).strip()
                    u = str(r['u_ext']).strip()
                    if p and u and not re.search(r'INCONFID[EÊ]NCIA\s+LOCADORA', u, re.I):
                        mapa[p] = u
        except Exception:
            pass
            
    # 3. Exceções Manuais Prioritárias
    mapa.update(EXCECOES_MANUAIS_UNIDADES)
    return mapa


# -------------------------------------------------------------------------
# TABELA DE EXCEÇÕES MANUAIS DE UNIDADES ESPECÍFICAS
# -------------------------------------------------------------------------
EXCECOES_MANUAIS_UNIDADES = {
    "1232704": "SERPRO - REFATURAMENTO VALOR DE CAFÉ ESTAVA ERRADO",
    "1234043": "BANRISUL AV DO FORTE - EMITIDO INTERNO POR ERRO DE SISTEMA",
    "1234574": "GRUPO CABURGO - CITROEN E JAECOO CANOAS"
}



