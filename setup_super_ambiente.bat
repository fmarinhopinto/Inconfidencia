@echo off
chcp 65001 >nul
echo ============================================================
echo   INICIANDO SUPER SETUP PYTHON - DATA SCIENCE & AUTOMATION
echo ============================================================

set PY_EXE="C:\Users\filipe.marinho\AppData\Local\Programs\Python\Python312\python.exe"

echo.
echo [1/7] Atualizando instaladores base (pip, setuptools, wheel)...
%PY_EXE% -m pip install --upgrade pip setuptools wheel --no-warn-script-location

echo.
echo [2/7] Instalando Manipulação de Dados e Arquivos (pandas, numpy, scipy, openpyxl, xlrd, xlsxwriter, pyarrow, fastparquet)...
%PY_EXE% -m pip install pandas numpy scipy openpyxl xlrd xlsxwriter pyarrow fastparquet --no-warn-script-location

echo.
echo [3/7] Instalando Dashboards e Visualização (streamlit, watchdog, plotly, matplotlib, seaborn, bokeh)...
%PY_EXE% -m pip install streamlit watchdog plotly matplotlib seaborn bokeh --no-warn-script-location

echo.
echo [4/7] Instalando Automação e Web Scraping (requests, beautifulsoup4, lxml, python-dotenv)...
%PY_EXE% -m pip install requests beautifulsoup4 lxml python-dotenv --no-warn-script-location

echo.
echo [5/7] Instalando Machine Learning e IA (scikit-learn, xgboost, statsmodels)...
%PY_EXE% -m pip install scikit-learn xgboost statsmodels --no-warn-script-location

echo.
echo [6/7] Instalando Conectores de Banco de Dados (sqlalchemy, pymysql, psycopg2-binary)...
%PY_EXE% -m pip install sqlalchemy pymysql psycopg2-binary --no-warn-script-location

echo.
echo [7/7] Instalando Ambiente de Experimentação (jupyter, ipykernel)...
%PY_EXE% -m pip install jupyter ipykernel --no-warn-script-location

echo.
echo ============================================================
echo   SUPER SETUP CONCLUÍDO COM SUCESSO!
echo ============================================================
