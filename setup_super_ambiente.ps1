Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  INICIANDO SUPER SETUP PYTHON - DATA SCIENCE & AUTOMATION" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$py = "C:\Users\filipe.marinho\AppData\Local\Programs\Python\Python312\python.exe"

# 1. Base
Write-Host "`n[1/7] Atualizando instaladores base (pip, setuptools, wheel)..." -ForegroundColor Yellow
& $py -m pip install --upgrade pip setuptools wheel --no-warn-script-location

# 2. Dados
Write-Host "`n[2/7] Instalando Manipulacao de Dados e Arquivos..." -ForegroundColor Yellow
& $py -m pip install pandas numpy scipy openpyxl xlrd xlsxwriter pyarrow fastparquet --no-warn-script-location

# 3. Dashboards
Write-Host "`n[3/7] Instalando Dashboards e Visualizacao..." -ForegroundColor Yellow
& $py -m pip install streamlit watchdog plotly matplotlib seaborn bokeh --no-warn-script-location

# 4. Automacao
Write-Host "`n[4/7] Instalando Automacao e Web Scraping..." -ForegroundColor Yellow
& $py -m pip install requests beautifulsoup4 lxml python-dotenv --no-warn-script-location

# 5. Machine Learning
Write-Host "`n[5/7] Instalando Inteligencia Artificial e Machine Learning..." -ForegroundColor Yellow
& $py -m pip install scikit-learn xgboost statsmodels --no-warn-script-location

# 6. Banco de Dados
Write-Host "`n[6/7] Instalando Conectores de Banco de Dados..." -ForegroundColor Yellow
& $py -m pip install sqlalchemy pymysql psycopg2-binary --no-warn-script-location

# 7. Jupyter
Write-Host "`n[7/7] Instalando Ambiente de Experimentacao..." -ForegroundColor Yellow
& $py -m pip install jupyter ipykernel --no-warn-script-location

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  SUPER SETUP CONCLUIDO COM SUCESSO!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
