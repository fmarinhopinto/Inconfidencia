@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

:: Garante diretorio de logs
if not exist "logs" mkdir "logs"
set "LOG_FILE=%~dp0logs\automacao_diaria.log"

echo ====================================================================== >> "%LOG_FILE%"
echo [%DATE% %TIME%] INICIANDO ATUALIZACAO AUTOMATICA DO DASHBOARD >> "%LOG_FILE%"

set "GIT_EXE=C:\Users\filipe.marinho\AppData\Local\Programs\Git\cmd\git.exe"
if not exist "%GIT_EXE%" set "GIT_EXE=git"

set "PY_EXE=C:\Users\filipe.marinho\AppData\Local\Programs\Python\Python312\python.exe"

:: 1. Sincroniza planilhas para o banco DuckDB se houver novas
if exist "%PY_EXE%" (
    echo [%DATE% %TIME%] Verificando planilhas e atualizando DuckDB... >> "%LOG_FILE%"
    "%PY_EXE%" -c "from modules.db_manager import sync_files_if_modified; res = sync_files_if_modified(); print('Sincronizados:', res)" >> "%LOG_FILE%" 2>&1
)

:: 2. Adiciona arquivos alterados
echo [%DATE% %TIME%] Coletando arquivos alterados... >> "%LOG_FILE%"
"%GIT_EXE%" add . >> "%LOG_FILE%" 2>&1

:: 3. Verifica se ha algo para commitar
"%GIT_EXE%" diff-index --quiet HEAD --
if %ERRORLEVEL% equ 0 (
    echo [%DATE% %TIME%] Nenhuma alteracao pendente detectada hoje. >> "%LOG_FILE%"
) else (
    echo [%DATE% %TIME%] Registrando alteracoes no Git... >> "%LOG_FILE%"
    "%GIT_EXE%" commit -m "Auto: sincronizacao matinal - %DATE% %TIME%" >> "%LOG_FILE%" 2>&1
    
    echo [%DATE% %TIME%] Enviando para o GitHub e nuvem Streamlit... >> "%LOG_FILE%"
    "%GIT_EXE%" push origin main >> "%LOG_FILE%" 2>&1
    
    if !ERRORLEVEL! equ 0 (
        echo [%DATE% %TIME%] SUCESSO: Dashboard na nuvem atualizado com exito! >> "%LOG_FILE%"
    ) else (
        echo [%DATE% %TIME%] AVISO: Nao foi possivel enviar para o GitHub no momento. >> "%LOG_FILE%"
    )
)

echo [%DATE% %TIME%] FINALIZADO COM SUCESSO. >> "%LOG_FILE%"
echo ====================================================================== >> "%LOG_FILE%"
exit /b 0
