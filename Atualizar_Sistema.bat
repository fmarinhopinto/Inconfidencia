@echo off
setlocal
cd /d "%~dp0"
title Atualizador Automatico - Grupo Fortpel e Inconfidencia
color 0b

echo ======================================================================
echo    ATUALIZADOR AUTOMATICO - GRUPO FORTPEL E INCONFIDENCIA
echo    Repositorio: https://github.com/fmarinhopinto/Inconfidencia
echo ======================================================================
echo.

set "GIT_EXE=C:\Users\filipe.marinho\AppData\Local\Programs\Git\cmd\git.exe"
if not exist "%GIT_EXE%" (
    set "GIT_EXE=git"
)

set "PY_EXE=C:\Users\filipe.marinho\AppData\Local\Programs\Python\Python312\python.exe"
if exist "%PY_EXE%" (
    echo [*] Verificando se ha novas planilhas para o DuckDB...
    "%PY_EXE%" -c "from modules.db_manager import sync_files_if_modified; sync_files_if_modified()"
)

echo [*] Preparando arquivos e banco de dados...
"%GIT_EXE%" add .

echo [*] Registrando alteracoes locais...
"%GIT_EXE%" commit -m "Atualizacao automatica - %date% %time%" >nul 2>&1

echo.
echo [*] Sincronizando com o GitHub...
echo (Se for a primeira vez, podera abrir o navegador para autorizar)
echo.

"%GIT_EXE%" push origin main

if %ERRORLEVEL% equ 0 (
    echo.
    color 0a
    echo ======================================================================
    echo   [SUCESSO] Sistema sincronizado com o GitHub com exito!
    echo   O Streamlit Cloud ja detectou e esta atualizando online.
    echo ======================================================================
) else (
    echo.
    color 0c
    echo ======================================================================
    echo   [ATENCAO] Nao foi possivel sincronizar automaticamente.
    echo   Verifique sua conexao ou login no GitHub.
    echo ======================================================================
)

echo.
pause
