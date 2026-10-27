@echo off
chcp 65001 >nul
title Atualizador Automático - Dashboard Grupo Fortpel & Inconfidência
color 0b

echo ======================================================================
echo    ATUALIZADOR AUTOMÁTICO - GRUPO FORTPEL & INCONFIDÊNCIA
echo    Repositório: https://github.com/fmarinhopinto/Inconfidencia
echo ======================================================================
echo.

:: 1. Localiza o executável do Git
set GIT_CMD=git
where git >nul 2>&1
if %ERRORLEVEL% neq 0 (
    if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" (
        set GIT_CMD="%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
    ) else if exist "C:\Program Files\Git\cmd\git.exe" (
        set GIT_CMD="C:\Program Files\Git\cmd\git.exe"
    ) else (
        color 0c
        echo [ERRO] O Git não foi encontrado no sistema!
        echo Por favor, verifique a instalação do Git.
        echo.
        pause
        exit /b 1
    )
)

echo [OK] Git detectado com sucesso.
echo.

:: 2. Garante vínculo com o repositório oficial
set REPO_URL=https://github.com/fmarinhopinto/Inconfidencia.git
%GIT_CMD% remote get-url origin >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [*] Vinculando repositório remoto: %REPO_URL%
    %GIT_CMD% remote add origin %REPO_URL%
) else (
    %GIT_CMD% remote set-url origin %REPO_URL%
)

:: 3. Adiciona arquivos modificados e banco de dados DuckDB
echo [*] Coletando arquivos alterados e banco de dados...
%GIT_CMD% add .

:: 4. Verifica se há algo para commitar
%GIT_CMD% diff-index --quiet HEAD --
if %ERRORLEVEL% equ 0 (
    echo [INFO] Nenhuma nova alteração nos arquivos locais.
) else (
    echo [*] Registrando alterações...
    %GIT_CMD% commit -m "Atualizacao automatica - %date% %time%"
)

:: 5. Envia para o GitHub
echo.
echo [*] Enviando atualizações para o GitHub...
echo (Se for a primeira vez, uma janela de login do navegador poderá se abrir)
echo.
%GIT_CMD% push origin main
if %ERRORLEVEL% neq 0 (
    echo.
    color 0e
    echo [!] Tentando reconciliar com a nuvem...
    %GIT_CMD% push -u origin main
)

if %ERRORLEVEL% equ 0 (
    echo.
    color 0a
    echo ======================================================================
    echo   [SUCESSO] Sistema sincronizado com o GitHub com êxito!
    echo   O Streamlit Cloud já detectou as alterações e o dashboard
    echo   online está sendo atualizado automaticamente.
    echo ======================================================================
) else (
    echo.
    color 0c
    echo ======================================================================
    echo   [ATENÇÃO] Houve uma falha ao enviar para o GitHub.
    echo   Verifique se sua conta do GitHub possui permissão de gravação.
    echo ======================================================================
)

echo.
pause
