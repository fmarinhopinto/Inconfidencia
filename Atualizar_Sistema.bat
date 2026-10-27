@echo off
chcp 65001 >nul
title Atualizador Automático - Dashboard Grupo Fortpel & Inconfidência
color 0b

echo ======================================================================
echo    ATUALIZADOR AUTOMÁTICO - GRUPO FORTPEL & INCONFIDÊNCIA
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

:: 2. Verifica se o repositório Git local existe
if not exist ".git" (
    echo [*] Inicializando repositório local...
    %GIT_CMD% init -b main
    %GIT_CMD% config user.name "Filipe Marinho"
    %GIT_CMD% config user.email "filipe.marinho@inconfidencia.com"
)

:: 3. Verifica se o remote origin está configurado
%GIT_CMD% remote get-url origin >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ======================================================================
    echo   CONFIGURAÇÃO INICIAL DO REPOSITÓRIO GITHUB
    echo ======================================================================
    echo O seu projeto ainda não está conectado ao seu repositório no GitHub.
    echo.
    echo Cole abaixo o link do seu repositório GitHub:
    echo (Exemplo: https://github.com/seu-usuario/inconfidencia-fortpel.git)
    echo.
    set /p REPO_URL="Link do GitHub: "
    if "%REPO_URL%"=="" (
        color 0c
        echo [AVISO] Nenhuma URL informada. A sincronização foi cancelada.
        pause
        exit /b 1
    )
    %GIT_CMD% remote add origin %REPO_URL%
    echo [OK] Repositório remoto configurado!
    echo.
)

:: 4. Adiciona arquivos modificados e banco de dados
echo [*] Coletando arquivos alterados e banco DuckDB...
%GIT_CMD% add .

:: 5. Verifica se há algo para commitar
%GIT_CMD% diff-index --quiet HEAD --
if %ERRORLEVEL% equ 0 (
    echo [INFO] Nenhuma alteração pendente detectada no momento.
) else (
    echo [*] Registrando alterações no histórico...
    %GIT_CMD% commit -m "Atualizacao automatica - %date% %time%"
)

:: 6. Envia para o GitHub
echo.
echo [*] Sincronizando com o GitHub...
echo (Se for a primeira vez, poderá abrir uma janela no navegador para autorizar)
echo.
%GIT_CMD% push -u origin main
if %ERRORLEVEL% neq 0 (
    echo.
    color 0e
    echo [!] Tentando sincronização com a branch atual...
    %GIT_CMD% push -u origin HEAD
)

if %ERRORLEVEL% equ 0 (
    echo.
    color 0a
    echo ======================================================================
    echo   [SUCESSO] Sincronização concluída com êxito!
    echo   O Streamlit Cloud já detectou as alterações e o dashboard
    echo   online será atualizado automaticamente em instantes.
    echo ======================================================================
) else (
    echo.
    color 0c
    echo ======================================================================
    echo   [ATENÇÃO] Houve uma falha ao enviar para o GitHub.
    echo   Verifique sua conexão ou se o link do repositório está correto.
    echo ======================================================================
)

echo.
pause
