@echo off
setlocal

cd /d "%~dp0"

echo ==========================================
echo  Setup do projeto verificacao_nf
echo ==========================================

set "PYTHON_CMD="

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo [ERRO] Python nao encontrado no PATH.
    echo Instale o Python 3 e rode este script novamente.
    exit /b 1
)

if not exist "requirements.txt" (
    echo [ERRO] Arquivo requirements.txt nao encontrado na pasta do projeto.
    exit /b 1
)

echo [1/6] Criando ambiente virtual...
if not exist ".venv\Scripts\python.exe" (
    call %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERRO] Falha ao criar o ambiente virtual.
        exit /b 1
    )
) else (
    echo Ambiente virtual .venv ja existe.
)

echo [2/6] Ativando ambiente virtual...
call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERRO] Falha ao ativar o ambiente virtual.
    exit /b 1
)

echo [3/6] Atualizando pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [ERRO] Falha ao atualizar o pip.
    exit /b 1
)

echo [4/6] Instalando dependencias do requirements.txt...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERRO] Falha ao instalar as dependencias.
    exit /b 1
)

echo [5/6] Preparando pastas locais...
if not exist "output" mkdir output
if not exist "output\screenshots" mkdir output\screenshots
if not exist "downloads" mkdir downloads

echo [6/6] Instalando navegador do Playwright...
python -m playwright install chromium
if errorlevel 1 (
    echo [ERRO] Falha ao instalar o Chromium do Playwright.
    exit /b 1
)

echo.
echo Setup concluido com sucesso.
if not exist ".env" (
    echo [AVISO] Arquivo .env nao encontrado.
    echo Crie o .env com as credenciais e configuracoes antes de executar o robo.
)
echo Para ativar o ambiente depois, use:
echo .venv\Scripts\activate
echo Para executar a automacao, use:
echo python main.py

endlocal
