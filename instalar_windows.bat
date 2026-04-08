@echo off
chcp 65001 >nul
title Recibo Agent - Instalador
echo.
echo ╔══════════════════════════════════════════╗
echo ║     RECIBO AGENT - INSTALADOR            ║
echo ║     Processador automático de recibos    ║
echo ╚══════════════════════════════════════════╝
echo.

set "AGENT_DIR=%~dp0"
set "VENV_DIR=%AGENT_DIR%\.venv"

:: ── 1. Verificar Python ──
echo [1/6] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo    Python nao encontrado. Instalando via winget...
    winget install Python.Python.3.13 --silent --accept-package-agreements --accept-source-agreements
    if %errorlevel% neq 0 (
        echo.
        echo    ❌ Nao foi possivel instalar Python automaticamente.
        echo    Baixe manualmente em: https://www.python.org/downloads/
        echo    Marque "Add to PATH" durante a instalacao.
        pause
        exit /b 1
    )
    :: Refresh PATH
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts;%PATH%"
    echo    ✅ Python instalado
) else (
    echo    ✅ Python encontrado
)

:: ── 2. Verificar Ollama ──
echo [2/6] Verificando Ollama...
ollama --version >nul 2>&1
if %errorlevel% neq 0 (
    echo    Ollama nao encontrado. Baixando instalador...
    echo    (Isso pode demorar alguns minutos)
    curl -L -o "%TEMP%\OllamaSetup.exe" "https://ollama.com/download/OllamaSetup.exe"
    if %errorlevel% neq 0 (
        echo    ❌ Falha ao baixar Ollama.
        echo    Baixe manualmente em: https://ollama.com/download
        pause
        exit /b 1
    )
    echo    Instalando Ollama...
    start /wait "%TEMP%\OllamaSetup.exe" /SILENT
    del "%TEMP%\OllamaSetup.exe"
    echo    ✅ Ollama instalado
) else (
    echo    ✅ Ollama encontrado
)

:: ── 3. Baixar modelo de visão ──
echo [3/6] Verificando modelo llama3.2-vision...
echo    (O modelo tem ~5GB - a primeira vez demora)
ollama pull llama3.2-vision
if %errorlevel% neq 0 (
    echo    ❌ Falha ao baixar modelo. Verifique sua conexao.
    pause
    exit /b 1
)
echo    ✅ Modelo pronto

:: ── 4. Criar ambiente virtual Python ──
echo [4/6] Configurando ambiente Python...
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    python -m venv "%VENV_DIR%"
)
call "%VENV_DIR%\Scripts\activate.bat"

pip install -q -r "%AGENT_DIR%\requirements.txt"
if %errorlevel% neq 0 (
    echo    ❌ Falha ao instalar dependencias Python.
    pause
    exit /b 1
)
echo    ✅ Dependencias instaladas

:: ── 5. Criar pastas OneDrive ──
echo [5/6] Verificando pastas...
set "ONEDRIVE_MAMI=%USERPROFILE%\OneDrive\MAMI"
if not exist "%ONEDRIVE_MAMI%\RECIBOS" (
    mkdir "%ONEDRIVE_MAMI%\RECIBOS"
    echo    Pasta criada: %ONEDRIVE_MAMI%\RECIBOS
)
if not exist "%ONEDRIVE_MAMI%\RECIBOS REVISADOS" (
    mkdir "%ONEDRIVE_MAMI%\RECIBOS REVISADOS"
    echo    Pasta criada: %ONEDRIVE_MAMI%\RECIBOS REVISADOS
)
echo    ✅ Pastas OK

:: ── 6. Autenticação Microsoft ──
echo [6/6] Autenticacao Microsoft (primeira vez)...
echo.
"%VENV_DIR%\Scripts\python.exe" -c "import sys; sys.path.insert(0,'%AGENT_DIR%'); from auth import get_token; get_token(); print('Autenticacao concluida!')"
if %errorlevel% neq 0 (
    echo    ⚠ Autenticacao pode ser feita depois ao rodar o programa.
)

:: ── 7. Criar atalho no Iniciar (auto-start) ──
echo.
echo Deseja iniciar o Recibo Agent automaticamente com o Windows? (S/N)
set /p AUTOSTART="> "
if /i "%AUTOSTART%"=="S" (
    set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
    echo @echo off > "%STARTUP%\ReciboAgent.bat"
    echo cd /d "%AGENT_DIR%" >> "%STARTUP%\ReciboAgent.bat"
    echo call .venv\Scripts\activate.bat >> "%STARTUP%\ReciboAgent.bat"
    echo start /min python run.py >> "%STARTUP%\ReciboAgent.bat"
    echo    ✅ Auto-start configurado
)

echo.
echo ╔══════════════════════════════════════════╗
echo ║    INSTALACAO CONCLUIDA!                 ║
echo ║                                          ║
echo ║    Para rodar manualmente:               ║
echo ║    1. Abra esta pasta no terminal        ║
echo ║    2. .venv\Scripts\activate             ║
echo ║    3. python run.py                      ║
echo ║                                          ║
echo ║    Recibos na pasta:                     ║
echo ║    OneDrive\MAMI\RECIBOS                 ║
echo ╚══════════════════════════════════════════╝
echo.
pause
