@echo off
setlocal

set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"

if not exist "dist\CONVERSOR - VEXPER.exe" (
    echo Executavel nao encontrado em dist\CONVERSOR - VEXPER.exe
    echo Gere primeiro o executavel atualizado.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Ambiente Python nao encontrado em .venv
    pause
    exit /b 1
)

if not exist "%ISCC%" (
    echo Inno Setup 6 nao encontrado.
    echo Instale o Inno Setup e rode este arquivo novamente.
    echo Script do instalador pronto em installer.iss
    pause
    exit /b 1
)

.venv\Scripts\python.exe generate_installer_assets.py
if errorlevel 1 (
    echo Falha ao gerar as imagens do instalador.
    pause
    exit /b 1
)

.venv\Scripts\python.exe generate_build_metadata.py
if errorlevel 1 (
    echo Falha ao gerar os metadados de build.
    pause
    exit /b 1
)

"%ISCC%" "installer.iss"
if errorlevel 1 (
    echo Falha ao compilar o instalador.
    pause
    exit /b 1
)

.venv\Scripts\python.exe publish_update.py
if errorlevel 1 (
    echo Falha ao publicar o feed de atualizacao.
    pause
    exit /b 1
)

echo.
echo Instalador gerado em installer\Instalador CONVERSOR - VEXPER.exe
pause