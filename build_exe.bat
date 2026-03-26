@echo off
setlocal

if not exist .venv (
    c:/python314/python.exe -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -c "from pathlib import Path; from PIL import Image; logo = Path('logo.png'); icon = Path('logo.ico'); image = Image.open(logo).convert('RGBA'); image.save(icon, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])"
if exist "dist\CONVERSOR - VEXPER.exe" del /f /q "dist\CONVERSOR - VEXPER.exe"
if exist "dist\CONVERSOR - VEXPER" rmdir /s /q "dist\CONVERSOR - VEXPER"
pyinstaller --noconfirm --clean --windowed --onedir --name "CONVERSOR - VEXPER" --icon "logo.ico" --add-data "logo.png;." --add-data "PLANILHA MODELO;PLANILHA MODELO" app.py

echo.
echo Aplicativo gerado em dist\CONVERSOR - VEXPER\CONVERSOR - VEXPER.exe
pause