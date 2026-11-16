@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv" (
    echo Criando ambiente virtual Python...
    python -m venv .venv
)

call ".venv\Scripts\activate.bat"
pip install -q -r requirements.txt
echo Iniciando SCEN - Gerador de Trabalhos...
python app_scen.py
pause
