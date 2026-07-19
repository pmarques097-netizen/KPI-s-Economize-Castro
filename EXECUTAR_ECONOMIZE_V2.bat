@echo off
setlocal
cd /d "%~dp0"
title REDE ECONOMIZE - KPI COMERCIAL

echo ================================================
echo  REDE ECONOMIZE - KPI COMERCIAL
echo ================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
)

call ".venv\Scripts\activate.bat"
pip install -r requirements.txt

echo.
echo Abrindo nova versao Rede Economize...
start "" http://localhost:8502
streamlit run app.py --server.port 8502 --server.address localhost

pause
