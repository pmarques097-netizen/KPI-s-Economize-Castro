$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=========================================="
Write-Host " PROJ_KPI'S_C - INICIANDO O SISTEMA"
Write-Host "=========================================="

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python não encontrado. Instale Python ou Miniconda."
    Read-Host "Pressione Enter para sair"
    exit 1
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Criando ambiente virtual..."
    python -m venv .venv
}

& ".venv\Scripts\Activate.ps1"
python -m pip install --upgrade pip | Out-Null
pip install -r requirements.txt

Start-Process "http://localhost:8501"
streamlit run app.py --server.port 8501 --server.address localhost
