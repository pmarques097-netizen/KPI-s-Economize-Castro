param([switch]$Stop)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectDir

$LogDir = Join-Path $ProjectDir "data\logs"
$VenvDir = Join-Path $ProjectDir ".venv"
$PythonExe = Join-Path $VenvDir "Scripts\python.exe"
$AppFile = Join-Path $ProjectDir "app.py"
$ReqFile = Join-Path $ProjectDir "requirements.txt"
$ReqHashFile = Join-Path $LogDir "requirements.sha256"
$PidFile = Join-Path $LogDir "streamlit.pid"
$LauncherLog = Join-Path $LogDir "launcher.log"
$StdoutLog = Join-Path $LogDir "streamlit_saida.log"
$StderrLog = Join-Path $LogDir "streamlit_erro.log"
$Url = "http://localhost:8501"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Log([string]$Message) {
    Add-Content -Path $LauncherLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Port-Open {
    try {
        return $null -ne (Get-NetTCPConnection -LocalPort 8501 -State Listen -ErrorAction Stop)
    } catch { return $false }
}

function Find-Python {
    foreach ($candidate in @("py -3.12", "py -3.11", "py -3", "python")) {
        try {
            $parts = $candidate.Split(" ")
            $cmd = $parts[0]
            $args = @()
            if ($parts.Count -gt 1) { $args = $parts[1..($parts.Count - 1)] }
            & $cmd @args --version *> $null
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch {}
    }
    return $null
}

function Stop-Server {
    if (Test-Path $PidFile) {
        try {
            $serverPid = [int](Get-Content $PidFile -Raw)
            if (Get-Process -Id $serverPid -ErrorAction SilentlyContinue) {
                Stop-Process -Id $serverPid -Force
                Log "Servidor encerrado. PID=$serverPid"
            }
        } catch { Log "Falha ao encerrar PID: $($_.Exception.Message)" }
        Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
    }
}

if ($Stop) { Stop-Server; exit 0 }

if (-not (Test-Path $AppFile)) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("O arquivo app.py não foi encontrado.", "Rede Economize - KPI Comercial", "OK", "Error") | Out-Null
    exit 1
}

if (Port-Open) { Start-Process $Url; exit 0 }

if (-not (Test-Path $PythonExe)) {
    $pythonCommand = Find-Python
    if ([string]::IsNullOrWhiteSpace($pythonCommand)) {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Python 3.11 ou 3.12 não foi encontrado.", "Rede Economize - KPI Comercial", "OK", "Error") | Out-Null
        exit 1
    }
    Log "Criando ambiente virtual com $pythonCommand"
    $parts = $pythonCommand.Split(" ")
    $cmd = $parts[0]
    $args = @()
    if ($parts.Count -gt 1) { $args = $parts[1..($parts.Count - 1)] }
    & $cmd @args -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) { throw "Falha ao criar o ambiente virtual." }
}

$mustInstall = $false
if (-not (Test-Path $ReqHashFile)) { $mustInstall = $true }
if (Test-Path $ReqFile) {
    $currentHash = (Get-FileHash $ReqFile -Algorithm SHA256).Hash
    $savedHash = if (Test-Path $ReqHashFile) { (Get-Content $ReqHashFile -Raw).Trim() } else { "" }
    if ($currentHash -ne $savedHash) { $mustInstall = $true }
}

# Verifica os módulos essenciais mesmo se o ambiente virtual já existir.
& $PythonExe -c "import streamlit, pandas, plotly, sqlalchemy, psycopg2, openpyxl, reportlab" *> $null
if ($LASTEXITCODE -ne 0) { $mustInstall = $true }

if ($mustInstall) {
    Log "Instalando/atualizando dependências."
    & $PythonExe -m pip install --upgrade pip *>> $LauncherLog
    & $PythonExe -m pip install -r $ReqFile *>> $LauncherLog
    if ($LASTEXITCODE -ne 0) {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Falha na instalação das dependências. Consulte data\logs\launcher.log.", "Rede Economize - KPI Comercial", "OK", "Error") | Out-Null
        exit 1
    }
    $currentHash = (Get-FileHash $ReqFile -Algorithm SHA256).Hash
    Set-Content $ReqHashFile $currentHash
}

$args = @(
    "-m", "streamlit", "run", $AppFile,
    "--server.address", "127.0.0.1",
    "--server.port", "8501",
    "--server.headless", "true",
    "--browser.gatherUsageStats", "false"
)

$process = Start-Process -FilePath $PythonExe -ArgumentList $args -WorkingDirectory $ProjectDir -WindowStyle Hidden -RedirectStandardOutput $StdoutLog -RedirectStandardError $StderrLog -PassThru
Set-Content $PidFile $process.Id
Log "Streamlit iniciado com app.py. PID=$($process.Id)"

$started = $false
for ($i = 0; $i -lt 90; $i++) {
    Start-Sleep -Seconds 1
    if ($process.HasExited) {
        $err = if (Test-Path $StderrLog) { Get-Content $StderrLog -Raw } else { "Sem detalhes." }
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Falha ao iniciar o sistema.`n`n$err", "Rede Economize - KPI Comercial", "OK", "Error") | Out-Null
        Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
        exit 1
    }
    if (Port-Open) { $started = $true; break }
}

if (-not $started) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("Tempo limite de inicialização. Consulte data\logs\streamlit_erro.log.", "Rede Economize - KPI Comercial", "OK", "Warning") | Out-Null
    exit 1
}

Start-Process $Url
