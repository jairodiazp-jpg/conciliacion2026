# start_all_port3000.ps1
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
Set-Location $here

# Preparar virtualenv
$venvPath = Join-Path $here '.venv'
$pythonExe = Join-Path $venvPath 'Scripts\python.exe'
if (-not (Test-Path $pythonExe)) {
    python -m venv .venv
}

# Instalar dependencias
& $pythonExe -m pip install --upgrade pip
if (Test-Path (Join-Path $here 'requirements.txt')) {
    & $pythonExe -m pip install -r requirements.txt
} else {
    & $pythonExe -m pip install openpyxl uvicorn fastapi python-multipart anyio
}

# Iniciar backend
$backendPort = 3000
$uvicornArgs = "-m uvicorn main:app --host 127.0.0.1 --port $backendPort"
$backendProc = Start-Process -FilePath $pythonExe -ArgumentList $uvicornArgs -WorkingDirectory $here -PassThru -WindowStyle Hidden

# Iniciar frontend
$frontendDir = Join-Path $here '..\frontend' | Resolve-Path -ErrorAction SilentlyContinue
if ($null -ne $frontendDir) {
    $frontendDir = $frontendDir.Path
    if (Test-Path (Join-Path $frontendDir 'package.json')) {
        $env:VITE_API_URL = "http://127.0.0.1:$backendPort"
        Start-Process -FilePath 'npm.cmd' -ArgumentList 'run', 'dev' -WorkingDirectory $frontendDir -WindowStyle Hidden
    }
}

# Esperar backend
$healthUrl = "http://127.0.0.1:$backendPort/health"
$maxRetries = 30
$retry = 0
while ($retry -lt $maxRetries) {
    try {
        Invoke-RestMethod -Uri $healthUrl -Method Get -ErrorAction Stop
        break
    } catch {
        Start-Sleep -Seconds 2
        $retry += 1
    }
}