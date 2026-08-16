# start_services.ps1
# Arranca backend (uvicorn) y frontend (Vite) en procesos separados (detached).
# Ejecutar desde PowerShell en la carpeta backend:
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   .\start_services.ps1

$ErrorActionPreference = 'Stop'
$here = Split-Path -LiteralPath $MyInvocation.MyCommand.Definition -Parent
Set-Location $here
Write-Host "Working dir: $here"

$venvPath = Join-Path $here '.venv'
$pythonExe = Join-Path $venvPath 'Scripts\python.exe'

if (-not (Test-Path $pythonExe)) {
    Write-Host "Virtualenv no encontrado en $venvPath. Ejecuta run_adq_tests.ps1 primero para crear el entorno." -ForegroundColor Red
    exit 1
}

# Levantar backend con uvicorn (detached)
$uvicornArgs = '-m uvicorn main:app --host 127.0.0.1 --port 8000'
Write-Host "Iniciando backend: $pythonExe $uvicornArgs"
Start-Process -FilePath $pythonExe -ArgumentList $uvicornArgs -WorkingDirectory $here -NoNewWindow -WindowStyle Hidden
Write-Host "Backend (uvicorn) iniciado en background en http://127.0.0.1:8000"

# Intentar arrancar frontend (si npm está disponible)
$frontendDir = Join-Path $here '..\frontend' | Resolve-Path -ErrorAction SilentlyContinue
if ($null -ne $frontendDir) {
    $frontendDir = $frontendDir.Path
    if (Test-Path (Join-Path $frontendDir 'package.json')) {
        Write-Host "Intentando iniciar frontend (Vite) en $frontendDir"
        # Usar npm si está en PATH
        try {
            Start-Process -FilePath 'npm' -ArgumentList 'run dev' -WorkingDirectory $frontendDir -NoNewWindow -WindowStyle Hidden
            Write-Host "Frontend (npm run dev) iniciado en background. Abre http://localhost:5173 (o la URL que muestre Vite)."
        } catch {
            Write-Host "No se pudo iniciar frontend con npm desde aquí. Asegúrate de tener Node.js y npm instalados y ejecuta 'npm run dev' en la carpeta frontend." -ForegroundColor Yellow
        }
    } else {
        Write-Host "package.json no encontrado en frontend. No se arrancó frontend." -ForegroundColor Yellow
    }
} else {
    Write-Host "Carpeta frontend no encontrada al buscar ../frontend" -ForegroundColor Yellow
}

Write-Host "\nServicios iniciados (si no hubo errores). Comprueba procesos y logs en esta terminal o usa el navegador para verificar las URLs."