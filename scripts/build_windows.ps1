param([switch]$SinDependencias)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
$python = if (Test-Path '.venv\Scripts\python.exe') { '.venv\Scripts\python.exe' } else { 'python' }

if (-not $SinDependencias) {
    & $python -m pip install -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw 'No se pudieron instalar las dependencias de build' }
}

$poppler = Get-Command pdftoppm.exe -ErrorAction SilentlyContinue
if (-not $poppler) {
    $packages = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    if (Test-Path $packages) {
        $poppler = Get-ChildItem $packages -Recurse -Filter pdftoppm.exe -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }
}
if ($poppler) {
    $popplerPath = if ($poppler.Source) { $poppler.Source } else { $poppler.FullName }
    $env:POPPLER_BIN_DIR = Split-Path -Parent $popplerPath
} else {
    throw 'Falta Poppler para incluir el lector de PDF escaneados en el paquete'
}

& $python -m PyInstaller --noconfirm --clean Regularizaciones.spec
if ($LASTEXITCODE -ne 0) { throw 'Falló el build de PyInstaller' }
$exe = Join-Path $root 'dist\Regularizaciones\Regularizaciones.exe'
if (-not (Test-Path $exe)) { throw "No se generó $exe" }

$forbidden = @('gestion.db', 'ui_prefs.json', 'proveedores_despacho.json', 'letter_identities.json')
foreach ($name in $forbidden) {
    if (Get-ChildItem (Join-Path $root 'dist\Regularizaciones') -Recurse -File -Filter $name) {
        throw "El paquete contiene un archivo privado: $name"
    }
}
Write-Host "Distribución lista: $exe"
