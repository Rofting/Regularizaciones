<#
  Instalador de Regularizaciones para Windows 10/11.

  Instala lo que falte (Python 3.12, LibreOffice y Poppler) con winget,
  prepara el entorno de Python del programa en la carpeta .venv, comprueba
  que todo funciona y crea los accesos directos. Se puede ejecutar las veces
  que haga falta: lo que ya está instalado no se vuelve a instalar.

  Uso normal: doble clic en INSTALAR.bat (en la carpeta del programa).
  Opciones:   -Desatendido          sin preguntas ni pausas (pruebas automáticas)
              -SinAccesosDirectos   no crea accesos en el escritorio ni en Inicio
#>
param(
    [switch]$Desatendido,
    [switch]$SinAccesosDirectos
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Raiz = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$Registro = Join-Path $Raiz 'instalacion.log'
try { Start-Transcript -Path $Registro -Append | Out-Null } catch { }

$WingetYaInstalado = @(-1978335189, -1978335135)  # sin actualización / ya instalado

function Titulo([string]$texto) {
    Write-Host ''
    Write-Host $texto -ForegroundColor Cyan
}
function Bien([string]$texto)  { Write-Host "   OK     $texto" -ForegroundColor Green }
function Aviso([string]$texto) { Write-Host "   AVISO  $texto" -ForegroundColor Yellow }

function Terminar([int]$codigo, [string]$mensaje) {
    Write-Host ''
    if ($codigo -eq 0) {
        Write-Host $mensaje -ForegroundColor Green
    } else {
        Write-Host $mensaje -ForegroundColor Red
        Write-Host "Detalle completo en: $Registro" -ForegroundColor Red
    }
    try { Stop-Transcript | Out-Null } catch { }
    if (-not $Desatendido) {
        Write-Host ''
        Read-Host 'Pulsa Intro para cerrar esta ventana' | Out-Null
    }
    exit $codigo
}

function Hay-Winget { return [bool](Get-Command winget -ErrorAction SilentlyContinue) }

function Instalar-ConWinget([string]$id, [string]$nombre, [string]$ambito) {
    if (-not (Hay-Winget)) {
        Terminar 1 ("Falta winget («Instalador de aplicación» de Microsoft). Ábrelo desde Microsoft Store, " +
                    "actualízalo y vuelve a ejecutar INSTALAR.bat.")
    }
    Write-Host "   Instalando $nombre (puede tardar unos minutos)..."
    $argumentos = @('install', '--exact', '--id', $id, '--silent',
                    '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
    if ($ambito) { $argumentos += @('--scope', $ambito) }
    & winget @argumentos | Out-Host
    $codigo = $LASTEXITCODE
    if ($codigo -ne 0 -and ($WingetYaInstalado -notcontains $codigo)) {
        return $false
    }
    return $true
}

function Probar-Python([string]$ruta) {
    if (-not $ruta -or -not (Test-Path $ruta)) { return $false }
    if ($ruta -like '*WindowsApps*') { return $false }   # acceso de Microsoft Store, no es Python
    & $ruta -c "import sys, tkinter; sys.exit(0 if sys.version_info[:2] == (3, 12) else 1)" 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Buscar-Python {
    $candidatos = New-Object System.Collections.Generic.List[string]
    $lanzador = Get-Command py -ErrorAction SilentlyContinue
    if ($lanzador) {
        $ruta = & $lanzador.Source -3.12 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $ruta) { $candidatos.Add(($ruta | Select-Object -First 1).Trim()) }
    }
    $candidatos.Add((Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'))
    $candidatos.Add((Join-Path $env:ProgramFiles 'Python312\python.exe'))
    foreach ($candidato in $candidatos) {
        if (Probar-Python $candidato) { return $candidato }
    }
    return $null
}

function Buscar-LibreOffice {
    foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $base) { continue }
        $ruta = Join-Path $base 'LibreOffice\program\soffice.com'
        if (Test-Path $ruta) { return $ruta }
    }
    return $null
}

function Buscar-Poppler {
    if (Get-Command pdftoppm -ErrorAction SilentlyContinue) { return 'PATH' }
    $paquetes = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    if (Test-Path $paquetes) {
        $encontrado = Get-ChildItem -Path $paquetes -Directory -Filter '*Poppler*' -ErrorAction SilentlyContinue |
            ForEach-Object { Get-ChildItem -Path $_.FullName -Recurse -Filter 'pdftoppm.exe' -ErrorAction SilentlyContinue } |
            Select-Object -First 1
        if ($encontrado) { return $encontrado.DirectoryName }
    }
    return $null
}

Write-Host '============================================================' -ForegroundColor Cyan
Write-Host '  Instalación de Regularizaciones' -ForegroundColor Cyan
Write-Host "  Carpeta del programa: $Raiz" -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan

if (-not (Test-Path (Join-Path $Raiz 'core\app.py'))) {
    Terminar 1 'No se encuentra core\app.py. Descomprime la carpeta completa y ejecuta INSTALAR.bat desde ella.'
}
if ($Raiz -match 'OneDrive') {
    Aviso 'El programa está dentro de OneDrive. Es mejor moverlo a C:\Regularizaciones: la sincronización puede bloquear la base de datos.'
}
# Los archivos descargados de Internet llevan una marca que hace que Windows pregunte al abrirlos.
Get-ChildItem -Path $Raiz -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notlike '*\.venv\*' } |
    Unblock-File -ErrorAction SilentlyContinue

# 1. Python --------------------------------------------------------------
Titulo '[1/5] Python 3.12'
$Python = Buscar-Python
if (-not $Python) {
    if (-not (Instalar-ConWinget 'Python.Python.3.12' 'Python 3.12' 'user')) {
        Terminar 1 'No se pudo instalar Python 3.12. Instálalo desde https://www.python.org/downloads/ (versión 3.12) y vuelve a ejecutar INSTALAR.bat.'
    }
    $Python = Buscar-Python
}
if (-not $Python) {
    Terminar 1 'Python 3.12 se instaló pero no se encuentra. Cierra esta ventana y vuelve a ejecutar INSTALAR.bat.'
}
Bien $Python

# 2. LibreOffice ---------------------------------------------------------
Titulo '[2/5] LibreOffice (necesario para generar el Excel oficial)'
$LibreOffice = Buscar-LibreOffice
if (-not $LibreOffice) {
    Write-Host '   Windows puede pedir permiso de administrador: acepta para continuar.'
    if (-not (Instalar-ConWinget 'TheDocumentFoundation.LibreOffice' 'LibreOffice' '')) {
        Terminar 1 'No se pudo instalar LibreOffice. Instálalo desde https://es.libreoffice.org/descarga/ y vuelve a ejecutar INSTALAR.bat.'
    }
    $LibreOffice = Buscar-LibreOffice
}
if ($LibreOffice) { Bien $LibreOffice } else {
    Terminar 1 'LibreOffice no aparece tras instalarlo. Reinicia el ordenador y vuelve a ejecutar INSTALAR.bat.'
}

# 3. Poppler ---------------------------------------------------------------
Titulo '[3/5] Poppler (lectura de PDF escaneados)'
$Poppler = Buscar-Poppler
if (-not $Poppler) {
    if (Instalar-ConWinget 'oschwartz10612.Poppler' 'Poppler' 'user') { $Poppler = Buscar-Poppler }
}
if ($Poppler) { Bien $Poppler } else {
    Aviso 'Poppler no se pudo instalar. El programa funciona, pero no podrá leer PDF escaneados (sin texto).'
}

# 4. Entorno de Python del programa ----------------------------------------
Titulo '[4/5] Librerías del programa (carpeta .venv)'
$Venv = Join-Path $Raiz '.venv'
$VenvPython = Join-Path $Venv 'Scripts\python.exe'
if ((Test-Path $VenvPython) -and -not (Probar-Python $VenvPython)) {
    Write-Host '   El entorno anterior no es válido; se vuelve a crear.'
    Remove-Item -Recurse -Force $Venv -ErrorAction SilentlyContinue
}
if (-not (Test-Path $VenvPython)) {
    & $Python -m venv $Venv | Out-Host
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython)) {
        Terminar 1 'No se pudo crear el entorno .venv. Comprueba que la carpeta no es de solo lectura.'
    }
}
& $VenvPython -m pip install --upgrade pip --disable-pip-version-check --quiet | Out-Host
& $VenvPython -m pip install -r (Join-Path $Raiz 'requirements.txt') --disable-pip-version-check | Out-Host
if ($LASTEXITCODE -ne 0) {
    Terminar 1 'No se pudieron instalar las librerías. Comprueba la conexión a Internet y vuelve a ejecutar INSTALAR.bat.'
}
Bien 'Librerías instaladas'

# 5. Comprobación y accesos directos ---------------------------------------
Titulo '[5/5] Comprobación final'
$env:PYTHONUTF8 = '1'
& $VenvPython (Join-Path $Raiz 'scripts\comprobar_instalacion.py') | Out-Host
$Comprobacion = $LASTEXITCODE

if (-not $SinAccesosDirectos) {
    $Pythonw = Join-Path $Venv 'Scripts\pythonw.exe'
    $Shell = New-Object -ComObject WScript.Shell
    $Destinos = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))
    foreach ($carpeta in $Destinos) {
        if (-not $carpeta -or -not (Test-Path $carpeta)) { continue }
        $Acceso = $Shell.CreateShortcut((Join-Path $carpeta 'Regularizaciones.lnk'))
        $Acceso.TargetPath = $Pythonw
        $Acceso.Arguments = '"' + (Join-Path $Raiz 'core\app.py') + '"'
        $Acceso.WorkingDirectory = $Raiz
        $Acceso.Description = 'Regularizaciones de agua caliente y calefacción'
        $Acceso.IconLocation = "$env:SystemRoot\System32\imageres.dll,109"
        $Acceso.Save()
    }
    Bien 'Accesos directos «Regularizaciones» en el escritorio y en el menú Inicio'
}

if ($Comprobacion -ne 0) {
    Terminar 1 'La instalación terminó, pero falta algún componente imprescindible (ver arriba).'
}
if (-not $Desatendido) {
    $Respuesta = Read-Host '¿Abrir Regularizaciones ahora? (S/N)'
    if ($Respuesta -match '^[sS]') {
        Start-Process -FilePath (Join-Path $Venv 'Scripts\pythonw.exe') -ArgumentList ('"' + (Join-Path $Raiz 'core\app.py') + '"') -WorkingDirectory $Raiz
    }
}
Terminar 0 'Instalación completada.'
