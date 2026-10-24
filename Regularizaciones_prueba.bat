@echo off
rem Abre Regularizaciones sobre una instalacion de PRUEBA con comunidades inventadas.
rem No abre ni modifica la base de datos real del despacho.
rem   Regularizaciones_prueba.bat          abre la prueba (la crea la primera vez)
rem   Regularizaciones_prueba.bat nueva    la borra y la vuelve a crear desde cero
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Regularizaciones aun no esta instalado en este equipo.
    echo Haz doble clic en INSTALAR.bat y vuelve a intentarlo.
    pause
    exit /b 1
)
set PYTHONUTF8=1
set "REGULARIZACIONES_PRUEBA=%LOCALAPPDATA%\Regularizaciones-prueba"
if /i "%~1"=="nueva" (
    ".venv\Scripts\python.exe" "%~dp0scripts\crear_comunidades_prueba.py" --destino "%REGULARIZACIONES_PRUEBA%" --reemplazar || (pause & exit /b 1)
) else if not exist "%REGULARIZACIONES_PRUEBA%\data\gestion.db" (
    echo Creando las comunidades de prueba, tarda alrededor de un minuto...
    ".venv\Scripts\python.exe" "%~dp0scripts\crear_comunidades_prueba.py" --destino "%REGULARIZACIONES_PRUEBA%" || (pause & exit /b 1)
)
set "REGULARIZACIONES_HOME=%REGULARIZACIONES_PRUEBA%"
start "" ".venv\Scripts\pythonw.exe" "%~dp0core\app.py"
