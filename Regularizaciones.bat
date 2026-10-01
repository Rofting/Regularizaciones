@echo off
rem Abre Regularizaciones. Con "Regularizaciones.bat diagnostico" se ve la consola con los errores.
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
    echo Regularizaciones aun no esta instalado en este equipo.
    echo Haz doble clic en INSTALAR.bat y vuelve a intentarlo.
    pause
    exit /b 1
)
set PYTHONUTF8=1
if /i "%~1"=="diagnostico" (
    ".venv\Scripts\python.exe" "%~dp0core\app.py"
    echo.
    echo El programa se ha cerrado. Si hay un error arriba, haz una captura de esta ventana.
    pause
    exit /b
)
start "" ".venv\Scripts\pythonw.exe" "%~dp0core\app.py"
