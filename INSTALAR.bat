@echo off
rem Instala o repara Regularizaciones en este equipo (Windows 10/11).
rem Se puede ejecutar todas las veces que haga falta.
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\windows\instalar.ps1" %*
exit /b %errorlevel%
