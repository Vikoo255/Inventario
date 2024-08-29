@echo off
setlocal

REM Obtener la letra de la unidad donde está el script
set scriptPath=%~dp0

REM Ejecutar el script de PowerShell como administrador, ignorando la política de ejecución
PowerShell -Command "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%scriptPath%a.ps1\"' -Verb RunAs"

endlocal
