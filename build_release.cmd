@echo off
cd /d "%~dp0"
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File ".\build_release.ps1"
if errorlevel 1 (
    echo.
    echo Build failed.
)
pause
