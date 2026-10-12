@echo off
setlocal

cd /d "%~dp0" || exit /b 1

where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo PowerShell не найден. Установите Windows PowerShell и повторите попытку.
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1" %*
exit /b %ERRORLEVEL%
