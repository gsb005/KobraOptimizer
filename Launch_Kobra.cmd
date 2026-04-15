@echo off
setlocal

set "APPDIR=%~dp0"
cd /d "%APPDIR%"

echo Starting KobraOptimizer...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath ""%APPDIR%"" -Recurse -File -ErrorAction SilentlyContinue ^| Unblock-File -ErrorAction SilentlyContinue" >nul 2>nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%APPDIR%Main.ps1"

endlocal
