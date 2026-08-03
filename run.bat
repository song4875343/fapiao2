@echo off
setlocal
cd /d "%~dp0"

where uv >nul 2>nul
if errorlevel 1 (
    echo [ERROR] uv is not installed or is not available in PATH.
    echo Install uv, reopen this terminal, and run this script again:
    echo   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 ^| iex"
    pause
    exit /b 1
)

echo Preparing the Python environment and starting PDF Merger Pro...
uv run --locked python pdfm_v3.py
set "exit_code=%ERRORLEVEL%"

if not "%exit_code%"=="0" (
    echo.
    echo [ERROR] Startup failed with exit code %exit_code%.
    pause
)

exit /b %exit_code%
