@echo off
setlocal

cd /d "%~dp0"
title Invoice Reimbursement

where uv >nul 2>&1
if errorlevel 1 (
    echo [ERROR] uv was not found in PATH.
    echo Install it from https://docs.astral.sh/uv/ and try again.
    pause
    exit /b 1
)

set "UV_CACHE_DIR=%TEMP%\baoxiao-uv-cache"

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Creating .venv...
    uv venv --no-config ".venv"
    if errorlevel 1 goto :failed
)

echo [INFO] Installing project dependencies...
uv pip install --no-config --python ".venv\Scripts\python.exe" -r "requirements.txt"
if errorlevel 1 goto :failed

echo [INFO] Starting Streamlit...
call ".venv\Scripts\activate.bat"
uv run --no-config --no-project --active python -m streamlit run "app.py"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] Streamlit exited with code %EXIT_CODE%.
    pause
)
exit /b %EXIT_CODE%

:failed
echo.
echo [ERROR] Environment setup failed.
pause
exit /b 1
