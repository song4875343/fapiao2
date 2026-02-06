@echo off
chcp 65001 >nul
echo ========================================
echo   PDF Merger - 服务器模式启动脚本
echo ========================================
echo.

REM 检查虚拟环境
if exist .venv\Scripts\activate.bat (
    echo 激活虚拟环境...
    call .venv\Scripts\activate.bat
) else (
    echo 警告: 未找到虚拟环境，使用全局 Python
)

echo.
echo 正在启动服务器...
echo.

".venv\Scripts\python.exe"  pdfm_v3.py --server --port 8000

pause
