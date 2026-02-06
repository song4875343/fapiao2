@echo off
REM 服务器部署脚本 (Windows)

echo 开始部署 PDF Merger Pro...

REM 检查 uv 是否安装
where uv >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo 错误: uv 未安装，请先安装 uv
    echo 安装命令: powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
    exit /b 1
)

REM 同步依赖
echo 正在安装依赖...
uv sync

REM 启动服务器
echo 启动服务器模式...
uv run python pdfm_v3.py --server

echo 部署完成！
