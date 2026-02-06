# PDF Merger - 服务器模式启动脚本 (PowerShell)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PDF Merger - 服务器模式启动脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查虚拟环境
if (Test-Path ".venv\Scripts\Activate.ps1") {
    Write-Host "激活虚拟环境..." -ForegroundColor Green
    & .venv\Scripts\Activate.ps1
} else {
    Write-Host "警告: 未找到虚拟环境，使用全局 Python" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "正在启动服务器..." -ForegroundColor Green
Write-Host ""

".venv\Scripts\python.exe" pdfm_v3.py --server --port 8000

Read-Host "按 Enter 键退出"
