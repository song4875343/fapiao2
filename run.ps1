# 如果策略禁止，先允许当前用户执行脚本（只需设一次）
# Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

# 定位到脚本所在目录
$projRoot = $PSScriptRoot
if (-not $projRoot) { $projRoot = Get-Location }

Set-Location $projRoot

# 激活 venv
$activate = Join-Path $projRoot ".venv\Scripts\Activate.ps1"
if (-not (Test-Path $activate)) {
    Write-Error "找不到 .venv\Scripts\Activate.ps1，请先创建虚拟环境！"
    exit 1
}
& $activate

# 运行主程序
python pdfm_v3.py

Write-Host "`n按任意键退出 ..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")