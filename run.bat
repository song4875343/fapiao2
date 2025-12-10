@echo off
cd /d "%~dp0"

:: 直接使用虚拟环境目录下的 python.exe 执行脚本
".venv\Scripts\python.exe" pdfm_v2.py

pause