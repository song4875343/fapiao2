#!/bin/bash
# 服务器部署脚本 (Linux/Mac)

echo "开始部署 PDF Merger Pro..."

# 检查 uv 是否安装
if ! command -v uv &> /dev/null
then
    echo "错误: uv 未安装，请先安装 uv"
    echo "安装命令: curl -LsSf https://astral.sh/uv/install.sh | sh"
    exit 1
fi

# 同步依赖
echo "正在安装依赖..."
uv sync

# 启动服务器
echo "启动服务器模式..."
uv run python pdfm_v3.py --server

echo "部署完成！"
