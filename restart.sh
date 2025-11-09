#!/bin/bash
# Excel工具箱 - 服务重启脚本

cd "$(dirname "$0")"

echo "🔄 正在停止服务..."
pkill -f "python.*app.py" 2>/dev/null && echo "✓ 已停止旧服务" || echo "ℹ️ 没有运行中的服务"

echo ""
echo "🚀 正在启动服务..."

# 激活虚拟环境并启动
if [ -d "venv" ]; then
    source venv/bin/activate
    echo "✓ 已激活虚拟环境"
else
    echo "⚠️ 未找到虚拟环境，使用系统Python"
fi

echo ""
echo "="*60
echo "Excel价格批量更新系统 - 启动中"
echo "="*60
echo ""

# 启动服务
python app.py

