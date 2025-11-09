#!/bin/bash
# Excel工具箱 - 服务停止脚本

cd "$(dirname "$0")"

echo "🛑 正在停止服务..."

if pkill -f "python.*app.py"; then
    echo "✅ 服务已停止"
else
    echo "ℹ️ 没有运行中的服务"
fi

echo ""
echo "要启动服务，请运行: ./start.sh"

