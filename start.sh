#!/bin/bash
# Excel工具箱 - 服务启动脚本（后台运行）

cd "$(dirname "$0")"

echo "🔄 检查服务状态..."
if pgrep -f "python.*app.py" > /dev/null; then
    echo "⚠️ 服务已在运行中"
    echo "如需重启，请先运行: ./stop.sh"
    exit 1
fi

echo ""
echo "🚀 正在启动服务..."

# 激活虚拟环境
if [ -d "venv" ]; then
    source venv/bin/activate
    echo "✓ 已激活虚拟环境"
fi

# 后台启动服务
nohup python app.py > server.log 2>&1 &
PID=$!

sleep 2

if ps -p $PID > /dev/null; then
    echo ""
    echo "✅ 服务启动成功！"
    echo "📝 进程ID: $PID"
    echo "🌐 访问地址: http://localhost:8800"
    echo "📄 日志文件: server.log"
    echo ""
    echo "停止服务: ./stop.sh"
    echo "查看日志: tail -f server.log"
else
    echo ""
    echo "❌ 服务启动失败"
    echo "查看错误: cat server.log"
fi

