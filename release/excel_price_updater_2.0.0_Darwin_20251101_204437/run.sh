#!/bin/bash

# Excel价格批量更新工具启动脚本

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "========================================"
echo " Excel价格批量更新工具 v2.0.0"
echo "========================================"
echo ""
echo "正在启动程序..."
echo "程序启动后会自动打开浏览器"
echo "访问地址: http://localhost:8800"
echo ""
echo "按 Ctrl+C 可以停止程序"
echo "========================================"
echo ""

./excel_price_updater

if [ $? -ne 0 ]; then
    echo ""
    echo "❌ 程序运行出错！"
    echo ""
    read -p "按 Enter 键继续..."
fi
