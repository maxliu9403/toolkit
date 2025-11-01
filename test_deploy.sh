#!/bin/bash

# Excel价格批量更新工具 - 部署测试脚本

echo "========================================"
echo "Excel价格批量更新工具 - 部署测试"
echo "========================================"
echo ""

# 1. 检查虚拟环境
echo "📋 1. 检查虚拟环境..."
if [ -d "venv" ]; then
    echo "   ✅ 虚拟环境存在"
    source venv/bin/activate
else
    echo "   ❌ 虚拟环境不存在，正在创建..."
    python3 -m venv venv
    source venv/bin/activate
fi
echo ""

# 2. 检查依赖
echo "📋 2. 检查依赖..."
python -c "import pandas, openpyxl" 2>/dev/null
if [ $? -eq 0 ]; then
    echo "   ✅ 核心依赖已安装"
else
    echo "   ⚠️  依赖缺失，正在安装..."
    pip install -r requirements.txt
fi
echo ""

# 3. 检查 PyInstaller
echo "📋 3. 检查 PyInstaller..."
python -c "import PyInstaller" 2>/dev/null
if [ $? -eq 0 ]; then
    echo "   ✅ PyInstaller 已安装"
else
    echo "   ⚠️  PyInstaller 未安装，正在安装..."
    pip install pyinstaller
fi
echo ""

# 4. 检查必需文件
echo "📋 4. 检查必需文件..."
required_files=("app.py" "main.py" "index.html" "config_editor.html" "config.json")
all_files_exist=true

for file in "${required_files[@]}"; do
    if [ -f "$file" ]; then
        echo "   ✅ $file"
    else
        echo "   ❌ $file 不存在"
        all_files_exist=false
    fi
done
echo ""

if [ "$all_files_exist" = false ]; then
    echo "❌ 缺少必需文件，无法继续"
    exit 1
fi

# 5. 运行部署脚本（测试模式）
echo "📋 5. 测试部署脚本..."
python deploy.py --help > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 部署脚本正常"
else
    echo "   ❌ 部署脚本有错误"
    exit 1
fi
echo ""

# 6. 显示部署选项
echo "========================================"
echo "✅ 所有检查通过！"
echo "========================================"
echo ""
echo "📦 可以开始部署："
echo ""
echo "选项 1: 单文件模式（推荐）"
echo "   python deploy.py"
echo ""
echo "选项 2: 保留临时文件（调试）"
echo "   python deploy.py --keep-temp"
echo ""
echo "选项 3: 目录模式"
echo "   python deploy.py --onedir"
echo ""
echo "========================================"

