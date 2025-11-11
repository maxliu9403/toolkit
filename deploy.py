#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel价格批量更新工具 部署脚本 - 使用 PyInstaller 打包项目
支持 Web 界面和配置文件的外部配置
优化版：自动安装依赖、CI/CD 友好、日志统一化、自动清理

使用方法:
    python deploy.py                    # 默认：单文件模式，自动清理
    python deploy.py --keep-temp        # 保留临时文件（build、dist、*.spec）
    python deploy.py --onedir           # 使用目录模式（而非单文件）
    python deploy.py --help             # 显示帮助信息
    python deploy.py --version          # 显示版本信息

特性:
    ✅ 自动检测并安装缺失的依赖
    ✅ 打包 Web 界面（index.html、config_editor.html）
    ✅ 包含配置文件模板（config.json）
    ✅ 自动清理构建临时文件
    ✅ 跨平台支持（Windows/Mac/Linux）
    ✅ 完整收集 pandas、openpyxl 等依赖
"""

import os
import sys
import shutil
import platform
import subprocess
from pathlib import Path
from datetime import datetime

class ExcelPriceUpdaterBuilder:
    """Excel价格批量更新工具 构建器"""
    
    def __init__(self, keep_temp=False, onefile=True):
        """初始化构建器"""
        self.project_root = Path(__file__).parent.resolve()
        self.system = platform.system()
        self.separator = ";" if self.system == "Windows" else ":"
        self.build_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.app_name = "excel_price_updater"
        self.version = "2.0.0"
        self.entry_point = "app.py"
        self.keep_temp = keep_temp
        self.onefile = onefile

        # 数据文件与目录
        self.data_includes = [
            ("index.html", "."),
            ("config_editor.html", "."),
            ("config.json", "."),
            ("README.md", "."),
        ]

        # 排除模块
        self.excludes = [
            "tkinter", "unittest", "test", "matplotlib", "scipy", 
            "IPython", "jupyter", "pkg_resources"
        ]

        # 隐藏导入
        self.hidden_imports = [
            # 第三方库
            "pandas", "openpyxl", "numpy", "tqdm",
            # pandas 依赖
            "pandas._libs", "pandas._libs.tslibs", "pandas._libs.tslibs.base",
            "pandas._libs.tslibs.timedeltas", "pandas._libs.tslibs.np_datetime",
            "pandas._libs.tslibs.nattype", "pandas._libs.tslibs.timestamps",
            # openpyxl 依赖
            "openpyxl.cell", "openpyxl.cell.cell", "openpyxl.styles",
            "openpyxl.worksheet", "openpyxl.worksheet.worksheet",
            # tqdm 依赖
            "tqdm.std", "tqdm.utils", "tqdm.auto", "tqdm.gui",
            # 标准库
            "json", "re", "random", "pathlib", "http.server",
            "urllib.parse", "email.parser", "io", "tempfile",
            # 避免 pkg_resources 相关错误
            "email", "email.mime", "email.mime.text"
        ]

    # ---------------------- 日志 ----------------------
    def log(self, msg, level="INFO"):
        """统一的日志输出"""
        icons = {
            "INFO": "ℹ️",
            "WARN": "⚠️",
            "ERROR": "❌",
            "SUCCESS": "✅"
        }
        icon = icons.get(level, "📝")
        print(f"{icon} [{level}] {msg}")

    # ---------------------- 环境检查 ----------------------
    def check_environment(self):
        """检查 Python 版本和依赖"""
        self.log("检查 Python 版本和依赖...")
        
        # 检查 Python 版本
        if sys.version_info < (3, 8):
            self.log("Python 版本过低，建议 >= 3.8", "WARN")
        else:
            self.log(f"Python 版本: {sys.version.split()[0]}", "SUCCESS")
        
        # 检查依赖包
        required_packages = {
            'pandas': 'pandas',
            'openpyxl': 'openpyxl',
            'numpy': 'numpy',
            'tqdm': 'tqdm',
        }
        
        missing_packages = []
        for pkg, mod in required_packages.items():
            try:
                __import__(mod)
                self.log(f"{pkg} 已安装", "SUCCESS")
            except ImportError:
                self.log(f"{pkg} 缺失", "WARN")
                missing_packages.append(pkg)
        
        # 安装缺失的依赖
        if missing_packages:
            self.log(f"正在安装缺失的依赖: {', '.join(missing_packages)}...")
            subprocess.run(
                [sys.executable, "-m", "pip", "install"] + missing_packages,
                check=True
            )
            self.log("依赖安装完成", "SUCCESS")

        # 检查 PyInstaller
        try:
            import PyInstaller
            self.log(f"PyInstaller 版本: {PyInstaller.__version__}", "SUCCESS")
        except ImportError:
            self.log("PyInstaller 未安装，正在安装...", "WARN")
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "pyinstaller"],
                check=True
            )
            try:
                import PyInstaller
                self.log(f"PyInstaller 安装完成: {PyInstaller.__version__}", "SUCCESS")
            except ImportError:
                self.log("PyInstaller 安装失败", "ERROR")
                sys.exit(1)

        # 检查入口文件
        entry_file = self.project_root / self.entry_point
        if not entry_file.exists():
            self.log(f"入口文件不存在: {entry_file}", "ERROR")
            sys.exit(1)
        self.log(f"入口文件: {entry_file}", "SUCCESS")

        # 检查数据文件
        missing_files = []
        for src, _ in self.data_includes:
            src_path = self.project_root / src
            if not src_path.exists():
                self.log(f"数据文件不存在，将跳过: {src}", "WARN")
                missing_files.append(src)
        
        if not missing_files:
            self.log("所有数据文件检查完成", "SUCCESS")

    # ---------------------- 清理旧产物 ----------------------
    def clean_build_artifacts(self):
        """清理旧的构建产物"""
        if self.keep_temp:
            self.log("保留旧构建产物", "INFO")
            return
        
        self.log("清理旧构建产物...")
        artifacts = ['build', 'dist', f'{self.app_name}.spec']
        
        cleaned_count = 0
        for artifact in artifacts:
            path = self.project_root / artifact
            if path.exists():
                if path.is_dir():
                    shutil.rmtree(path)
                else:
                    path.unlink()
                cleaned_count += 1
        
        # 清理 __pycache__
        pycache_count = 0
        for pyc in self.project_root.rglob("__pycache__"):
            if pyc.is_dir():
                shutil.rmtree(pyc)
                pycache_count += 1
        
        if cleaned_count > 0 or pycache_count > 0:
            self.log(f"清理完成（{cleaned_count} 个构建文件，{pycache_count} 个缓存目录）", "SUCCESS")

    # ---------------------- 构建 PyInstaller 命令 ----------------------
    def build_pyinstaller_command(self):
        """构建 PyInstaller 打包命令"""
        cmd = [
            "pyinstaller",
            "--noconfirm",
            "--clean",
            "--log-level=INFO"
        ]
        
        # 单文件或目录模式
        if self.onefile:
            cmd.append("--onefile")
        else:
            cmd.append("--onedir")
        
        # 应用名称
        cmd.extend(["--name", self.app_name])
        
        # 添加项目路径
        cmd.extend(["--paths", str(self.project_root)])
        
        # 禁用 UPX 压缩（避免某些兼容性问题）
        cmd.append("--noupx")
        
        # 收集子模块（确保所有依赖都被打包）
        cmd.extend(["--collect-all", "pandas"])
        cmd.extend(["--collect-all", "openpyxl"])
        cmd.extend(["--collect-all", "numpy"])
        cmd.extend(["--collect-all", "tqdm"])

        # 添加数据文件
        for src, dst in self.data_includes:
            src_path = self.project_root / src
            if src_path.exists():
                cmd.extend(["--add-data", f"{src}{self.separator}{dst}"])

        # 添加隐藏导入
        for mod in self.hidden_imports:
            cmd.extend(["--hidden-import", mod])

        # 排除模块（避免 pkg_resources 相关错误）
        for mod in self.excludes:
            cmd.extend(["--exclude-module", mod])
        
        # 禁用控制台窗口（如果是 Windows）
        # 注释掉此行可以看到控制台输出，方便调试
        # if self.system == "Windows":
        #     cmd.append("--noconsole")

        # 添加入口文件
        cmd.append(str(self.project_root / self.entry_point))
        
        return cmd

    # ---------------------- 执行构建 ----------------------
    def run_build(self):
        """执行 PyInstaller 打包"""
        cmd = self.build_pyinstaller_command()
        self.log(f"执行打包命令...")
        self.log(f"命令: {' '.join(cmd)}", "INFO")
        
        result = subprocess.run(cmd)
        
        if result.returncode != 0:
            self.log("打包失败", "ERROR")
            sys.exit(1)
        
        self.log("打包完成", "SUCCESS")

    # ---------------------- 创建发布包 ----------------------
    def create_release_package(self):
        """创建发布包"""
        release_name = f"{self.app_name}_{self.version}_{self.system}_{self.build_time}"
        release_dir = self.project_root / "release" / release_name
        release_dir.mkdir(parents=True, exist_ok=True)
        
        self.log(f"创建发布包: {release_name}...")
        
        # 获取可执行文件
        if self.onefile:
            exe_file = self.project_root / 'dist' / (
                f"{self.app_name}.exe" if self.system == "Windows" else self.app_name
            )
        else:
            dist_dir = self.project_root / 'dist' / self.app_name
            exe_file = dist_dir / (
                f"{self.app_name}.exe" if self.system == "Windows" else self.app_name
            )

        # 复制可执行文件或目录
        if self.onefile:
            if exe_file.exists():
                shutil.copy2(exe_file, release_dir)
                self.log(f"复制可执行文件: {exe_file.name}", "SUCCESS")
        else:
            if exe_file.parent.exists():
                shutil.copytree(exe_file.parent, release_dir / self.app_name)
                self.log(f"复制应用目录: {self.app_name}", "SUCCESS")

        # 复制配置文件模板
        config_src = self.project_root / 'config.json'
        config_dst = release_dir / 'config_template.json'
        if config_src.exists():
            shutil.copy2(config_src, config_dst)
            self.log("复制配置模板", "SUCCESS")

        # 复制 HTML 文件（作为备份）
        for html_file in ['index.html', 'config_editor.html']:
            html_src = self.project_root / html_file
            if html_src.exists():
                shutil.copy2(html_src, release_dir / html_file)

        # 复制文档
        for doc_file in ['README.md', 'requirements.txt']:
            doc_src = self.project_root / doc_file
            if doc_src.exists():
                shutil.copy2(doc_src, release_dir / doc_file)
                self.log(f"复制文档: {doc_file}", "SUCCESS")

        # 生成使用说明
        self._create_usage_guide(release_dir, exe_file.name)

        # 生成启动脚本
        self._create_startup_scripts(release_dir, exe_file.name if self.onefile else self.app_name)

        self.log(f"发布包创建成功: {release_dir}", "SUCCESS")
        return release_dir, exe_file

    # ---------------------- 生成使用说明 ----------------------
    def _create_usage_guide(self, release_dir, exe_name):
        """生成使用说明文档"""
        usage_content = f"""
========================================
Excel价格批量更新工具 使用说明
========================================

版本: {self.version}
系统: {self.system}
构建时间: {self.build_time}

========================================
📦 主要文件
========================================

- {exe_name}                主程序可执行文件
- config_template.json      配置文件模板
- index.html                Web界面（已内嵌）
- config_editor.html        配置编辑器（已内嵌）
- README.md                 详细文档
- USAGE.txt                 本文件

========================================
🚀 快速开始
========================================

方法一：使用启动脚本（推荐）
{'  - Windows: 双击 run.bat' if self.system == 'Windows' else '  - Mac/Linux: 双击 run.sh 或在终端运行 ./run.sh'}

方法二：命令行启动
  1. 打开终端/命令提示符
  2. 进入本目录
  3. 运行: ./{exe_name}

========================================
📝 使用步骤
========================================

1. 启动程序后，浏览器会自动打开
   访问地址: http://localhost:8800

2. 配置产品价格（第一次使用）
   - 点击"⚙️ 配置管理"标签
   - 添加产品和价格规则
   - 支持固定价格或区间定价
   - 点击"💾 保存配置"

3. 批量更新Excel价格
   - 点击"🔄 价格更新"标签
   - 拖拽或选择Excel文件
   - 选择要更新的地域（HK/SG/MY等）
   - 点击"开始处理"
   - 下载处理后的文件

========================================
⚙️ 配置说明
========================================

配置文件格式（config.json）：

{{
  "产品名称": {{
    "hk": [最小价, 最大价, 步长],  // 区间定价
    "sg": 固定价格,                 // 固定定价
    "my": [min, max, step]
  }}
}}

示例：
{{
  "Nike Air Force 1": {{
    "hk": [550, 580, 10],  // HK: 550-580之间，10的倍数
    "sg": [70, 85, 5],     // SG: 70-85之间，5的倍数
    "my": [50, 60, 10]     // MY: 50-60之间，10的倍数
  }},
  "Adidas Samba": {{
    "hk": 450,             // HK: 固定价格450
    "sg": 60,              // SG: 固定价格60
    "my": 45               // MY: 固定价格45
  }}
}}

========================================
📊 Excel文件格式要求
========================================

必需列：
  - ProductNameCn  （产品中文名称）
  - {{REGION}}Price   （各地域价格列，如 HKPrice, SGPrice）

示例：
  | ProductNameCn        | HKPrice | SGPrice | MYPrice |
  |---------------------|---------|---------|---------|
  | Nike Air Force 1    | 565     | 75      | 55      |
  | Adidas Samba       | 450     | 60      | 45      |

========================================
🔍 匹配规则
========================================

产品名称匹配规则：
  - 优先匹配最具体的名称
  - 不区分大小写
  - 支持部分匹配

示例：
  配置中有 "samba" 和 "samba og"
  Excel中 "Adidas Samba OG Triple Black"
  → 匹配到 "samba og"（更具体）

========================================
❓ 常见问题
========================================

Q: 如何添加新地域？
A: 在配置编辑器中，点击"🌍 添加地域"按钮

Q: 如何修改已有产品价格？
A: 在配置编辑器中，找到产品并修改价格

Q: 支持哪些地域？
A: 支持任意地域，常见的有：
   HK(香港), SG(新加坡), MY(马来西亚), TW(台湾),
   JP(日本), KR(韩国), ID(印尼), TH(泰国), PH(菲律宾)

Q: 程序无法启动？
A: 检查端口8800是否被占用，或联系技术支持

Q: 如何批量处理多个文件？
A: 可以一次选择多个Excel文件进行处理

========================================
🛠️ 技术支持
========================================

遇到问题？
  1. 查看 README.md 获取详细文档
  2. 检查配置文件格式是否正确
  3. 查看 USAGE.txt 获取使用说明

========================================
"""
        
        with open(release_dir / 'USAGE.txt', 'w', encoding='utf-8') as f:
            f.write(usage_content.strip())
        
        self.log("生成使用说明", "SUCCESS")

    # ---------------------- 生成启动脚本 ----------------------
    def _create_startup_scripts(self, release_dir, exe_name):
        """生成启动脚本"""
        if self.system == "Windows":
            # Windows 批处理脚本
            bat_content = f"""@echo off
chcp 65001 >nul
title Excel价格批量更新工具
cls
echo ========================================
echo  Excel价格批量更新工具 v{self.version}
echo ========================================
echo.
echo 正在启动程序...
echo 程序启动后会自动打开浏览器
echo 访问地址: http://localhost:8800
echo.
echo 按 Ctrl+C 可以停止程序
echo ========================================
echo.

{exe_name}

if errorlevel 1 (
    echo.
    echo ❌ 程序运行出错！
    echo.
    pause
)
"""
            bat_file = release_dir / 'run.bat'
            with open(bat_file, 'w', encoding='utf-8') as f:
                f.write(bat_content)
            self.log("生成启动脚本: run.bat", "SUCCESS")
        else:
            # Unix/Mac Shell 脚本
            sh_content = f"""#!/bin/bash

# Excel价格批量更新工具启动脚本

SCRIPT_DIR="$( cd "$( dirname "${{BASH_SOURCE[0]}}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "========================================"
echo " Excel价格批量更新工具 v{self.version}"
echo "========================================"
echo ""
echo "正在启动程序..."
echo "程序启动后会自动打开浏览器"
echo "访问地址: http://localhost:8800"
echo ""
echo "按 Ctrl+C 可以停止程序"
echo "========================================"
echo ""

./{exe_name}

if [ $? -ne 0 ]; then
    echo ""
    echo "❌ 程序运行出错！"
    echo ""
    read -p "按 Enter 键继续..."
fi
"""
            sh_file = release_dir / 'run.sh'
            with open(sh_file, 'w', encoding='utf-8') as f:
                f.write(sh_content)
            os.chmod(sh_file, 0o755)
            self.log("生成启动脚本: run.sh", "SUCCESS")

    # ---------------------- 自动清理临时文件 ----------------------
    def auto_cleanup_temp_files(self):
        """构建完成后自动清理临时文件"""
        self.log("自动清理构建临时文件...")
        
        temp_items = ['build', 'dist', f'{self.app_name}.spec']
        cleaned_count = 0
        
        for item in temp_items:
            item_path = self.project_root / item
            if item_path.exists():
                if item_path.is_dir():
                    shutil.rmtree(item_path)
                else:
                    item_path.unlink()
                cleaned_count += 1
                self.log(f"删除: {item}", "INFO")
        
        # 清理 __pycache__
        pycache_count = 0
        for pycache in self.project_root.rglob('__pycache__'):
            if pycache.is_dir():
                shutil.rmtree(pycache)
                pycache_count += 1
        
        if pycache_count > 0:
            self.log(f"删除 {pycache_count} 个 __pycache__ 目录", "INFO")
        
        self.log(f"临时文件清理完成（共 {cleaned_count + pycache_count} 项）", "SUCCESS")

    # ---------------------- 构建流程 ----------------------
    def build(self):
        """执行完整构建流程"""
        try:
            print("\n" + "=" * 60)
            self.log("🚀 Excel价格批量更新工具 构建开始")
            print("=" * 60 + "\n")
            
            # 1. 环境检查
            self.check_environment()
            print()
            
            # 2. 清理旧产物
            self.clean_build_artifacts()
            print()
            
            # 3. 执行构建
            self.run_build()
            print()
            
            # 4. 创建发布包
            release_dir, exe_file = self.create_release_package()
            print()
            
            # 5. 自动清理临时文件（除非设置了 keep_temp）
            if not self.keep_temp:
                self.auto_cleanup_temp_files()
                print()
            
            # 6. 显示完成信息
            print("\n" + "=" * 60)
            self.log(f"🎉 构建完成！", "SUCCESS")
            print("=" * 60)
            print(f"\n📦 可执行文件: {exe_file.name}")
            print(f"📂 发布包位置: {release_dir}")
            print(f"📊 发布包大小: {self._get_dir_size(release_dir):.2f} MB")
            print(f"\n💡 提示:")
            print(f"   1. 进入发布目录: cd {release_dir}")
            print(f"   2. 运行程序: {'run.bat' if self.system == 'Windows' else './run.sh'}")
            print(f"   3. 访问: http://localhost:8800")
            print("\n" + "=" * 60 + "\n")
            
        except KeyboardInterrupt:
            print("\n")
            self.log("用户中断构建", "WARN")
            sys.exit(1)
        except Exception as e:
            import traceback
            print("\n")
            self.log(f"构建出错: {e}", "ERROR")
            traceback.print_exc()
            sys.exit(1)

    def _get_dir_size(self, path):
        """计算目录大小（MB）"""
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(path):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                if os.path.exists(filepath):
                    total_size += os.path.getsize(filepath)
        return total_size / (1024 * 1024)

# ---------------------- 主函数 ----------------------
def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Excel价格批量更新工具 部署脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python deploy.py                # 默认构建（单文件模式，自动清理）
  python deploy.py --keep-temp    # 保留临时文件
  python deploy.py --onedir       # 使用目录模式
        """
    )
    
    parser.add_argument(
        '--keep-temp',
        action='store_true',
        help='保留临时文件（build、dist、*.spec）'
    )
    parser.add_argument(
        '--onedir',
        action='store_true',
        help='使用目录模式（默认为单文件模式）'
    )
    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 2.0.0'
    )
    
    args = parser.parse_args()
    
    builder = ExcelPriceUpdaterBuilder(
        keep_temp=args.keep_temp,
        onefile=not args.onedir
    )
    builder.build()

if __name__ == '__main__':
    main()

