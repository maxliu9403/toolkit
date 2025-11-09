# Excel 工具箱

一个集成多功能的 Excel 处理工具，提供价格批量更新、BrowserID 替换、Excel 拆分合并等功能。

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## ✨ 功能特点

### 📋 Excel 拆分/合并
- **拆分模式**: 按 BrowserID 将 Excel 文件拆分成多个文件
- **合并模式**: 将多个 Excel 文件合并成一个文件
- **智能处理**: 自动添加来源信息，保持数据完整性
- **批量操作**: 支持同时处理多个文件

### 🔄 BrowserID 替换
- 根据封号数据表批量替换 BrowserID
- 支持多个目标文件同时处理
- 提供详细的替换统计信息
- 自动匹配和替换

### 📈 价格批量更新
- **智能匹配**: 通过正则表达式匹配产品名称
- **灵活定价**: 支持固定价格和价格区间两种方式
- **多地域支持**: 支持任意地域（HK、SG、MY、TW、JP、KR等）
- **随机定价**: 在指定区间内自动生成符合步长要求的价格

### ⚙️ 配置管理
- **可视化编辑**: Web界面可视化编辑价格配置
- **实时保存**: 自动保存配置修改
- **格式验证**: 自动验证配置格式

## 🚀 快速开始

### 安装依赖

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate  # macOS/Linux
# venv\Scripts\activate    # Windows

# 安装依赖
pip install -r requirements.txt
```

### 启动服务

```bash
# 方法一：使用启动脚本（推荐）
./start.sh

# 方法二：手动启动
source venv/bin/activate
python app.py

# 方法三：重启服务
./restart.sh
```

服务启动后，访问: **http://localhost:8800**

### 停止服务

```bash
./stop.sh
```

## 📖 使用指南

### 1. Excel 拆分/合并

#### 拆分模式
1. 访问第一个标签页 "📋 Excel拆分/合并"
2. 选择 "拆分模式"
3. 设置店铺数量（默认20）
4. 上传需要拆分的 Excel 文件（支持拖拽）
5. 点击 "开始拆分"
6. 自动下载 ZIP 压缩包

#### 合并模式
1. 选择 "合并模式"
2. 上传多个 Excel 文件
3. 点击 "开始合并"
4. 自动下载合并后的文件

### 2. BrowserID 替换

1. 访问 "🔄 BrowserID替换" 标签页
2. 上传封号数据表（包含"封号ID"和"新对应ID"列）
3. 上传需要替换的目标 Excel 文件
4. 点击 "开始替换BrowserID"
5. 自动下载替换后的文件

**封号数据表格式**:
```
| 封号ID | 新对应ID |
|--------|----------|
| 123    | 456      |
| 789    | 101      |
```

### 3. 价格更新

1. 访问 "📈 价格更新" 标签页
2. 上传 Excel 文件
3. 选择需要更新的地域（如 HK、SG、MY）
4. 点击 "开始更新价格"
5. 自动下载更新后的文件

**Excel 文件格式要求**:
- 必须包含 `ProductNameCn` 列（产品中文名称）
- 必须包含对应的价格列（如 `HKPrice`、`SGPrice`）

### 4. 配置管理

1. 访问 "⚙️ 配置管理" 标签页
2. 添加/编辑产品配置
3. 设置价格（固定价格或区间价格）
4. 点击 "保存配置"

**配置格式示例**:

```json
{
  "Nike Air Force 1": {
    "hk": [550, 580, 10],
    "sg": [70, 85, 5],
    "my": [50, 60, 10]
  },
  "Adidas Samba": {
    "hk": 450,
    "sg": 60,
    "my": 45
  }
}
```

**价格区间格式**: `[最小值, 最大值, 步长]`
- 程序会在区间内随机生成符合步长的价格
- 示例：`[550, 580, 10]` 会生成 550、560、570 或 580

## 🔧 服务管理

### 启动服务
```bash
./start.sh
```
- 后台运行
- 日志输出到 `server.log`
- 显示进程ID和访问地址

### 停止服务
```bash
./stop.sh
```

### 重启服务
```bash
./restart.sh
```
- 前台运行，适合开发调试
- 按 Ctrl+C 停止

### 查看日志
```bash
# 实时查看
tail -f server.log

# 查看全部
cat server.log
```

### 检查服务状态
```bash
# 查看进程
ps aux | grep "python.*app.py"

# 测试API
curl http://localhost:8800/api/regions
```

## 📁 项目结构

```
toolkit/
├── modules/                      # 功能模块
│   ├── browserid_replacer.py    # BrowserID替换模块
│   ├── price_updater.py         # 价格更新模块
│   └── split_excel.py           # Excel拆分/合并模块
│
├── app.py                        # Web服务器（主程序）
├── main.py                       # 命令行工具
│
├── index.html                    # Web主界面
├── config_editor.html            # 配置编辑器页面
├── config.json                   # 价格配置文件
│
├── start.sh                      # 启动脚本
├── stop.sh                       # 停止脚本
├── restart.sh                    # 重启脚本
├── server.log                    # 服务日志
│
├── deploy.py                     # 部署打包脚本
├── requirements.txt              # 依赖列表
└── README.md                     # 本文件
```

## 🎯 API 接口

### Excel 拆分/合并
- `POST /api/split_excel` - 拆分 Excel 文件
- `POST /api/merge_excel` - 合并 Excel 文件

### BrowserID 替换
- `POST /api/upload_ban_data` - 上传封号数据表
- `POST /api/replace_browserid` - 替换 BrowserID

### 价格更新
- `POST /api/process` - 处理 Excel 文件更新价格

### 配置管理
- `GET /api/config` - 获取配置
- `POST /api/config` - 保存配置
- `GET /api/regions` - 获取可用地域

### 文件下载
- `GET /api/download/<filename>` - 下载处理后的文件

## 💻 命令行使用

如果需要在命令行中使用（不启动Web服务）：

```bash
source venv/bin/activate
python main.py
```

按照提示输入文件路径和地域即可。

## 📦 部署打包

### 打包为可执行文件

```bash
# 1. 激活虚拟环境
source venv/bin/activate

# 2. 运行部署脚本
python deploy.py

# 3. 发布包位置
cd release/excel_price_updater_*

# 4. 运行
./run.sh  # Mac/Linux
run.bat   # Windows
```

**部署选项**:
- `python deploy.py` - 单文件模式（默认）
- `python deploy.py --keep-temp` - 保留临时文件
- `python deploy.py --onedir` - 目录模式

## ❓ 常见问题

### Q: 服务无法启动？
A: 检查以下几点：
1. 端口 8800 是否被占用
   ```bash
   lsof -i :8800
   ```
2. 是否安装了所有依赖
   ```bash
   pip install -r requirements.txt
   ```
3. 是否激活了虚拟环境
   ```bash
   source venv/bin/activate
   ```

### Q: 如何添加新地域？
A: 在配置编辑器中直接添加新地域配置即可，系统会自动检测。

### Q: 如何修改端口？
A: 编辑 `app.py`，修改 `start_server(port=8800)` 中的端口号。

### Q: 支持哪些Excel格式？
A: 支持 `.xlsx` 和 `.xls` 格式。

### Q: 如何批量处理多个文件？
A: 在Web界面中可以一次选择多个文件进行处理。

### Q: 匹配规则是什么？
A: 
- 使用正则表达式匹配产品名称
- 不区分大小写
- 优先匹配更长（更具体）的关键字

示例：
```
配置: "samba" 和 "samba og"
产品: "Adidas Samba OG"
结果: 匹配到 "samba og"（更具体）
```

## 🔄 更新日志

### v2.0.0 (2025-11-09)
- ✨ 新增 Excel 拆分/合并功能
- ✨ 集成所有功能到 Web 界面
- 🎨 优化项目结构，模块化代码
- 📝 完善文档和使用说明
- 🚀 添加启动/停止脚本
- 🧹 清理无用代码和文件

### v1.0.0
- 🎉 初始版本
- 价格批量更新功能
- BrowserID 替换功能
- 配置编辑器

## 🛠️ 技术栈

- **Python 3.8+**
- **pandas** - Excel 数据处理
- **openpyxl** - Excel 读写
- **tqdm** - 进度条显示
- **http.server** - Web 服务器

## 📝 依赖包

主要依赖：
```
pandas>=2.0.0
openpyxl>=3.0.0
numpy>=1.24.0
tqdm>=4.65.0
```

完整依赖列表请查看 `requirements.txt`。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT License

---

## 📚 附录

### 匹配逻辑详解

1. 读取 Excel 中的 `ProductNameCn` 列
2. 在配置文件中查找所有关键字
3. 使用正则表达式（不区分大小写）匹配
4. **优先匹配更长的关键字**
5. 找到匹配后使用对应配置更新价格

### 价格配置说明

**固定价格**:
```json
{
  "product": {
    "hk": 100,
    "sg": 50
  }
}
```

**价格区间**:
```json
{
  "product": {
    "hk": [500, 600, 10],
    "sg": [70, 85, 5]
  }
}
```

**混合配置**:
```json
{
  "product": {
    "hk": [500, 600, 10],
    "sg": 70,
    "my": 50
  }
}
```

### 清理维护

**定期清理**:
```bash
# 清理缓存
rm -rf __pycache__ modules/__pycache__
find . -name "*.pyc" -delete

# 清空日志
> server.log
```

**重建虚拟环境**:
```bash
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Git 忽略

建议 `.gitignore` 包含：
```
__pycache__/
*.pyc
*.pyo
venv/
server.log
*.xlsx
*.xls
!sample*.xlsx
```

---

**更新时间**: 2025-11-09  
**版本**: v2.0.0  
**维护**: 持续更新中

如有问题或建议，欢迎反馈！🎉
