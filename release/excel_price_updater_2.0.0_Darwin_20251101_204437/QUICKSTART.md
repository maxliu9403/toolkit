# 快速开始指南

## 1. 环境准备

首先创建并激活虚拟环境：

```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate  # macOS/Linux
# 或在Windows上: venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

## 2. 准备配置文件和Excel文件

### 配置文件 (config.json)

确保你的 `config.json` 文件包含产品关键字和对应的价格。支持两种配置方式：

**方式1：固定价格**
```json
{
  "example_product": {
    "hk": 100,
    "sg": 50,
    "my": 80
  }
}
```

**方式2：价格区间（随机生成）**
```json
{
  "Nike Air force 1": {
    "hk": [550, 580, 10],
    "sg": [70, 85, 5],
    "my": [50, 60, 10]
  }
}
```

**价格区间格式**: `[最小值, 最大值, 步长]`

程序会在区间内随机生成符合步长要求的价格。例如`[550, 580, 10]`会随机生成 550, 560, 570, 580 中的一个值。

**注意**：
- 配置多个相关关键字时（如"samba"和"samba a"），更长的关键字会优先匹配
- 使用价格区间时，最小值和最大值都必须是步长的倍数

### Excel文件格式

你的Excel文件必须包含以下列：

- `ProductNameCn` - 产品中文名称（用于正则匹配）
- `HKPrice` - 香港价格
- `SGPrice` - 新加坡价格  
- `MYPrice` - 马来西亚价格

## 3. 运行程序

```bash
source venv/bin/activate  # 如果还没激活虚拟环境
python main.py
```

## 4. 交互式使用示例

### 步骤1: 输入Excel文件

```
请输入Excel文件路径（直接回车结束输入）: sample_products.xlsx
请输入Excel文件路径（直接回车结束输入）: 
总共选择了 1 个文件:
  1. sample_products.xlsx
```

### 步骤2: 选择地域

```
可用地域:
  HK   -> HKPrice
  SG   -> SGPrice
  MY   -> MYPrice

请输入需要更新的地域（多个用逗号分隔，如: hk,sg,my）:
地域代码: hk,sg,my
```

### 步骤3: 查看结果

```
正在验证配置文件...
✓ 配置文件验证通过

============================================================
开始处理文件...
============================================================

正在处理: sample_products.xlsx
✓ 成功更新 5 条记录
✓ 已保存到: sample_products_updated.xlsx

============================================================
处理完成!
============================================================
成功: 1 个文件
失败: 0 个文件
```

## 5. 使用示例文件测试

项目包含了一个示例Excel文件 `sample_products.xlsx`，你可以用它来测试程序：

```bash
python main.py
# 然后输入: sample_products.xlsx
# 地域选择: hk,sg,my
```

## 常见问题

### Q: 程序提示"无法匹配产品的价格配置"

A: 检查你的 `config.json` 中是否包含该产品的关键字。例如，如果Excel中有产品叫"Samba产品A"，你的配置文件中需要有"samba"这个key。

### Q: 程序提示"产品缺少某地域的价格配置"

A: 确保配置文件中每个产品都包含了所选地域的价格。如果你选择了hk,sg,my三个地域，每个产品的配置都必须包含这三个地域。

### Q: 程序提示"Excel文件缺少必需的列"

A: 确保你的Excel文件包含所有必需的列：ProductNameCn, HKPrice, SGPrice, MYPrice。

### Q: 如何批量处理多个文件？

A: 在输入文件时，你可以使用通配符，例如：
- `*.xlsx` - 处理当前目录下所有xlsx文件
- `data/*.xlsx` - 处理data目录下所有xlsx文件
- 或者手动输入多个文件路径（每行一个）
