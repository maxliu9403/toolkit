#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试BrowserID替换功能
"""

import pandas as pd
from pathlib import Path

# 创建测试用的封号数据表
ban_data = {
    '封号ID': ['BR001', 'BR002', 'BR003', 'BR004', 'BR005'],
    '新对应ID': ['NEW001', 'NEW002', 'NEW003', 'NEW004', 'NEW005']
}
ban_df = pd.DataFrame(ban_data)
ban_file = Path('test_ban_data.xlsx')
ban_df.to_excel(ban_file, index=False, engine='openpyxl')
print(f"✓ 创建测试封号数据表: {ban_file}")

# 创建测试用的目标Excel文件
target_data = {
    'BrowserID': ['BR001', 'BR002', 'BR999', 'BR003', 'BR888', 'BR004'],
    'ProductName': ['Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product F'],
    'Price': [100, 200, 300, 400, 500, 600]
}
target_df = pd.DataFrame(target_data)
target_file = Path('test_target.xlsx')
target_df.to_excel(target_file, index=False, engine='openpyxl')
print(f"✓ 创建测试目标文件: {target_file}")

# 测试BrowserID替换功能
from main import BrowserIDReplacer

replacer = BrowserIDReplacer()
replacer.load_ban_data(str(ban_file))

result = replacer.replace_browser_id(str(target_file))

print(f"\n测试结果:")
print(f"  总记录数: {result['total_count']}")
print(f"  替换数: {result['replaced_count']}")
print(f"  未匹配数: {result['not_found_count']}")
print(f"  输出文件: {result['output_file']}")

# 验证结果
output_df = pd.read_excel(result['output_file'])
print(f"\n输出文件内容:")
print(output_df)

# 验证替换是否正确
expected_ids = ['NEW001', 'NEW002', 'BR999', 'NEW003', 'BR888', 'NEW004']
actual_ids = output_df['BrowserID'].tolist()

if actual_ids == expected_ids:
    print(f"\n✓ 测试通过！BrowserID替换正确")
else:
    print(f"\n✗ 测试失败！")
    print(f"  期望: {expected_ids}")
    print(f"  实际: {actual_ids}")

# 清理测试文件
# ban_file.unlink()
# target_file.unlink()
# Path(result['output_file']).unlink()
print(f"\n测试文件已保留供查看")

