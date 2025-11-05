#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试BrowserID替换修复 - 整数类型处理
"""

import pandas as pd
from pathlib import Path

# 创建测试用的封号数据表（使用整数）
ban_data = {
    '封号ID': [520, 478, 514, 521, 463],
    '新对应ID': [1001, 1002, 1003, 1004, 1005]
}
ban_df = pd.DataFrame(ban_data)
ban_file = Path('test_ban_integers.xlsx')
ban_df.to_excel(ban_file, index=False, engine='openpyxl')
print(f"✓ 创建测试封号数据表: {ban_file}")
print(f"  封号ID类型: {ban_df['封号ID'].dtype}")
print(f"  封号ID值: {ban_df['封号ID'].tolist()}")

# 创建测试用的目标Excel文件（BrowserID为整数）
target_data = {
    'SKU': [1330294857, 1391251104, 1351577734, 1334111858, 1331858619, 1234567890],
    'BrowserID': [520, 478, 999, 514, 521, 463],  # 整数类型
    'ProductName': ['Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product F']
}
target_df = pd.DataFrame(target_data)
target_file = Path('test_target_integers.xlsx')
target_df.to_excel(target_file, index=False, engine='openpyxl')
print(f"\n✓ 创建测试目标文件: {target_file}")
print(f"  BrowserID类型: {target_df['BrowserID'].dtype}")
print(f"  BrowserID值: {target_df['BrowserID'].tolist()}")

# 测试BrowserID替换功能
from main import BrowserIDReplacer

print("\n" + "="*60)
print("开始测试BrowserID替换...")
print("="*60)

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
expected_ids = [1001, 1002, 999, 1003, 1004, 1005]
actual_ids = output_df['BrowserID'].tolist()

print(f"\n详细对比:")
print(f"  期望值: {expected_ids}")
print(f"  实际值: {actual_ids}")
print(f"  期望类型: {type(expected_ids[0])}")
print(f"  实际类型: {type(actual_ids[0])}")

if actual_ids == expected_ids:
    print(f"\n✓✓✓ 测试通过！BrowserID替换正确且保持了整数类型")
else:
    print(f"\n✗✗✗ 测试失败！")
    for i, (exp, act) in enumerate(zip(expected_ids, actual_ids)):
        if exp != act:
            print(f"  行{i}: 期望={exp}, 实际={act}")

# 清理测试文件
import os
print(f"\n清理测试文件...")
os.remove(ban_file)
os.remove(target_file)
os.remove(result['output_file'])
print(f"✓ 测试文件已清理")

