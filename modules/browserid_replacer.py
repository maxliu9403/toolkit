#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BrowserID替换器模块
"""

import os
from pathlib import Path
from typing import List, Dict
import pandas as pd
import numpy as np


class BrowserIDReplacer:
    """BrowserID替换器"""
    
    def __init__(self):
        """初始化BrowserID替换器"""
        self.ban_data = None  # 封号数据表
    
    def load_ban_data(self, ban_file: str) -> bool:
        """
        加载封号数据表
        
        Args:
            ban_file: 封号数据表文件路径
            
        Returns:
            是否成功加载
            
        Raises:
            FileNotFoundError: 文件不存在
            KeyError: 必需的列不存在
        """
        if not os.path.exists(ban_file):
            raise FileNotFoundError(f"封号数据表文件不存在: {ban_file}")
        
        print(f"\n正在加载封号数据表: {ban_file}")
        self.ban_data = pd.read_excel(ban_file)
        
        # 检查必需的列是否存在
        required_columns = ['封号ID', '新对应ID']
        missing_columns = [col for col in required_columns if col not in self.ban_data.columns]
        
        if missing_columns:
            raise KeyError(
                f"封号数据表缺少必需的列: {', '.join(missing_columns)}\n"
                f"当前列: {', '.join(self.ban_data.columns)}"
            )
        
        # 创建封号ID到新ID的映射字典
        # 先转为字符串，去除可能的空格，处理NaN值
        ban_ids = []
        new_ids = []
        for idx, row in self.ban_data.iterrows():
            ban_id = row['封号ID']
            new_id = row['新对应ID']
            
            # 跳过NaN值
            if pd.isna(ban_id) or pd.isna(new_id):
                continue
            
            # 如果是数字，转为整数再转字符串（避免520.0这样的浮点数）
            if isinstance(ban_id, (int, float)):
                ban_id = str(int(ban_id))
            else:
                ban_id = str(ban_id).strip()
            
            if isinstance(new_id, (int, float)):
                new_id = str(int(new_id))
            else:
                new_id = str(new_id).strip()
            
            ban_ids.append(ban_id)
            new_ids.append(new_id)
        
        self.ban_mapping = dict(zip(ban_ids, new_ids))
        
        print(f"✓ 成功加载封号数据表，共 {len(self.ban_mapping)} 条记录")
        print(f"  示例映射（前3条）:")
        for i, (old_id, new_id) in enumerate(list(self.ban_mapping.items())[:3]):
            print(f"    {old_id} -> {new_id}")
        return True
    
    def replace_browser_id(self, excel_file: str, output_suffix: str = "_replaced") -> Dict:
        """
        替换Excel文件中的BrowserID
        
        Args:
            excel_file: Excel文件路径
            output_suffix: 输出文件后缀
            
        Returns:
            处理结果字典，包含成功/失败信息和统计数据
            
        Raises:
            FileNotFoundError: 文件不存在
            KeyError: 必需的列不存在
            ValueError: 数据格式错误
        """
        if self.ban_data is None:
            raise ValueError("请先加载封号数据表")
        
        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"目标文件不存在: {excel_file}")
        
        # 读取Excel文件
        print(f"\n正在处理: {excel_file}")
        df = pd.read_excel(excel_file)
        
        # 检查BrowserID列是否存在
        if 'BrowserID' not in df.columns:
            raise KeyError(
                f"目标Excel文件缺少必需的列: BrowserID\n"
                f"当前列: {', '.join(df.columns)}"
            )
        
        # 统计信息
        replaced_count = 0
        not_found_count = 0
        replaced_details = []  # 记录替换详情
        
        # 遍历每一行，替换BrowserID
        for idx, row in df.iterrows():
            original_id = row['BrowserID']
            
            # 转换BrowserID为字符串（处理整数/浮点数）
            if pd.isna(original_id):
                not_found_count += 1
                continue
                
            if isinstance(original_id, (int, float)):
                browser_id = str(int(original_id))
            else:
                browser_id = str(original_id).strip()
            
            # 检查是否在封号列表中
            if browser_id in self.ban_mapping:
                new_id = self.ban_mapping[browser_id]
                
                # 根据原始列的数据类型来决定新值的类型
                if isinstance(original_id, (int, np.integer)):
                    # 如果原始是整数，尝试将新ID也转为整数
                    try:
                        df.at[idx, 'BrowserID'] = int(new_id)
                    except ValueError:
                        df.at[idx, 'BrowserID'] = new_id
                elif isinstance(original_id, (float, np.floating)):
                    # 如果原始是浮点数，尝试将新ID转为浮点数
                    try:
                        df.at[idx, 'BrowserID'] = float(new_id)
                    except ValueError:
                        df.at[idx, 'BrowserID'] = new_id
                else:
                    df.at[idx, 'BrowserID'] = new_id
                
                replaced_count += 1
                replaced_details.append(f"{browser_id} -> {new_id}")
            else:
                not_found_count += 1
        
        # 保存更新后的文件
        output_file = self._get_output_filename(excel_file, output_suffix)
        df.to_excel(output_file, index=False, engine='openpyxl')
        
        result = {
            'success': True,
            'output_file': output_file,
            'total_count': len(df),
            'replaced_count': replaced_count,
            'not_found_count': not_found_count
        }
        
        print(f"✓ 处理完成")
        print(f"  总记录数: {result['total_count']}")
        print(f"  替换数: {result['replaced_count']}")
        print(f"  未匹配数: {result['not_found_count']}")
        if replaced_details:
            print(f"  替换详情（前5条）:")
            for detail in replaced_details[:5]:
                print(f"    {detail}")
        print(f"✓ 已保存到: {output_file}")
        
        return result
    
    def batch_replace(self, excel_files: List[str], ban_file: str, 
                     output_suffix: str = "_replaced") -> Dict:
        """
        批量替换多个Excel文件中的BrowserID
        
        Args:
            excel_files: Excel文件路径列表
            ban_file: 封号数据表文件路径
            output_suffix: 输出文件后缀
            
        Returns:
            批处理结果字典
        """
        results = {
            'success_files': [],
            'failed_files': [],
            'total_replaced': 0,
            'total_not_found': 0
        }
        
        # 加载封号数据表
        try:
            self.load_ban_data(ban_file)
        except (FileNotFoundError, KeyError) as e:
            return {
                'success': False,
                'error': str(e)
            }
        
        # 批量处理文件
        print("\n" + "="*60)
        print("开始批量处理文件...")
        print("="*60)
        
        for excel_file in excel_files:
            try:
                result = self.replace_browser_id(excel_file, output_suffix)
                results['success_files'].append({
                    'file': excel_file,
                    'output': result['output_file'],
                    'replaced_count': result['replaced_count'],
                    'not_found_count': result['not_found_count']
                })
                results['total_replaced'] += result['replaced_count']
                results['total_not_found'] += result['not_found_count']
            except (FileNotFoundError, KeyError, ValueError) as e:
                results['failed_files'].append({
                    'file': excel_file,
                    'error': str(e)
                })
        
        results['success'] = True
        return results
    
    def _get_output_filename(self, filepath: str, suffix: str) -> str:
        """
        生成输出文件名
        
        Args:
            filepath: 原始文件路径
            suffix: 后缀
            
        Returns:
            输出文件路径
        """
        path = Path(filepath)
        output_path = path.parent / f"{path.stem}{suffix}{path.suffix}"
        return str(output_path)

