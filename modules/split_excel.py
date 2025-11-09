#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel拆分/合并模块
按 BrowserID 拆分或合并多个 Excel 文件
"""

import os
from glob import glob
from pathlib import Path
from typing import List, Dict, Optional
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime


class ExcelSplitter:
    """Excel拆分/合并工具类"""
    
    def __init__(self):
        """初始化Excel拆分器"""
        pass
    
    @staticmethod
    def log(msg: str, level: str = "INFO"):
        """统一日志输出"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] [{level}] {msg}")
    
    def collect_excels(self, root_folder: str) -> List[str]:
        """
        递归搜集所有子文件夹下的 Excel 文件路径
        
        Args:
            root_folder: 根目录路径
            
        Returns:
            Excel文件路径列表
        """
        self.log(f"开始扫描目录: {root_folder}")
        
        # 检查目录是否存在
        if not os.path.exists(root_folder):
            self.log(f"目录不存在: {root_folder}", "ERROR")
            return []
        
        if not os.path.isdir(root_folder):
            self.log(f"路径不是目录: {root_folder}", "ERROR")
            return []
        
        # 搜集所有 .xlsx 和 .xls 文件
        xlsx_files = glob(os.path.join(root_folder, "**/*.xlsx"), recursive=True)
        xls_files = glob(os.path.join(root_folder, "**/*.xls"), recursive=True)
        excel_files = xlsx_files + xls_files
        
        # 过滤掉临时文件（以 ~$ 开头）
        excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
        
        self.log(f"共检测到 {len(excel_files)} 个 Excel 文件 (.xlsx: {len(xlsx_files)}, .xls: {len(xls_files)})")
        
        if excel_files:
            self.log("找到的文件列表:")
            for idx, file in enumerate(excel_files, 1):
                rel_path = os.path.relpath(file, root_folder)
                self.log(f"  {idx}. {rel_path}")
        
        return excel_files
    
    def split_excel(self, file_path: str, group_size: int) -> Optional[pd.DataFrame]:
        """
        读取 Excel 文件并添加来源信息
        
        Args:
            file_path: Excel文件路径
            group_size: 分组大小（此参数保留用于兼容性）
            
        Returns:
            包含来源信息的DataFrame，如果失败返回None
        """
        try:
            # 读取 Excel 文件
            df = pd.read_excel(file_path)
            
            # 检查是否为空
            if df.empty:
                self.log(f"文件为空: {os.path.basename(file_path)}", "WARNING")
                return None
            
            # 检查是否存在 BrowserID 列
            if 'BrowserID' not in df.columns:
                self.log(f"文件缺少 BrowserID 列: {os.path.basename(file_path)}", "WARNING")
                return None
            
            # 添加来源信息
            df['SourceFile'] = os.path.basename(file_path)
            df['SourceFolder'] = os.path.basename(os.path.dirname(file_path))
            
            return df
            
        except Exception as e:
            self.log(f"文件读取失败: {os.path.basename(file_path)} - {str(e)}", "ERROR")
            return None
    
    def split_by_browser_id(self, root_folder: str, group_size: int, output_folder: str) -> Dict:
        """
        按 BrowserID 拆分Excel文件
        
        Args:
            root_folder: 输入文件夹路径
            group_size: 每组文件的数量
            output_folder: 输出文件夹路径
            
        Returns:
            处理结果字典
        """
        self.log("=" * 80)
        self.log("开始处理 Excel 按 BrowserID 拆分任务")
        self.log("=" * 80)
        
        # 显示配置参数
        self.log(f"配置参数:")
        self.log(f"  - 输入目录: {root_folder}")
        self.log(f"  - 分组大小: {group_size} 个文件")
        self.log(f"  - 输出目录: {output_folder}")
        self.log(f"  - 拆分逻辑: 按 BrowserID 聚合")
        self.log("=" * 80)

        # 1. 收集文件
        excel_files = self.collect_excels(root_folder)
        if not excel_files:
            self.log("❌ 未找到任何 Excel 文件，任务终止", "ERROR")
            return {'success': False, 'error': '未找到任何 Excel 文件'}

        # 2. 并行加载 Excel 文件
        self.log(f"\n开始加载 {len(excel_files)} 个 Excel 文件...")
        all_data = []
        success_count = 0
        failed_count = 0

        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(self.split_excel, f, group_size): f for f in excel_files}
            for future in tqdm(as_completed(futures), total=len(futures), desc="加载文件", unit="个"):
                file_path = futures[future]
                try:
                    df = future.result()
                    if df is not None:
                        all_data.append(df)
                        success_count += 1
                        self.log(f"✓ 成功加载: {os.path.basename(file_path)} ({len(df)} 行)")
                    else:
                        failed_count += 1
                except Exception as e:
                    failed_count += 1
                    self.log(f"✗ 加载失败: {os.path.basename(file_path)} - {str(e)}", "ERROR")

        self.log(f"\n文件加载完成:")
        self.log(f"  - 成功: {success_count} 个")
        self.log(f"  - 失败/空文件: {failed_count} 个")
        
        if not all_data:
            self.log("❌ 没有有效的数据可以处理", "ERROR")
            return {'success': False, 'error': '没有有效的数据可以处理'}

        # 3. 按 BrowserID 重组并保存
        self.log("\n" + "=" * 80)
        result = self._regroup_and_save_by_browser_id(all_data, group_size, output_folder)
        
        return result
    
    def _regroup_and_save_by_browser_id(self, all_data: List[pd.DataFrame], 
                                       group_size: int, output_folder: str) -> Dict:
        """
        按 BrowserID 重新组合并保存为新的 Excel
        
        Args:
            all_data: 所有数据的DataFrame列表
            group_size: 分组大小
            output_folder: 输出文件夹路径
            
        Returns:
            处理结果字典
        """
        self.log(f"创建输出目录: {output_folder}")
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        
        # 合并所有数据
        if not all_data:
            self.log("没有有效的数据", "ERROR")
            return {'success': False, 'error': '没有有效的数据'}
        
        combined_df = pd.concat(all_data, ignore_index=True)
        self.log(f"合并后总数据行数: {len(combined_df)}")
        
        # 检查是否存在 BrowserID 列
        if 'BrowserID' not in combined_df.columns:
            self.log("数据中没有 BrowserID 列", "ERROR")
            return {'success': False, 'error': '数据中没有 BrowserID 列'}
        
        # 按 BrowserID 分组
        browser_groups = {}
        for browser_id, group_df in combined_df.groupby('BrowserID'):
            browser_groups[browser_id] = group_df.reset_index(drop=True)
        
        self.log(f"共发现 {len(browser_groups)} 个不同的 BrowserID:")
        for browser_id, group_df in browser_groups.items():
            self.log(f"  - BrowserID {browser_id}: {len(group_df)} 行数据")
        
        # 计算需要生成的文件数（取最大值）
        max_rows = max(len(df) for df in browser_groups.values())
        num_files = min(max_rows, group_size)
        
        self.log(f"将生成 {num_files} 个输出文件（分组大小: {group_size}）")
        
        total_output = 0
        failed_output = 0
        
        # 生成每个输出文件
        for file_idx in tqdm(range(num_files), desc="生成文件", unit="个"):
            try:
                rows_for_this_file = []
                
                # 收集每个 BrowserID 的第 file_idx 行数据
                for browser_id, group_df in browser_groups.items():
                    if file_idx < len(group_df):
                        # 直接使用原始数据，不对齐列结构
                        row = group_df.iloc[[file_idx]].copy()
                        rows_for_this_file.append(row)
                
                if not rows_for_this_file:
                    self.log(f"第 {file_idx + 1} 个文件没有数据，跳过", "WARNING")
                    continue
                
                # 合并所有行，保持各自的列结构和列顺序
                combined = pd.concat(rows_for_this_file, ignore_index=True, sort=False)
                
                # 保存文件
                out_name = f"output_{file_idx + 1:03d}.xlsx"
                out_path = os.path.join(output_folder, out_name)
                combined.to_excel(out_path, index=False)
                total_output += 1
                
                self.log(f"  ✓ 生成文件 {out_name}: {len(combined)} 行数据（{len(rows_for_this_file)} 个 BrowserID）")
                
            except Exception as e:
                failed_output += 1
                self.log(f"保存文件失败 (文件{file_idx+1}): {str(e)}", "ERROR")
        
        self.log(f"✅ 处理完成:")
        self.log(f"  - 成功输出: {total_output} 个文件")
        if failed_output > 0:
            self.log(f"  - 失败: {failed_output} 个文件", "WARNING")
        
        return {
            'success': True,
            'total_output': total_output,
            'failed_output': failed_output,
            'browser_id_count': len(browser_groups)
        }
    
    def merge_excel_files(self, excel_files: List[str], output_path: str) -> Dict:
        """
        合并多个 Excel 文件到一个文件中
        
        Args:
            excel_files: Excel文件路径列表
            output_path: 输出文件路径
            
        Returns:
            处理结果字典
        """
        self.log("开始合并 Excel 文件")
        
        if not excel_files:
            self.log("没有找到 Excel 文件", "ERROR")
            return {'success': False, 'error': '没有找到 Excel 文件'}
        
        try:
            all_data = []
            total_rows = 0
            
            # 读取所有文件
            for i, file_path in enumerate(excel_files, 1):
                self.log(f"正在读取文件 {i}/{len(excel_files)}: {os.path.basename(file_path)}")
                
                try:
                    df = pd.read_excel(file_path)
                    if not df.empty:
                        # 添加来源文件列
                        df['SourceFile'] = os.path.basename(file_path)
                        df['SourceFolder'] = os.path.basename(os.path.dirname(file_path))
                        all_data.append(df)
                        total_rows += len(df)
                        self.log(f"  ✓ 成功读取 {len(df)} 行数据")
                    else:
                        self.log(f"  ⚠️ 文件为空: {os.path.basename(file_path)}", "WARNING")
                except Exception as e:
                    self.log(f"  ✗ 读取失败: {os.path.basename(file_path)} - {str(e)}", "ERROR")
            
            if not all_data:
                self.log("没有有效的数据可以合并", "ERROR")
                return {'success': False, 'error': '没有有效的数据可以合并'}
            
            # 合并所有数据
            self.log(f"开始合并 {len(all_data)} 个文件的数据...")
            merged_df = pd.concat(all_data, ignore_index=True)
            
            # 重新排列列，将来源信息放在最后
            cols = list(merged_df.columns)
            source_cols = ['SourceFile', 'SourceFolder']
            other_cols = [col for col in cols if col not in source_cols]
            final_cols = other_cols + source_cols
            merged_df = merged_df[final_cols]
            
            # 保存合并后的文件
            self.log(f"正在保存合并文件: {output_path}")
            merged_df.to_excel(output_path, index=False)
            
            self.log(f"✅ 合并完成!")
            self.log(f"  - 合并文件数: {len(all_data)}")
            self.log(f"  - 总行数: {total_rows}")
            self.log(f"  - 输出文件: {output_path}")
            
            return {
                'success': True,
                'merged_files_count': len(all_data),
                'total_rows': total_rows,
                'output_file': output_path
            }
            
        except Exception as e:
            self.log(f"合并过程中出错: {str(e)}", "ERROR")
            return {'success': False, 'error': str(e)}

