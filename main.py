#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel价格批量更新工具
交互式获取多个Excel文件，通过正则匹配ProductNameCn更新价格
"""

import os
import json
import re
import glob
import random
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Union
import pandas as pd
import numpy as np
from openpyxl import load_workbook


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


class ExcelPriceUpdater:
    """Excel价格批量更新器"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        初始化更新器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.config = self._load_config()
        self.product_column = "ProductNameCn"
        # 自动检测所有地域并生成价格列名映射
        self.price_columns = self._build_price_columns()
    
    def _load_config(self) -> Dict:
        """
        加载配置文件，如果不存在则自动创建默认配置
        
        Returns:
            配置字典
            
        Raises:
            json.JSONDecodeError: 配置文件格式错误
        """
        if not os.path.exists(self.config_file):
            # 自动创建默认配置文件
            print(f"⚠️  配置文件 {self.config_file} 不存在，正在创建默认配置...")
            default_config = {
                "Nike Air Force 1": {
                    "hk": [550, 580, 10],
                    "sg": [70, 85, 5],
                    "my": [50, 60, 10]
                },
                "New Balance 327": {
                    "hk": [480, 510, 10],
                    "sg": [75, 90, 5],
                    "my": [60, 70, 10]
                }
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            
            print(f"✓ 已创建默认配置文件: {self.config_file}")
            print(f"💡 提示：可以使用配置编辑器修改配置: python open_config_editor.py\n")
            
            return default_config
        
        with open(self.config_file, 'r', encoding='utf-8') as f:
            try:
                config = json.load(f)
            except json.JSONDecodeError as e:
                raise json.JSONDecodeError(
                    f"配置文件格式错误：{e.msg}\n"
                    f"请检查 {self.config_file} 文件的JSON格式是否正确\n"
                    f"可以使用配置编辑器修复：python open_config_editor.py",
                    e.doc, e.pos
                )
        
        if not isinstance(config, dict):
            raise ValueError(
                f"配置文件格式错误：根对象必须是字典类型\n"
                f"当前类型：{type(config).__name__}\n"
                f"请使用配置编辑器修复：python open_config_editor.py"
            )
        
        if not config:
            # 配置文件为空时，自动填充默认配置
            print(f"⚠️  配置文件 {self.config_file} 为空，正在创建默认配置...")
            default_config = {
                "Nike Air Force 1": {
                    "hk": [550, 580, 10],
                    "sg": [70, 85, 5],
                    "my": [50, 60, 10]
                },
                "New Balance 327": {
                    "hk": [480, 510, 10],
                    "sg": [75, 90, 5],
                    "my": [60, 70, 10]
                }
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            
            print(f"✓ 已创建默认配置")
            print(f"💡 提示：可以使用配置编辑器修改配置: python open_config_editor.py\n")
            
            return default_config
        
        return config
    
    def _build_price_columns(self) -> Dict[str, str]:
        """
        自动检测配置文件中所有地域并生成价格列名映射
        
        Returns:
            地域代码到价格列名的映射字典
        """
        regions = set()
        
        # 收集所有产品中出现的地域
        for product_config in self.config.values():
            if isinstance(product_config, dict):
                regions.update(product_config.keys())
        
        # 生成价格列名映射：{region} -> {REGION}Price
        # 例如：hk -> HKPrice, sg -> SGPrice, tw -> TWPrice
        price_columns = {}
        for region in regions:
            # 将地区代码转换为首字母大写的列名
            price_columns[region.lower()] = f"{region.upper()}Price"
        
        return price_columns
    
    def _get_region_price(self, region: str) -> str:
        """
        获取地域对应的价格列名
        
        Args:
            region: 地域代码
            
        Returns:
            价格列名
        """
        return self.price_columns.get(region.lower())
    
    def _generate_random_price(self, price_config: Union[int, list]) -> int:
        """
        生成随机价格
        
        Args:
            price_config: 价格配置，可以是固定价格（int）或价格区间（list）
                        区间格式: [最小值, 最大值, 步长]
            
        Returns:
            生成的价格
        """
        # 如果是固定价格
        if isinstance(price_config, (int, float)):
            return int(price_config)
        
        # 如果是价格区间
        if isinstance(price_config, list):
            if len(price_config) != 3:
                raise ValueError(
                    f"价格区间配置格式错误：应为 [最小值, 最大值, 步长]，"
                    f"但得到 {price_config}"
                )
            
            min_price, max_price, step = price_config
            min_price = int(min_price)
            max_price = int(max_price)
            step = int(step)
            
            if min_price > max_price:
                raise ValueError(
                    f"价格区间配置错误：最小值 {min_price} 大于最大值 {max_price}"
                )
            
            if step <= 0:
                raise ValueError(
                    f"价格区间配置错误：步长 {step} 必须大于0"
                )
            
            # 验证最小值和步长的关系
            if min_price % step != 0:
                raise ValueError(
                    f"价格区间配置错误：最小值 {min_price} 必须是步长 {step} 的倍数"
                )
            
            # 计算可能的取值数量
            num_values = (max_price - min_price) // step + 1
            
            # 生成随机索引
            random_index = random.randint(0, num_values - 1)
            
            # 生成随机价格
            random_price = min_price + random_index * step
            
            return random_price
        
        raise ValueError(
            f"价格配置格式错误：应为固定价格（数字）或价格区间（[最小值, 最大值, 步长]），"
            f"但得到 {type(price_config)}: {price_config}"
        )
    
    def _match_price_key(self, product_name: str) -> Optional[str]:
        """
        通过正则匹配ProductNameCn找到对应的价格配置key
        优先匹配更具体（更长）的关键字
        
        Args:
            product_name: 产品名称
            
        Returns:
            匹配到的配置key，如果未匹配到返回None
        """
        if not product_name or pd.isna(product_name):
            return None
        
        product_name_str = str(product_name)
        
        # 按关键字长度降序排序，优先匹配更具体（更长）的关键字
        # 这样"samba a"会优先于"samba"匹配
        sorted_keys = sorted(self.config.keys(), key=len, reverse=True)
        
        # 遍历配置文件中的所有key，尝试匹配
        for key in sorted_keys:
            # 使用正则匹配，支持大小写不敏感
            pattern = re.compile(key, re.IGNORECASE)
            if pattern.search(product_name_str):
                return key
        
        return None
    
    def _validate_config(self, regions: List[str]) -> None:
        """
        验证配置文件是否包含所需地域的价格配置
        
        Args:
            regions: 需要更新的地域列表
            
        Raises:
            ValueError: 配置不完整
        """
        for product_key in self.config.keys():
            product_config = self.config[product_key]
            if not isinstance(product_config, dict):
                raise ValueError(
                    f"配置错误：产品 '{product_key}' 的价格配置必须是字典类型"
                )
            
            missing_regions = []
            for region in regions:
                if region not in product_config:
                    missing_regions.append(region)
            
            if missing_regions:
                raise ValueError(
                    f"产品 '{product_key}' 缺少以下地域的价格配置: {', '.join(missing_regions)}"
                )
            
            # 验证价格配置格式
            for region in regions:
                price_config = product_config[region]
                # 尝试生成价格以验证配置格式
                try:
                    self._generate_random_price(price_config)
                except ValueError as e:
                    raise ValueError(
                        f"产品 '{product_key}' 的地域 '{region}' 价格配置错误: {e}"
                    )
    
    def update_prices(self, excel_file: str, regions: List[str], 
                     output_suffix: str = "_updated") -> bool:
        """
        更新Excel文件中的价格
        
        Args:
            excel_file: Excel文件路径
            regions: 需要更新的地域列表
            output_suffix: 输出文件后缀
            
        Returns:
            是否成功更新
            
        Raises:
            FileNotFoundError: 文件不存在
            KeyError: 必需的列不存在
            ValueError: 配置错误或匹配失败
        """
        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"文件不存在: {excel_file}")
        
        # 读取Excel文件
        print(f"\n正在处理: {excel_file}")
        df = pd.read_excel(excel_file)
        
        # 检查必需的列是否存在
        if self.product_column not in df.columns:
            raise KeyError(
                f"Excel文件缺少必需的列: {self.product_column}"
            )
        
        # 检查价格列是否存在
        missing_price_columns = []
        for region in regions:
            price_col = self._get_region_price(region)
            if price_col not in df.columns:
                missing_price_columns.append(price_col)
        
        if missing_price_columns:
            raise KeyError(
                f"Excel文件缺少必需的价格列: {', '.join(missing_price_columns)}"
            )
        
        # 统计信息
        updated_count = 0
        not_found_products = []
        
        # 遍历每一行，更新价格
        for idx, row in df.iterrows():
            product_name = row[self.product_column]
            matched_key = self._match_price_key(product_name)
            
            if matched_key:
                # 找到匹配的配置，更新价格
                for region in regions:
                    price_col = self._get_region_price(region)
                    price_config = self.config[matched_key][region]
                    # 生成随机价格（如果配置是区间）或使用固定价格
                    price = self._generate_random_price(price_config)
                    df.at[idx, price_col] = price
                updated_count += 1
            else:
                # 记录未匹配到的产品
                not_found_products.append(str(product_name))
        
        # 如果有没有匹配到的产品，报错
        if not_found_products:
            unique_not_found = list(set(not_found_products))
            raise ValueError(
                f"无法匹配以下产品的价格配置:\n" +
                "\n".join(f"  - {product}" for product in unique_not_found[:10]) +
                (f"\n  ... 还有 {len(unique_not_found) - 10} 个产品未显示" 
                 if len(unique_not_found) > 10 else "") +
                f"\n\n请检查配置文件，补充这些产品的价格配置。"
            )
        
        # 保存更新后的文件
        output_file = self._get_output_filename(excel_file, output_suffix)
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"✓ 成功更新 {updated_count} 条记录")
        print(f"✓ 已保存到: {output_file}")
        
        return True
    
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
    
    def get_excel_files_interactive(self) -> List[str]:
        """
        交互式获取Excel文件列表
        
        Returns:
            Excel文件路径列表
        """
        print("\n" + "="*60)
        print("Excel价格批量更新工具")
        print("="*60)
        
        files = []
        
        while True:
            file_path = input("\n请输入Excel文件路径（直接回车结束输入）: ").strip()
            
            if not file_path:
                break
            
            # 支持通配符
            if '*' in file_path or '?' in file_path:
                matched_files = glob.glob(file_path)
                if matched_files:
                    files.extend(matched_files)
                    print(f"找到 {len(matched_files)} 个文件")
                else:
                    print(f"未找到匹配的文件: {file_path}")
            else:
                if os.path.exists(file_path):
                    if os.path.isfile(file_path):
                        files.append(file_path)
                    else:
                        print(f"不是文件: {file_path}")
                else:
                    print(f"文件不存在: {file_path}")
        
        if not files:
            raise ValueError("未选择任何文件")
        
        print(f"\n总共选择了 {len(files)} 个文件:")
        for i, file in enumerate(files, 1):
            print(f"  {i}. {file}")
        
        return files
    
    def get_regions_interactive(self) -> List[str]:
        """
        交互式获取需要更新的地域列表
        
        Returns:
            地域代码列表
        """
        print("\n可用地域:")
        for region, column in self.price_columns.items():
            print(f"  {region.upper():4s} -> {column}")
        
        print("\n请输入需要更新的地域（多个用逗号分隔，如: hk,sg,my）:")
        regions_input = input("地域代码: ").strip().lower()
        
        if not regions_input:
            raise ValueError("未选择任何地域")
        
        regions = [r.strip() for r in regions_input.split(',')]
        
        # 验证地域代码
        invalid_regions = [r for r in regions if r not in self.price_columns]
        if invalid_regions:
            raise ValueError(
                f"无效的地域代码: {', '.join(invalid_regions)}"
            )
        
        return regions


def main():
    """主函数"""
    try:
        # 初始化更新器
        updater = ExcelPriceUpdater()
        
        # 交互式获取文件
        excel_files = updater.get_excel_files_interactive()
        
        # 交互式获取地域
        regions = updater.get_regions_interactive()
        
        # 验证配置
        print("\n正在验证配置文件...")
        updater._validate_config(regions)
        print("✓ 配置文件验证通过")
        
        # 批量处理文件
        print("\n" + "="*60)
        print("开始处理文件...")
        print("="*60)
        
        success_count = 0
        fail_count = 0
        
        for excel_file in excel_files:
            try:
                updater.update_prices(excel_file, regions)
                success_count += 1
            except (FileNotFoundError, KeyError, ValueError) as e:
                print(f"\n✗ 处理失败: {excel_file}")
                print(f"  错误: {e}")
                fail_count += 1
        
        # 输出统计信息
        print("\n" + "="*60)
        print("处理完成!")
        print("="*60)
        print(f"成功: {success_count} 个文件")
        print(f"失败: {fail_count} 个文件")
        
    except (FileNotFoundError, ValueError, KeyError) as e:
        print(f"\n✗ 错误: {e}")
        return 1
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
        return 1
    except Exception as e:
        print(f"\n✗ 未预期的错误: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
