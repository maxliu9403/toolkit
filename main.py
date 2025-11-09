#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel价格批量更新工具 - 命令行版本
交互式获取多个Excel文件，通过正则匹配ProductNameCn更新价格
"""

from modules import BrowserIDReplacer, ExcelPriceUpdater


def main():
    """主函数"""
    try:
        # 初始化更新器（使用新的模块）
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
