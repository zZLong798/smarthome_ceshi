#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查模具库中的产品ID格式
"""

import pandas as pd

def check_product_ids():
    """检查模具库中的产品ID格式"""
    
    excel_path = 'E:\\Programs\\smarthome\\智能家居模具库.xlsx'
    
    try:
        df = pd.read_excel(excel_path)
        print('📋 模具库产品ID列表:')
        print('='*60)
        
        for _, row in df.iterrows():
            product_id = row['产品ID']
            device_name = row['设备名称']
            brand = row['品牌']
            
            print(f'产品ID: {product_id}')
            print(f'设备名称: {device_name}')
            print(f'品牌: {brand}')
            print('-'*40)
        
        print(f'\n📊 总计: {len(df)} 个产品')
        
    except Exception as e:
        print(f'❌ 读取Excel文件失败: {e}')

def main():
    """主函数"""
    check_product_ids()

if __name__ == "__main__":
    main()