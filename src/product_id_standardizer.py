#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
产品ID标准化模块 - 任务1：创建产品ID标准化模块
将Excel模具库中的产品ID重新编号为纯数字递增格式
"""

import pandas as pd
import os
import json
from typing import Dict, List, Tuple

class ProductIDStandardizer:
    """产品ID标准化器"""
    
    def __init__(self, excel_path: str):
        """
        初始化标准化器
        
        Args:
            excel_path: Excel文件路径
        """
        self.excel_path = excel_path
        self.df = None
        self.mapping = {}
        
    def load_excel_data(self) -> bool:
        """
        加载Excel数据
        
        Returns:
            bool: 是否成功加载
        """
        try:
            self.df = pd.read_excel(self.excel_path)
            print(f"✅ 成功加载Excel文件: {self.excel_path}")
            print(f"📊 数据形状: {self.df.shape}")
            print(f"📋 列名: {list(self.df.columns)}")
            return True
        except Exception as e:
            print(f"❌ 加载Excel文件失败: {e}")
            return False
    
    def get_current_product_ids(self) -> List[str]:
        """
        获取当前产品ID列表
        
        Returns:
            List[str]: 产品ID列表
        """
        if self.df is None:
            return []
        
        if '产品ID' not in self.df.columns:
            print("❌ Excel文件中没有'产品ID'列")
            return []
        
        product_ids = self.df['产品ID'].tolist()
        print(f"📋 当前产品ID列表: {product_ids}")
        return product_ids
    
    def generate_new_ids(self, product_ids: List[str]) -> Dict[str, int]:
        """
        生成新的产品ID映射
        
        Args:
            product_ids: 原始产品ID列表
            
        Returns:
            Dict[str, int]: 原ID到新ID的映射
        """
        mapping = {}
        for i, old_id in enumerate(product_ids, 1):
            mapping[old_id] = i
        
        print(f"🔄 生成产品ID映射:")
        for old_id, new_id in mapping.items():
            print(f"   {old_id} -> {new_id}")
        
        return mapping
    
    def apply_standardization(self) -> bool:
        """
        应用产品ID标准化
        
        Returns:
            bool: 是否成功应用
        """
        if self.df is None:
            print("❌ 请先加载Excel数据")
            return False
        
        # 获取当前产品ID
        current_ids = self.get_current_product_ids()
        if not current_ids:
            return False
        
        # 生成新的ID映射
        self.mapping = self.generate_new_ids(current_ids)
        
        # 应用新的产品ID
        self.df['产品ID'] = self.df['产品ID'].map(self.mapping)
        
        print("✅ 产品ID标准化完成")
        print(f"📊 标准化后数据:")
        print(self.df[['产品ID', '设备名称', '品牌']].to_string(index=False))
        
        return True
    
    def save_standardized_excel(self, output_path: str = None) -> bool:
        """
        保存标准化后的Excel文件
        
        Args:
            output_path: 输出文件路径，如果为None则覆盖原文件
            
        Returns:
            bool: 是否成功保存
        """
        if self.df is None:
            print("❌ 没有数据可保存")
            return False
        
        if output_path is None:
            output_path = self.excel_path
        
        try:
            self.df.to_excel(output_path, index=False)
            print(f"✅ 标准化Excel文件已保存: {output_path}")
            return True
        except Exception as e:
            print(f"❌ 保存Excel文件失败: {e}")
            return False
    
    def save_mapping_table(self, mapping_path: str) -> bool:
        """
        保存产品ID映射表
        
        Args:
            mapping_path: 映射表文件路径
            
        Returns:
            bool: 是否成功保存
        """
        if not self.mapping:
            print("❌ 没有映射数据可保存")
            return False
        
        try:
            # 保存为JSON格式
            with open(mapping_path, 'w', encoding='utf-8') as f:
                json.dump(self.mapping, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 产品ID映射表已保存: {mapping_path}")
            return True
        except Exception as e:
            print(f"❌ 保存映射表失败: {e}")
            return False

def standardize_product_ids(excel_path: str, output_path: str = None, mapping_path: str = None) -> Dict[str, int]:
    """
    标准化产品ID的主函数
    
    Args:
        excel_path: 输入Excel文件路径
        output_path: 输出Excel文件路径，如果为None则覆盖原文件
        mapping_path: 映射表文件路径
        
    Returns:
        Dict[str, int]: 产品ID映射表
    """
    print("=" * 60)
    print("🔧 开始产品ID标准化 - 任务1")
    print("=" * 60)
    
    # 初始化标准化器
    standardizer = ProductIDStandardizer(excel_path)
    
    # 加载数据
    if not standardizer.load_excel_data():
        return {}
    
    # 应用标准化
    if not standardizer.apply_standardization():
        return {}
    
    # 保存结果
    if not standardizer.save_standardized_excel(output_path):
        return {}
    
    # 保存映射表
    if mapping_path and not standardizer.save_mapping_table(mapping_path):
        return {}
    
    print("=" * 60)
    print("✅ 产品ID标准化任务完成")
    print("=" * 60)
    
    return standardizer.mapping

if __name__ == "__main__":
    # 测试函数
    excel_path = "E:\\Programs\\smarthome\\智能家居模具库.xlsx"
    mapping = standardize_product_ids(excel_path)
    
    if mapping:
        print("🎯 标准化结果:")
        for old_id, new_id in mapping.items():
            print(f"   {old_id} -> {new_id}")
    else:
        print("❌ 标准化失败")