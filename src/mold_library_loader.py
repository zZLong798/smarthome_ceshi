#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模具库加载器模块
负责加载和验证模具库Excel文件，提取产品ID信息，提供产品数据映射
"""

import os
import pandas as pd
from typing import Dict, List, Any, Optional, Tuple


class MoldLibraryLoader:
    """模具库加载器类"""
    
    def __init__(self):
        """初始化模具库加载器"""
        self.dataframe: Optional[pd.DataFrame] = None
        self.mold_info: Dict[str, Any] = {}
        
    def load_mold_library(self, excel_path: str) -> bool:
        """
        加载模具库Excel文件
        
        Args:
            excel_path: Excel文件路径
            
        Returns:
            bool: 是否加载成功
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(excel_path):
                print(f"❌ 模具库文件不存在: {excel_path}")
                return False
            
            # 检查文件格式
            if not excel_path.lower().endswith(('.xlsx', '.xls')):
                print(f"❌ 模具库文件格式不正确，必须是.xlsx或.xls格式: {excel_path}")
                return False
            
            # 加载Excel文件
            print(f"🔍 加载模具库文件: {excel_path}")
            self.dataframe = pd.read_excel(excel_path)
            
            # 分析模具库结构
            if not self._analyze_mold_library():
                print("❌ 模具库结构分析失败")
                return False
            
            print("✅ 模具库文件加载成功")
            return True
            
        except Exception as e:
            print(f"❌ 加载模具库文件失败: {e}")
            return False
    
    def _analyze_mold_library(self) -> bool:
        """
        分析模具库结构
        
        Returns:
            bool: 是否分析成功
        """
        try:
            if self.dataframe is None:
                return False
            
            # 获取模具库基本信息
            self.mold_info = {
                'row_count': len(self.dataframe),
                'column_count': len(self.dataframe.columns),
                'column_names': list(self.dataframe.columns),
                'product_ids': [],
                'device_categories': [],
                'brands': []
            }
            
            # 提取产品ID
            if '产品ID' in self.dataframe.columns:
                self.mold_info['product_ids'] = self.dataframe['产品ID'].dropna().unique().tolist()
            
            # 提取设备品类
            if '设备品类' in self.dataframe.columns:
                self.mold_info['device_categories'] = self.dataframe['设备品类'].dropna().unique().tolist()
            
            # 提取品牌
            if '品牌' in self.dataframe.columns:
                self.mold_info['brands'] = self.dataframe['品牌'].dropna().unique().tolist()
            
            print(f"📊 模具库结构分析完成:")
            print(f"   • 产品数量: {self.mold_info['row_count']}")
            print(f"   • 列数: {self.mold_info['column_count']}")
            print(f"   • 列名: {self.mold_info['column_names']}")
            print(f"   • 产品ID数量: {len(self.mold_info['product_ids'])}")
            print(f"   • 设备品类: {self.mold_info['device_categories']}")
            print(f"   • 品牌: {self.mold_info['brands']}")
            
            return True
            
        except Exception as e:
            print(f"❌ 分析模具库结构失败: {e}")
            return False
    
    def extract_product_ids(self) -> List[int]:
        """
        提取产品ID列表
        
        Returns:
            List[int]: 产品ID列表
        """
        if self.dataframe is None or '产品ID' not in self.dataframe.columns:
            return []
        
        try:
            product_ids = self.dataframe['产品ID'].dropna().astype(int).unique().tolist()
            print(f"📋 提取到 {len(product_ids)} 个产品ID: {product_ids}")
            return product_ids
        except Exception as e:
            print(f"❌ 提取产品ID失败: {e}")
            return []
    
    def get_product_info(self, product_id: int) -> Optional[Dict[str, Any]]:
        """
        根据产品ID获取产品信息
        
        Args:
            product_id: 产品ID
            
        Returns:
            Optional[Dict[str, Any]]: 产品信息字典，如果找不到返回None
        """
        if self.dataframe is None or '产品ID' not in self.dataframe.columns:
            return None
        
        try:
            # 查找匹配的产品ID
            product_row = self.dataframe[self.dataframe['产品ID'] == product_id]
            
            if len(product_row) == 0:
                print(f"⚠️ 未找到产品ID为 {product_id} 的产品")
                return None
            
            # 转换为字典格式
            product_info = product_row.iloc[0].to_dict()
            
            # 清理NaN值
            for key, value in product_info.items():
                if pd.isna(value):
                    product_info[key] = None
            
            print(f"🔍 找到产品ID {product_id} 的信息")
            return product_info
            
        except Exception as e:
            print(f"❌ 获取产品信息失败: {e}")
            return None
    
    def validate_mold_library(self) -> Tuple[bool, List[str]]:
        """
        验证模具库格式
        
        Returns:
            Tuple[bool, List[str]]: (是否验证通过, 错误信息列表)
        """
        errors = []
        
        if self.dataframe is None:
            errors.append("模具库未加载")
            return False, errors
        
        # 检查必要的列名
        required_columns = ['产品ID', '设备品类', '设备名称', '品牌', '单价']
        existing_columns = self.dataframe.columns.tolist()
        
        for required_col in required_columns:
            if required_col not in existing_columns:
                errors.append(f"缺少必要列: {required_col}")
        
        # 检查产品ID的唯一性
        if '产品ID' in self.dataframe.columns:
            duplicate_ids = self.dataframe[self.dataframe.duplicated('产品ID', keep=False)]
            if len(duplicate_ids) > 0:
                errors.append(f"存在重复的产品ID: {duplicate_ids['产品ID'].unique().tolist()}")
        
        # 检查数据完整性
        if len(self.dataframe) == 0:
            errors.append("模具库为空")
        
        if errors:
            return False, errors
        else:
            return True, []
    
    def get_mold_info(self) -> Dict[str, Any]:
        """
        获取模具库信息
        
        Returns:
            Dict[str, Any]: 模具库信息
        """
        return self.mold_info.copy()
    
    def get_dataframe(self) -> Optional[pd.DataFrame]:
        """
        获取数据框对象
        
        Returns:
            Optional[pd.DataFrame]: 数据框对象
        """
        return self.dataframe
    
    def search_products(self, keyword: str, search_columns: List[str] = None) -> pd.DataFrame:
        """
        根据关键词搜索产品
        
        Args:
            keyword: 搜索关键词
            search_columns: 搜索的列名列表，如果为None则搜索所有文本列
            
        Returns:
            pd.DataFrame: 搜索结果
        """
        if self.dataframe is None:
            return pd.DataFrame()
        
        try:
            if search_columns is None:
                # 默认搜索所有文本列
                text_columns = self.dataframe.select_dtypes(include=['object']).columns.tolist()
                search_columns = text_columns
            
            # 创建搜索条件
            search_condition = False
            for column in search_columns:
                if column in self.dataframe.columns:
                    search_condition = search_condition | self.dataframe[column].astype(str).str.contains(keyword, case=False, na=False)
            
            results = self.dataframe[search_condition]
            print(f"🔍 搜索关键词 '{keyword}' 找到 {len(results)} 个结果")
            return results
            
        except Exception as e:
            print(f"❌ 搜索产品失败: {e}")
            return pd.DataFrame()


def load_and_validate_mold_library(excel_path: str) -> Tuple[bool, Optional[MoldLibraryLoader], List[str]]:
    """
    加载并验证模具库文件的便捷函数
    
    Args:
        excel_path: Excel文件路径
        
    Returns:
        Tuple[bool, Optional[MoldLibraryLoader], List[str]]: (是否成功, 模具库加载器对象, 错误信息)
    """
    loader = MoldLibraryLoader()
    
    # 加载模具库
    if not loader.load_mold_library(excel_path):
        return False, None, ["模具库加载失败"]
    
    # 验证模具库
    is_valid, errors = loader.validate_mold_library()
    
    if not is_valid:
        return False, None, errors
    
    return True, loader, []


def main():
    """主函数 - 测试模具库加载器"""
    print("=" * 60)
    print("📋 模具库加载器测试")
    print("=" * 60)
    
    # 测试默认模具库
    mold_library_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '智能家居模具库.xlsx')
    
    if not os.path.exists(mold_library_path):
        print(f"❌ 模具库文件不存在: {mold_library_path}")
        return
    
    # 加载并验证模具库
    success, loader, errors = load_and_validate_mold_library(mold_library_path)
    
    if success:
        print("\n✅ 模具库验证通过")
        
        # 显示模具库信息
        mold_info = loader.get_mold_info()
        print(f"\n📊 模具库详细信息:")
        print(f"   • 产品数量: {mold_info['row_count']}")
        print(f"   • 列数: {mold_info['column_count']}")
        print(f"   • 产品ID数量: {len(mold_info['product_ids'])}")
        print(f"   • 设备品类: {mold_info['device_categories']}")
        print(f"   • 品牌: {mold_info['brands']}")
        
        # 提取产品ID
        product_ids = loader.extract_product_ids()
        print(f"\n📋 产品ID列表: {product_ids}")
        
        # 测试产品信息查询
        if product_ids:
            print(f"\n🔍 测试产品信息查询:")
            for product_id in product_ids[:3]:  # 测试前3个产品
                product_info = loader.get_product_info(product_id)
                if product_info:
                    print(f"   • 产品ID {product_id}: {product_info.get('设备名称', '未知')} - {product_info.get('品牌', '未知')}")
        
        # 测试搜索功能
        print(f"\n🔍 测试搜索功能:")
        search_results = loader.search_products("智能开关")
        if len(search_results) > 0:
            print(f"   找到 {len(search_results)} 个智能开关产品")
        
    else:
        print(f"\n❌ 模具库验证失败:")
        for error in errors:
            print(f"   • {error}")


if __name__ == "__main__":
    main()