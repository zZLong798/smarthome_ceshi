#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设备信息查询模块 - 任务7：开发设备信息查询模块
根据pdid值从智能家居模具库中查询设备信息
"""

import pandas as pd
import os
from typing import Dict, List, Optional, Any

class DeviceInfoQuery:
    """设备信息查询器"""
    
    def __init__(self, excel_path: str = None):
        """
        初始化设备信息查询器
        
        Args:
            excel_path: Excel文件路径（智能家居模具库）
        """
        if excel_path is None:
            # 默认使用项目根目录下的模具库文件
            excel_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '智能家居模具库.xlsx')
        
        self.excel_path = excel_path
        self.product_df = None
        self.loaded = False
    
    def load_product_library(self) -> bool:
        """
        加载智能家居模具库
        
        Returns:
            bool: 是否成功加载
        """
        if not os.path.exists(self.excel_path):
            print(f"❌ 模具库文件不存在: {self.excel_path}")
            return False
        
        try:
            self.product_df = pd.read_excel(self.excel_path)
            print(f"✅ 成功加载模具库，共 {len(self.product_df)} 个产品")
            
            # 检查必要的列是否存在
            required_columns = ['产品ID', '品牌', '设备名称', '主规格']
            missing_columns = [col for col in required_columns if col not in self.product_df.columns]
            
            if missing_columns:
                print(f"⚠️ 模具库缺少必要的列: {missing_columns}")
                return False
            
            self.loaded = True
            return True
            
        except Exception as e:
            print(f"❌ 加载模具库失败: {e}")
            return False
    
    def query_device_by_pdid(self, pdid: int) -> Optional[Dict[str, Any]]:
        """
        根据pdid查询设备信息
        
        Args:
            pdid: 产品ID值
            
        Returns:
            Optional[Dict]: 设备信息字典，未找到返回None
        """
        if not self.loaded:
            print("❌ 请先加载模具库")
            return None
        
        try:
            # 在模具库中查找产品信息
            product_info = self.product_df[self.product_df['产品ID'] == pdid]
            
            if product_info.empty:
                print(f"⚠️ 未找到产品ID {pdid} 对应的设备信息")
                return None
            
            # 获取第一条匹配的记录
            device_info = product_info.iloc[0].to_dict()
            
            # 格式化返回信息
            result = {
                'pdid': pdid,
                'brand': device_info.get('品牌', ''),
                'device_name': device_info.get('设备名称', ''),
                'specification': device_info.get('主规格', ''),
                'model': device_info.get('型号', ''),
                'price': device_info.get('价格', ''),
                'supplier': device_info.get('供应商', ''),
                'notes': device_info.get('备注', '')
            }
            
            print(f"✅ 找到产品ID {pdid} 的设备: {result['brand']} {result['device_name']}")
            return result
            
        except Exception as e:
            print(f"❌ 查询产品ID {pdid} 失败: {e}")
            return None
    
    def query_devices_by_pdid_list(self, pdid_list: List[int]) -> Dict[int, Dict[str, Any]]:
        """
        批量查询设备信息
        
        Args:
            pdid_list: pdid值列表
            
        Returns:
            Dict[int, Dict]: pdid到设备信息的映射
        """
        if not self.loaded:
            print("❌ 请先加载模具库")
            return {}
        
        device_mapping = {}
        
        print(f"\n🔍 开始批量查询 {len(pdid_list)} 个pdid的设备信息...")
        
        for pdid in pdid_list:
            device_info = self.query_device_by_pdid(pdid)
            if device_info:
                device_mapping[pdid] = device_info
        
        print(f"📊 成功查询到 {len(device_mapping)} 个设备的详细信息")
        return device_mapping
    
    def get_all_products(self) -> List[Dict[str, Any]]:
        """
        获取模具库中所有产品信息
        
        Returns:
            List[Dict]: 所有产品信息列表
        """
        if not self.loaded:
            print("❌ 请先加载模具库")
            return []
        
        products = []
        for _, row in self.product_df.iterrows():
            product_info = {
                'pdid': row['产品ID'],
                'brand': row.get('品牌', ''),
                'device_name': row.get('设备名称', ''),
                'specification': row.get('主规格', ''),
                'model': row.get('型号', ''),
                'price': row.get('价格', ''),
                'supplier': row.get('供应商', ''),
                'notes': row.get('备注', '')
            }
            products.append(product_info)
        
        return products
    
    def save_device_query_report(self, device_mapping: Dict[int, Dict], output_path: str = "device_query_report.json") -> bool:
        """
        保存设备查询报告
        
        Args:
            device_mapping: 设备信息映射
            output_path: 输出文件路径
            
        Returns:
            bool: 是否成功保存
        """
        try:
            report = {
                'query_time': pd.Timestamp.now().isoformat(),
                'total_queried_pdids': len(device_mapping),
                'devices': device_mapping
            }
            
            import json
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2)
            
            print(f"\n💾 设备查询报告已保存至: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ 保存设备查询报告失败: {e}")
            return False


def test_device_info_query():
    """测试设备信息查询模块"""
    
    print("🔧 设备信息查询模块测试")
    print("="*60)
    
    # 创建查询器
    query = DeviceInfoQuery()
    
    # 加载模具库
    if not query.load_product_library():
        print("❌ 模具库加载失败，测试终止")
        return
    
    # 测试单个查询
    print("\n🧪 测试单个设备查询:")
    device_info = query.query_device_by_pdid(1)
    if device_info:
        print(f"   ✅ 查询结果: {device_info}")
    
    # 测试批量查询
    print("\n🧪 测试批量设备查询:")
    pdid_list = [1, 2, 3, 4, 5, 6, 7, 8]
    device_mapping = query.query_devices_by_pdid_list(pdid_list)
    
    # 显示查询结果
    print("\n📋 查询结果汇总:")
    for pdid, info in device_mapping.items():
        print(f"   🏷️ PDID {pdid}: {info['brand']} {info['device_name']} - {info['specification']}")
    
    # 保存查询报告
    query.save_device_query_report(device_mapping)
    
    print("\n" + "="*60)
    print("✅ 设备信息查询模块测试完成")


if __name__ == "__main__":
    test_device_info_query()