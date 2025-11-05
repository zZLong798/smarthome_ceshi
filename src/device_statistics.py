#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设备统计模块 - 任务8：开发设备统计模块
统计从PPT中提取的pdid标签对应的设备数量和分类信息
"""

import json
from typing import Dict, List, Any
from collections import defaultdict

class DeviceStatistics:
    """设备统计器"""
    
    def __init__(self):
        """初始化设备统计器"""
        self.device_counts = defaultdict(int)
        self.brand_stats = defaultdict(list)
        self.category_stats = defaultdict(list)
        self.total_devices = 0
        
    def load_pdid_data(self, pdid_extraction_report_path: str = "pdid_extraction_report.json") -> Dict[int, int]:
        """
        加载pdid提取数据
        
        Args:
            pdid_extraction_report_path: pdid提取报告文件路径
            
        Returns:
            Dict[int, int]: pdid到数量的映射
        """
        try:
            with open(pdid_extraction_report_path, 'r', encoding='utf-8') as f:
                pdid_data = json.load(f)
            
            # 从报告中提取pdid统计信息
            pdid_counts = defaultdict(int)
            
            # 遍历所有幻灯片中的pdid标签
            for slide_data in pdid_data.get('pdid_labels', {}).values():
                for label in slide_data:
                    pdid = label.get('pdid')
                    if pdid is not None:
                        pdid_counts[pdid] += 1
            
            print(f"✅ 成功加载pdid提取数据，共发现 {len(pdid_counts)} 种pdid")
            return dict(pdid_counts)
            
        except Exception as e:
            print(f"❌ 加载pdid提取数据失败: {e}")
            return {}
    
    def count_devices_by_pdid(self, pdid_counts: Dict[int, int], device_mapping: Dict[int, Dict]) -> Dict[str, Any]:
        """
        根据pdid统计设备数量和分类
        
        Args:
            pdid_counts: pdid到数量的映射
            device_mapping: pdid到设备信息的映射
            
        Returns:
            Dict[str, Any]: 统计结果
        """
        # 重置统计结果
        self.device_counts.clear()
        self.brand_stats.clear()
        self.category_stats.clear()
        self.total_devices = 0
        
        # 统计设备数量和分类
        for pdid, count in pdid_counts.items():
            # 处理PDID类型不匹配问题：pdid_counts中是整数，device_mapping中是字符串
            pdid_key = str(pdid)
            if pdid_key in device_mapping:
                device_info = device_mapping[pdid_key]
                
                # 统计设备数量
                self.device_counts[pdid] = count
                self.total_devices += count
                
                # 按品牌统计
                brand = device_info.get('brand', '未知品牌')
                self.brand_stats[brand].append({
                    'pdid': pdid,
                    'device_name': device_info.get('device_name', ''),
                    'specification': device_info.get('specification', ''),
                    'count': count
                })
                
                # 按设备类型统计（从设备名称中提取类型）
                device_name = device_info.get('device_name', '')
                category = self._extract_device_category(device_name)
                self.category_stats[category].append({
                    'pdid': pdid,
                    'device_name': device_name,
                    'brand': brand,
                    'specification': device_info.get('specification', ''),
                    'count': count
                })
        
        return {
            'total_devices': self.total_devices,
            'unique_pdids': len(self.device_counts),
            'brands': len(self.brand_stats),
            'categories': len(self.category_stats),
            'device_counts': dict(self.device_counts),
            'brand_stats': dict(self.brand_stats),
            'category_stats': dict(self.category_stats)
        }
    
    def _extract_device_category(self, device_name: str) -> str:
        """
        从设备名称中提取设备类型
        
        Args:
            device_name: 设备名称
            
        Returns:
            str: 设备类型
        """
        if '开关' in device_name:
            return '智能开关'
        elif '插座' in device_name:
            return '智能插座'
        elif '传感器' in device_name:
            return '传感器'
        elif '网关' in device_name:
            return '网关'
        elif '面板' in device_name:
            return '控制面板'
        else:
            return '其他设备'
    
    def generate_statistics_report(self, statistics: Dict[str, Any]) -> str:
        """
        生成统计报告
        
        Args:
            statistics: 统计结果
            
        Returns:
            str: 统计报告文本
        """
        report = []
        report.append("📊 设备统计报告")
        report.append("=" * 60)
        
        # 总体统计
        report.append(f"📈 总体统计:")
        report.append(f"   • 设备总数: {statistics['total_devices']} 个")
        report.append(f"   • 设备种类: {statistics['unique_pdids']} 种")
        report.append(f"   • 品牌数量: {statistics['brands']} 个")
        report.append(f"   • 设备分类: {statistics['categories']} 类")
        
        # 按品牌统计
        if statistics['brand_stats']:
            report.append(f"\n🏷️ 按品牌统计:")
            for brand, devices in statistics['brand_stats'].items():
                brand_total = sum(device['count'] for device in devices)
                report.append(f"   • {brand}: {brand_total} 个设备")
                
                for device in devices:
                    report.append(f"      - {device['device_name']}: {device['count']} 个")
        
        # 按分类统计
        if statistics['category_stats']:
            report.append(f"\n🔧 按设备分类统计:")
            for category, devices in statistics['category_stats'].items():
                category_total = sum(device['count'] for device in devices)
                report.append(f"   • {category}: {category_total} 个设备")
                
                for device in devices:
                    report.append(f"      - {device['brand']} {device['device_name']}: {device['count']} 个")
        
        # 详细设备统计
        if statistics['device_counts']:
            report.append(f"\n📋 详细设备统计:")
            for pdid, count in statistics['device_counts'].items():
                report.append(f"   • PDID {pdid}: {count} 个")
        
        return '\n'.join(report)
    
    def save_statistics_report(self, statistics: Dict[str, Any], output_path: str = "device_statistics_report.json") -> bool:
        """
        保存统计报告
        
        Args:
            statistics: 统计结果
            output_path: 输出文件路径
            
        Returns:
            bool: 是否成功保存
        """
        try:
            report = {
                'statistics_time': statistics.get('statistics_time', ''),
                'total_devices': statistics['total_devices'],
                'unique_pdids': statistics['unique_pdids'],
                'brands': statistics['brands'],
                'categories': statistics['categories'],
                'device_counts': statistics['device_counts'],
                'brand_stats': statistics['brand_stats'],
                'category_stats': statistics['category_stats']
            }
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2)
            
            print(f"💾 设备统计报告已保存至: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ 保存设备统计报告失败: {e}")
            return False


def test_device_statistics():
    """测试设备统计模块"""
    
    print("🔧 设备统计模块测试")
    print("=" * 60)
    
    # 创建统计器
    stats = DeviceStatistics()
    
    # 从实际pdid提取报告中获取数据
    pdid_counts = stats.load_pdid_data("pdid_extraction_report.json")
    
    if not pdid_counts:
        print("❌ 无法加载pdid提取数据，测试终止")
        return
    
    print(f"📊 实际发现的PDID标签: {pdid_counts}")
    
    # 从设备查询报告中获取设备映射数据
    try:
        with open("device_query_report.json", 'r', encoding='utf-8') as f:
            device_query_data = json.load(f)
        device_mapping = device_query_data.get('devices', {})
        print(f"📋 可查询的设备数量: {len(device_mapping)}")
    except Exception as e:
        print(f"❌ 无法加载设备查询数据: {e}")
        return
    
    # 统计设备
    statistics = stats.count_devices_by_pdid(pdid_counts, device_mapping)
    statistics['statistics_time'] = '2025-10-31T01:30:00'
    
    # 生成报告
    report_text = stats.generate_statistics_report(statistics)
    print(report_text)
    
    # 保存报告
    stats.save_statistics_report(statistics)
    
    print("\n" + "=" * 60)
    print("✅ 设备统计模块测试完成")


if __name__ == "__main__":
    test_device_statistics()