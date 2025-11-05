#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简要报告生成模块 - 任务4：开发简要报告生成模块
生成简要设备清单报告，包含品牌、主规格、设备名称和数量信息
"""

import json
from typing import Dict, List, Any
from datetime import datetime


class BriefReportGenerator:
    """简要报告生成器"""
    
    def __init__(self):
        """初始化报告生成器"""
        self.report_data = {}
        
    def load_statistics_data(self, statistics_report_path: str = "device_statistics_report.json") -> Dict[str, Any]:
        """
        加载设备统计数据
        
        Args:
            statistics_report_path: 设备统计报告文件路径
            
        Returns:
            Dict[str, Any]: 设备统计数据
        """
        try:
            with open(statistics_report_path, 'r', encoding='utf-8') as f:
                statistics_data = json.load(f)
            
            print(f"✅ 成功加载设备统计数据")
            return statistics_data
            
        except Exception as e:
            print(f"❌ 加载设备统计数据失败: {e}")
            return {}
    
    def generate_brief_report(self, statistics_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        生成简要设备清单报告
        
        Args:
            statistics_data: 设备统计数据
            
        Returns:
            Dict[str, Any]: 简要报告数据
        """
        # 提取关键信息
        brand_stats = statistics_data.get('brand_stats', {})
        device_counts = statistics_data.get('device_counts', {})
        
        # 生成简要设备清单
        device_list = []
        
        # 按品牌和设备名称组织数据
        for brand, devices in brand_stats.items():
            for device_info in devices:
                device_list.append({
                    'brand': brand,
                    'device_name': device_info.get('device_name', ''),
                    'specification': device_info.get('specification', ''),
                    'count': device_info.get('count', 0)
                })
        
        # 按品牌排序
        device_list.sort(key=lambda x: x['brand'])
        
        # 生成简要报告
        brief_report = {
            'report_time': datetime.now().isoformat(),
            'total_devices': statistics_data.get('total_devices', 0),
            'total_brands': len(brand_stats),
            'device_list': device_list,
            'summary': {
                'brands': list(brand_stats.keys()),
                'device_types': list(set([device['device_name'] for device in device_list]))
            }
        }
        
        return brief_report
    
    def generate_console_output(self, brief_report: Dict[str, Any]) -> str:
        """
        生成控制台输出格式
        
        Args:
            brief_report: 简要报告数据
            
        Returns:
            str: 控制台输出文本
        """
        output = []
        output.append("📋 简要设备清单报告")
        output.append("=" * 60)
        
        # 总体统计
        output.append(f"📈 总体统计:")
        output.append(f"   • 设备总数: {brief_report['total_devices']} 个")
        output.append(f"   • 品牌数量: {brief_report['total_brands']} 个")
        output.append(f"   • 设备种类: {len(brief_report['device_list'])} 种")
        
        # 设备清单
        output.append(f"\n📋 设备清单:")
        
        current_brand = ""
        for device in brief_report['device_list']:
            if device['brand'] != current_brand:
                current_brand = device['brand']
                output.append(f"\n🏷️  {current_brand}:")
            
            output.append(f"   📱 {device['device_name']}")
            output.append(f"      • 规格: {device['specification']}")
            output.append(f"      • 数量: {device['count']} 个")
        
        # 汇总信息
        output.append(f"\n📊 汇总信息:")
        output.append(f"   • 品牌列表: {', '.join(brief_report['summary']['brands'])}")
        output.append(f"   • 设备类型: {', '.join(brief_report['summary']['device_types'])}")
        
        return '\n'.join(output)
    
    def save_brief_report(self, brief_report: Dict[str, Any], output_path: str = "brief_device_report.json") -> bool:
        """
        保存简要报告
        
        Args:
            brief_report: 简要报告数据
            output_path: 输出文件路径
            
        Returns:
            bool: 是否成功保存
        """
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(brief_report, f, ensure_ascii=False, indent=2)
            
            print(f"💾 简要设备清单报告已保存至: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ 保存简要设备清单报告失败: {e}")
            return False
    
    def save_text_report(self, console_output: str, output_path: str = "brief_device_report.txt") -> bool:
        """
        保存文本格式报告
        
        Args:
            console_output: 控制台输出文本
            output_path: 输出文件路径
            
        Returns:
            bool: 是否成功保存
        """
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(console_output)
            
            print(f"💾 文本格式报告已保存至: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ 保存文本格式报告失败: {e}")
            return False


def test_brief_report_generator():
    """测试简要报告生成模块"""
    
    print("🔧 简要报告生成模块测试")
    print("=" * 60)
    
    # 创建报告生成器
    generator = BriefReportGenerator()
    
    # 模拟设备统计数据
    statistics_data = {
        'statistics_time': '2025-10-31T01:30:00',
        'total_devices': 21,
        'unique_pdids': 8,
        'brands': 2,
        'categories': 1,
        'device_counts': {1: 5, 2: 3, 3: 2, 4: 1, 5: 4, 6: 2, 7: 3, 8: 1},
        'brand_stats': {
            '领普': [
                {'pdid': 1, 'device_name': '一键智能开关', 'specification': '白色四开', 'count': 5},
                {'pdid': 2, 'device_name': '二键智能开关', 'specification': '白色四开', 'count': 3},
                {'pdid': 3, 'device_name': '三键智能开关', 'specification': '白色四开', 'count': 2},
                {'pdid': 4, 'device_name': '四键智能开关', 'specification': '白色四开', 'count': 1}
            ],
            '易来': [
                {'pdid': 5, 'device_name': '一键智能开关', 'specification': '灰色', 'count': 4},
                {'pdid': 6, 'device_name': '二键智能开关', 'specification': '灰色', 'count': 2},
                {'pdid': 7, 'device_name': '三键智能开关', 'specification': '灰色', 'count': 3},
                {'pdid': 8, 'device_name': '四键智能开关', 'specification': '灰色', 'count': 1}
            ]
        },
        'category_stats': {
            '智能开关': [
                {'pdid': 1, 'device_name': '一键智能开关', 'brand': '领普', 'specification': '白色四开', 'count': 5},
                {'pdid': 2, 'device_name': '二键智能开关', 'brand': '领普', 'specification': '白色四开', 'count': 3},
                {'pdid': 3, 'device_name': '三键智能开关', 'brand': '领普', 'specification': '白色四开', 'count': 2},
                {'pdid': 4, 'device_name': '四键智能开关', 'brand': '领普', 'specification': '白色四开', 'count': 1},
                {'pdid': 5, 'device_name': '一键智能开关', 'brand': '易来', 'specification': '灰色', 'count': 4},
                {'pdid': 6, 'device_name': '二键智能开关', 'brand': '易来', 'specification': '灰色', 'count': 2},
                {'pdid': 7, 'device_name': '三键智能开关', 'brand': '易来', 'specification': '灰色', 'count': 3},
                {'pdid': 8, 'device_name': '四键智能开关', 'brand': '易来', 'specification': '灰色', 'count': 1}
            ]
        }
    }
    
    # 生成简要报告
    brief_report = generator.generate_brief_report(statistics_data)
    
    # 生成控制台输出
    console_output = generator.generate_console_output(brief_report)
    print(console_output)
    
    # 保存报告
    generator.save_brief_report(brief_report)
    generator.save_text_report(console_output)
    
    print("\n" + "=" * 60)
    print("✅ 简要报告生成模块测试完成")


if __name__ == "__main__":
    test_brief_report_generator()