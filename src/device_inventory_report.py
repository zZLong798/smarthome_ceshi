#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设备清单报告生成模块 - 任务7：生成设备清单报告
生成详细的设备清单报告，包含完整的设备信息和统计
"""

import json
import os
from datetime import datetime
from typing import Dict, List, Any


class DeviceInventoryReport:
    """设备清单报告生成器"""
    
    def __init__(self):
        """初始化报告生成器"""
        self.report_data = {}
        
    def load_statistics_data(self) -> Dict[str, Any]:
        """
        加载设备统计数据
        
        Returns:
            Dict[str, Any]: 设备统计数据
        """
        try:
            # 加载设备统计报告
            stats_file = "device_statistics_report.json"
            if os.path.exists(stats_file):
                with open(stats_file, 'r', encoding='utf-8') as f:
                    statistics_data = json.load(f)
                print(f"✅ 成功加载设备统计数据")
                return statistics_data
            else:
                print(f"⚠️ 设备统计报告文件不存在: {stats_file}")
                return {}
        except Exception as e:
            print(f"❌ 加载设备统计数据失败: {e}")
            return {}
    
    def generate_inventory_report(self, statistics_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        生成设备清单报告
        
        Args:
            statistics_data: 设备统计数据
            
        Returns:
            Dict[str, Any]: 设备清单报告
        """
        print("📋 开始生成设备清单报告...")
        
        if not statistics_data:
            print("❌ 设备统计数据为空")
            return {}
        
        # 创建报告结构
        inventory_report = {
            'report_type': '设备清单报告',
            'generated_time': datetime.now().isoformat(),
            'summary': {},
            'inventory_by_brand': {},
            'inventory_by_category': {},
            'detailed_inventory': [],
            'statistical_analysis': {}
        }
        
        # 提取总体统计信息
        inventory_report['summary'] = {
            'total_devices': statistics_data.get('total_devices', 0),
            'unique_pdids': statistics_data.get('unique_pdids', 0),
            'brands': statistics_data.get('brands', 0),
            'categories': statistics_data.get('categories', 0),
            'total_price': statistics_data.get('total_price', 0)
        }
        
        # 按品牌分类的设备清单
        if 'brand_stats' in statistics_data:
            inventory_report['inventory_by_brand'] = statistics_data['brand_stats']
        
        # 按分类分类的设备清单
        if 'category_stats' in statistics_data:
            inventory_report['inventory_by_category'] = statistics_data['category_stats']
        
        # 详细设备清单 - 从品牌统计中提取所有设备
        if 'brand_stats' in statistics_data:
            detailed_inventory = []
            for brand, devices in statistics_data['brand_stats'].items():
                for device in devices:
                    device_info = device.copy()
                    device_info['brand'] = brand
                    detailed_inventory.append(device_info)
            inventory_report['detailed_inventory'] = detailed_inventory
        
        # 统计分析
        inventory_report['statistical_analysis'] = self._generate_statistical_analysis(statistics_data)
        
        print("✅ 设备清单报告生成完成")
        return inventory_report
    
    def _generate_statistical_analysis(self, statistics_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        生成统计分析
        
        Args:
            statistics_data: 设备统计数据
            
        Returns:
            Dict[str, Any]: 统计分析结果
        """
        analysis = {
            'device_distribution': {},
            'brand_distribution': {},
            'category_distribution': {},
            'key_insights': []
        }
        
        # 设备分布分析
        if 'brand_stats' in statistics_data:
            devices = []
            for brand, brand_devices in statistics_data['brand_stats'].items():
                devices.extend(brand_devices)
            
            # 按设备类型分布
            type_distribution = {}
            for device in devices:
                device_type = device.get('device_name', '未知')
                count = device.get('count', 0)
                if device_type in type_distribution:
                    type_distribution[device_type] += count
                else:
                    type_distribution[device_type] = count
            
            analysis['device_distribution'] = type_distribution
        
        # 品牌分布分析
        if 'brand_stats' in statistics_data:
            brand_stats = statistics_data['brand_stats']
            brand_distribution = {}
            
            for brand, devices in brand_stats.items():
                total_count = 0
                for device in devices:
                    total_count += device.get('count', 0)
                brand_distribution[brand] = total_count
            
            analysis['brand_distribution'] = brand_distribution
        
        # 分类分布分析
        if 'category_stats' in statistics_data:
            category_stats = statistics_data['category_stats']
            category_distribution = {}
            
            for category, devices in category_stats.items():
                total_count = 0
                for device in devices:
                    total_count += device.get('count', 0)
                category_distribution[category] = total_count
            
            analysis['category_distribution'] = category_distribution
        
        # 关键洞察
        analysis['key_insights'] = self._generate_key_insights(statistics_data)
        
        return analysis
    
    def _generate_key_insights(self, statistics_data: Dict[str, Any]) -> List[str]:
        """
        生成关键洞察
        
        Args:
            statistics_data: 设备统计数据
            
        Returns:
            List[str]: 关键洞察列表
        """
        insights = []
        
        total_devices = statistics_data.get('total_devices', 0)
        unique_pdids = statistics_data.get('unique_pdids', 0)
        brands = statistics_data.get('brands', 0)
        categories = statistics_data.get('categories', 0)
        
        if total_devices > 0:
            insights.append(f"设备总数: {total_devices} 个")
            insights.append(f"设备种类: {unique_pdids} 种")
            insights.append(f"品牌数量: {brands} 个")
            insights.append(f"设备分类: {categories} 类")
        
        # 品牌分析
        if 'brand_statistics' in statistics_data:
            brand_stats = statistics_data['brand_statistics']
            if brand_stats:
                brand_counts = {}
                for brand, devices in brand_stats.items():
                    total_count = sum(device.get('count', 0) for device in devices)
                    brand_counts[brand] = total_count
                
                if brand_counts:
                    max_brand = max(brand_counts, key=brand_counts.get)
                    max_count = brand_counts[max_brand]
                    percentage = (max_count / total_devices) * 100
                    insights.append(f"主要品牌: {max_brand} (占比: {percentage:.1f}%)")
        
        # 设备类型分析
        if 'detailed_statistics' in statistics_data:
            devices = statistics_data['detailed_statistics']
            if devices:
                type_counts = {}
                for device in devices:
                    device_type = device.get('device_name', '未知')
                    count = device.get('count', 0)
                    if device_type in type_counts:
                        type_counts[device_type] += count
                    else:
                        type_counts[device_type] = count
                
                if type_counts:
                    max_type = max(type_counts, key=type_counts.get)
                    max_count = type_counts[max_type]
                    percentage = (max_count / total_devices) * 100
                    insights.append(f"主要设备类型: {max_type} (占比: {percentage:.1f}%)")
        
        return insights
    
    def generate_console_output(self, inventory_report: Dict[str, Any]) -> str:
        """
        生成控制台输出格式
        
        Args:
            inventory_report: 设备清单报告
            
        Returns:
            str: 控制台输出内容
        """
        output = []
        
        # 报告标题
        output.append("📋 设备清单报告")
        output.append("=" * 60)
        
        # 总体统计
        summary = inventory_report.get('summary', {})
        if summary:
            output.append("📊 总体统计:")
            output.append(f"   • 设备总数: {summary.get('total_devices', 0)} 个")
            output.append(f"   • 设备种类: {summary.get('unique_pdids', 0)} 种")
            output.append(f"   • 品牌数量: {summary.get('brands', 0)} 个")
            output.append(f"   • 设备分类: {summary.get('categories', 0)} 类")
            if summary.get('total_price', 0) > 0:
                output.append(f"   • 总价值: ¥{summary.get('total_price', 0):,.2f}")
        
        # 按品牌分类的设备清单
        inventory_by_brand = inventory_report.get('inventory_by_brand', {})
        if inventory_by_brand:
            output.append("\n🏷️ 按品牌分类的设备清单:")
            for brand, devices in inventory_by_brand.items():
                total_count = sum(device.get('count', 0) for device in devices)
                output.append(f"\n   📍 {brand} (总计: {total_count} 个):")
                for device in devices:
                    output.append(f"      📱 {device.get('device_name', '未知')}")
                    output.append(f"         • 规格: {device.get('specification', '未知')}")
                    output.append(f"         • 数量: {device.get('count', 0)} 个")
                    if device.get('unit_price', 0) > 0:
                        output.append(f"         • 单价: ¥{device.get('unit_price', 0):,.2f}")
        
        # 统计分析
        statistical_analysis = inventory_report.get('statistical_analysis', {})
        if statistical_analysis:
            output.append("\n📈 统计分析:")
            
            # 关键洞察
            key_insights = statistical_analysis.get('key_insights', [])
            if key_insights:
                output.append("   🔍 关键洞察:")
                for insight in key_insights:
                    output.append(f"      • {insight}")
            
            # 品牌分布
            brand_distribution = statistical_analysis.get('brand_distribution', {})
            if brand_distribution:
                output.append("\n   🏷️ 品牌分布:")
                for brand, count in brand_distribution.items():
                    percentage = (count / summary.get('total_devices', 1)) * 100
                    output.append(f"      • {brand}: {count} 个 ({percentage:.1f}%)")
            
            # 设备类型分布
            device_distribution = statistical_analysis.get('device_distribution', {})
            if device_distribution:
                output.append("\n   📱 设备类型分布:")
                for device_type, count in device_distribution.items():
                    percentage = (count / summary.get('total_devices', 1)) * 100
                    output.append(f"      • {device_type}: {count} 个 ({percentage:.1f}%)")
        
        # 报告生成时间
        generated_time = inventory_report.get('generated_time', '')
        if generated_time:
            try:
                dt = datetime.fromisoformat(generated_time.replace('Z', '+00:00'))
                output.append(f"\n⏰ 报告生成时间: {dt.strftime('%Y-%m-%d %H:%M:%S')}")
            except:
                output.append(f"\n⏰ 报告生成时间: {generated_time}")
        
        return '\n'.join(output)
    
    def save_inventory_report(self, inventory_report: Dict[str, Any]) -> str:
        """
        保存设备清单报告为JSON文件
        
        Args:
            inventory_report: 设备清单报告
            
        Returns:
            str: 保存的文件路径
        """
        report_file = "device_inventory_report.json"
        
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(inventory_report, f, ensure_ascii=False, indent=2)
            print(f"💾 设备清单报告已保存至: {report_file}")
            return report_file
        except Exception as e:
            print(f"❌ 保存设备清单报告失败: {e}")
            return ""
    
    def save_text_report(self, console_output: str) -> str:
        """
        保存文本格式的设备清单报告
        
        Args:
            console_output: 控制台输出内容
            
        Returns:
            str: 保存的文件路径
        """
        report_file = "device_inventory_report.txt"
        
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write(console_output)
            print(f"💾 文本格式设备清单报告已保存至: {report_file}")
            return report_file
        except Exception as e:
            print(f"❌ 保存文本格式设备清单报告失败: {e}")
            return ""


def test_device_inventory_report():
    """测试设备清单报告生成功能"""
    print("🧪 测试设备清单报告生成功能")
    print("=" * 50)
    
    # 创建报告生成器
    report_generator = DeviceInventoryReport()
    
    # 加载统计数据
    statistics_data = report_generator.load_statistics_data()
    
    if not statistics_data:
        print("❌ 无法加载设备统计数据，测试终止")
        return
    
    # 生成设备清单报告
    inventory_report = report_generator.generate_inventory_report(statistics_data)
    
    if not inventory_report:
        print("❌ 设备清单报告生成失败")
        return
    
    # 生成控制台输出
    console_output = report_generator.generate_console_output(inventory_report)
    print(console_output)
    
    # 保存报告
    json_report_path = report_generator.save_inventory_report(inventory_report)
    text_report_path = report_generator.save_text_report(console_output)
    
    if json_report_path and text_report_path:
        print("✅ 设备清单报告生成测试完成")
    else:
        print("❌ 设备清单报告保存失败")


if __name__ == "__main__":
    test_device_inventory_report()