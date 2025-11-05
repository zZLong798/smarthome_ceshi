#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdid设备识别分析主程序 - 任务6：执行pdid设备识别分析
整合四个模块执行完整的pdid设备识别分析流程
"""

import os
import sys
import json
from datetime import datetime
from typing import Dict, Any

# 导入各个模块
from pdid_extractor import PDIDExtractor
from device_info_query import DeviceInfoQuery
from device_statistics import DeviceStatistics
from brief_report_generator import BriefReportGenerator


class PDIDAnalysisMain:
    """pdid设备识别分析主程序"""
    
    def __init__(self):
        """初始化主程序"""
        self.pdid_extractor = None
        self.device_query = DeviceInfoQuery()
        self.device_stats = DeviceStatistics()
        self.report_generator = BriefReportGenerator()
        self.analysis_results = {}
        
    def run_analysis(self, ppt_file_path: str) -> Dict[str, Any]:
        """
        运行完整的pdid设备识别分析
        
        Args:
            ppt_file_path: PPT文件路径
            
        Returns:
            Dict[str, Any]: 分析结果
        """
        print("🚀 开始pdid设备识别分析")
        print("=" * 80)
        print(f"📊 分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📄 目标文件: {ppt_file_path}")
        print("=" * 80)
        
        analysis_results = {
            'analysis_time': datetime.now().isoformat(),
            'ppt_file': ppt_file_path,
            'success': True,
            'errors': [],
            'warnings': [],
            'step_results': {}
        }
        
        try:
            # 步骤1: pdid标签提取
            print("\n📋 步骤1: pdid标签提取")
            print("-" * 40)
            
            if not os.path.exists(ppt_file_path):
                error_msg = f"PPT文件不存在: {ppt_file_path}"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            # 创建pdid提取器实例
            self.pdid_extractor = PDIDExtractor(ppt_file_path)
            
            # 加载PPT文件
            if not self.pdid_extractor.load_presentation():
                error_msg = "PPT文件加载失败"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            # 提取pdid标签
            pdid_labels = self.pdid_extractor.extract_pdid_labels()
            
            if not pdid_labels:
                error_msg = "未提取到任何pdid标签"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            # 计算pdid总数
            total_pdid_labels = sum(len(labels) for labels in pdid_labels.values())
            print(f"✅ 成功提取 {total_pdid_labels} 个pdid标签")
            
            # 保存pdid提取报告
            pdid_report_path = self.pdid_extractor.save_pdid_report(pdid_labels)
            print(f"💾 pdid提取报告已保存至: {pdid_report_path}")
            
            analysis_results['step_results']['pdid_extraction'] = {
                'status': 'success',
                'pdid_count': total_pdid_labels,
                'report_path': pdid_report_path
            }
            
            # 步骤2: 设备信息查询
            print("\n📋 步骤2: 设备信息查询")
            print("-" * 40)
            
            # 加载模具库
            if not self.device_query.load_product_library():
                error_msg = "加载模具库失败"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            # 从pdid提取报告中获取pdid列表
            pdid_counts = self.device_stats.load_pdid_data()
            if not pdid_counts:
                error_msg = "无法从pdid提取报告中获取pdid数据"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            print(f"📊 需要查询的pdid种类: {len(pdid_counts)} 种")
            
            # 查询设备信息
            device_mapping = self.device_query.query_devices_by_pdid_list(list(pdid_counts.keys()))
            
            if not device_mapping:
                error_msg = "未查询到任何设备信息"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            print(f"✅ 成功查询到 {len(device_mapping)} 种设备的详细信息")
            
            # 保存设备查询报告
            device_report_path = self.device_query.save_device_query_report(device_mapping)
            print(f"💾 设备查询报告已保存至: {device_report_path}")
            
            analysis_results['step_results']['device_query'] = {
                'status': 'success',
                'device_types_count': len(device_mapping),
                'report_path': device_report_path
            }
            
            # 步骤3: 设备统计
            print("\n📋 步骤3: 设备统计")
            print("-" * 40)
            
            # 统计设备数量和分类
            statistics = self.device_stats.count_devices_by_pdid(pdid_counts, device_mapping)
            
            if not statistics or statistics['total_devices'] == 0:
                error_msg = "设备统计失败或未统计到设备"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            statistics['statistics_time'] = datetime.now().isoformat()
            
            print(f"✅ 成功统计设备信息:")
            print(f"   • 设备总数: {statistics['total_devices']} 个")
            print(f"   • 设备种类: {statistics['unique_pdids']} 种")
            print(f"   • 品牌数量: {statistics['brands']} 个")
            print(f"   • 设备分类: {statistics['categories']} 类")
            
            # 保存统计报告
            stats_report_path = self.device_stats.save_statistics_report(statistics)
            print(f"💾 设备统计报告已保存至: {stats_report_path}")
            
            analysis_results['step_results']['device_statistics'] = {
                'status': 'success',
                'total_devices': statistics['total_devices'],
                'unique_pdids': statistics['unique_pdids'],
                'brands': statistics['brands'],
                'categories': statistics['categories'],
                'report_path': stats_report_path
            }
            
            # 步骤4: 简要报告生成
            print("\n📋 步骤4: 简要报告生成")
            print("-" * 40)
            
            # 加载统计数据进行报告生成
            statistics_data = self.report_generator.load_statistics_data()
            
            if not statistics_data:
                error_msg = "无法加载设备统计数据"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            # 生成简要报告
            brief_report = self.report_generator.generate_brief_report(statistics_data)
            
            if not brief_report or brief_report['total_devices'] == 0:
                error_msg = "简要报告生成失败"
                analysis_results['errors'].append(error_msg)
                analysis_results['success'] = False
                print(f"❌ {error_msg}")
                return analysis_results
            
            # 生成控制台输出
            console_output = self.report_generator.generate_console_output(brief_report)
            print(console_output)
            
            # 保存报告
            json_report_path = self.report_generator.save_brief_report(brief_report)
            text_report_path = self.report_generator.save_text_report(console_output)
            
            print(f"💾 JSON格式报告已保存至: {json_report_path}")
            print(f"💾 文本格式报告已保存至: {text_report_path}")
            
            analysis_results['step_results']['brief_report'] = {
                'status': 'success',
                'total_devices': brief_report['total_devices'],
                'total_brands': brief_report['total_brands'],
                'json_report_path': json_report_path,
                'text_report_path': text_report_path
            }
            
            # 分析完成
            print("\n" + "=" * 80)
            print("🎉 pdid设备识别分析完成！")
            print("=" * 80)
            
            # 生成分析总结
            self._generate_analysis_summary(analysis_results)
            
        except Exception as e:
            error_msg = f"分析过程中发生错误: {e}"
            analysis_results['errors'].append(error_msg)
            analysis_results['success'] = False
            print(f"❌ {error_msg}")
            
        return analysis_results
    
    def _generate_analysis_summary(self, analysis_results: Dict[str, Any]) -> None:
        """
        生成分析总结
        
        Args:
            analysis_results: 分析结果
        """
        print("\n📊 分析总结")
        print("-" * 40)
        
        if analysis_results['success']:
            print("✅ 分析成功完成")
            
            # 获取统计信息
            stats = analysis_results['step_results']['device_statistics']
            brief = analysis_results['step_results']['brief_report']
            
            print(f"📈 关键指标:")
            print(f"   • 设备总数: {stats['total_devices']} 个")
            print(f"   • 设备种类: {stats['unique_pdids']} 种")
            print(f"   • 品牌数量: {stats['brands']} 个")
            print(f"   • 设备分类: {stats['categories']} 类")
            
            print(f"\n📋 生成报告:")
            print(f"   • pdid提取报告: {analysis_results['step_results']['pdid_extraction']['report_path']}")
            print(f"   • 设备查询报告: {analysis_results['step_results']['device_query']['report_path']}")
            print(f"   • 设备统计报告: {analysis_results['step_results']['device_statistics']['report_path']}")
            print(f"   • 简要设备清单(JSON): {brief['json_report_path']}")
            print(f"   • 简要设备清单(文本): {brief['text_report_path']}")
            
            # 保存分析结果
            summary_path = self._save_analysis_summary(analysis_results)
            print(f"\n💾 分析总结已保存至: {summary_path}")
            
        else:
            print("❌ 分析失败")
            print(f"错误数量: {len(analysis_results['errors'])}")
            for error in analysis_results['errors']:
                print(f"   • {error}")
    
    def _save_analysis_summary(self, analysis_results: Dict[str, Any]) -> str:
        """
        保存分析总结
        
        Args:
            analysis_results: 分析结果
            
        Returns:
            str: 保存路径
        """
        summary_file = "pdid_analysis_summary.json"
        
        # 简化分析结果，只保留关键信息
        summary = {
            'analysis_time': analysis_results['analysis_time'],
            'ppt_file': analysis_results['ppt_file'],
            'success': analysis_results['success'],
            'key_metrics': {},
            'report_paths': {}
        }
        
        if analysis_results['success']:
            # 添加关键指标
            stats = analysis_results['step_results']['device_statistics']
            summary['key_metrics'] = {
                'total_devices': stats['total_devices'],
                'unique_pdids': stats['unique_pdids'],
                'brands': stats['brands'],
                'categories': stats['categories']
            }
            
            # 添加报告路径
            summary['report_paths'] = {
                'pdid_extraction': analysis_results['step_results']['pdid_extraction']['report_path'],
                'device_query': analysis_results['step_results']['device_query']['report_path'],
                'device_statistics': analysis_results['step_results']['device_statistics']['report_path'],
                'brief_report_json': analysis_results['step_results']['brief_report']['json_report_path'],
                'brief_report_text': analysis_results['step_results']['brief_report']['text_report_path']
            }
        
        # 保存总结文件
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        return summary_file


def main():
    """主函数"""
    
    # 创建主程序实例
    analyzer = PDIDAnalysisMain()
    
    # 设置PPT文件路径
    ppt_file_path = "../output/修复后的全屋智能方案.pptx"
    
    # 检查文件是否存在
    if not os.path.exists(ppt_file_path):
        print(f"⚠️ 目标PPT文件不存在: {ppt_file_path}")
        print("💡 请确保PPT文件已放置在正确位置")
        print("📁 预期路径: ../output/修复后的全屋智能方案.pptx")
        return
    
    # 运行分析
    results = analyzer.run_analysis(ppt_file_path)
    
    # 输出最终结果
    print("\n" + "=" * 80)
    if results['success']:
        print("🎉 pdid设备识别分析流程执行完成！")
        print("📊 所有报告已生成，请查看相关文件。")
    else:
        print("❌ pdid设备识别分析流程执行失败")
        print("📋 请检查错误信息并重新运行。")
    print("=" * 80)


if __name__ == "__main__":
    main()