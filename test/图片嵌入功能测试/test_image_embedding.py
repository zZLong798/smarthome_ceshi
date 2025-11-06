#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片嵌入功能测试脚本
测试采购清单生成器中的DISPIMG公式替换功能
"""

import sys
import os
import json

# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from procurement_list_generator import ProcurementListGenerator
from excel_image_replacer import ExcelImageReplacer

def test_procurement_list_generation():
    """测试采购清单生成器"""
    print("🧪 开始测试采购清单生成器...")
    
    try:
        # 创建采购清单生成器实例
        generator = ProcurementListGenerator()
        
        # 加载统计报告数据
        statistics_data = generator.load_statistics_data("device_statistics_report.json")
        if not statistics_data:
            print("❌ 无法加载统计报告数据")
            return False
        
        print(f"✅ 成功加载统计报告数据，共 {len(statistics_data)} 个设备")
        
        # 生成采购清单数据
        procurement_list = generator.generate_device_procurement_list(statistics_data)
        if not procurement_list:
            print("❌ 无法生成采购清单数据")
            return False
        
        print(f"✅ 成功生成采购清单数据，共 {len(procurement_list)} 个条目")
        
        # 测试1：生成包含DISPIMG公式的Excel文件
        print("\n📝 测试1：生成包含DISPIMG公式的Excel文件...")
        dispimg_output = "test_dispimg_formulas.xlsx"
        success = generator.save_procurement_list(
            procurement_list, 
            dispimg_output, 
            use_dispimg_formulas=True
        )
        
        if success:
            print(f"✅ 成功生成DISPIMG公式文件: {dispimg_output}")
        else:
            print("❌ 生成DISPIMG公式文件失败")
            return False
        
        # 测试2：生成直接嵌入图片的Excel文件
        print("\n🖼️  测试2：生成直接嵌入图片的Excel文件...")
        direct_output = "test_direct_images.xlsx"
        success = generator.save_procurement_list(
            procurement_list, 
            direct_output, 
            use_dispimg_formulas=False
        )
        
        if success:
            print(f"✅ 成功生成直接嵌入图片文件: {direct_output}")
        else:
            print("❌ 生成直接嵌入图片文件失败")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试过程中发生错误: {e}")
        return False

def test_excel_image_replacer():
    """测试Excel图片替换器"""
    print("\n🔄 开始测试Excel图片替换器...")
    
    try:
        # 创建图片替换器实例
        replacer = ExcelImageReplacer()
        
        # 测试图片映射加载
        print("📋 测试图片映射加载...")
        # 检查映射是否已加载
        if hasattr(replacer, 'image_mapping') and replacer.image_mapping:
            mapping = replacer.image_mapping
            print(f"✅ 成功加载图片映射，共 {len(mapping)} 个映射关系")
            # 显示前5个映射关系
            for i, (pdid, image_path) in enumerate(list(mapping.items())[:5]):
                print(f"   {i+1}. PDID: {pdid} -> 图片: {image_path}")
        else:
            print("❌ 图片映射加载失败")
            return False
        
        # 测试单个图片路径查找
        print("\n🔍 测试单个图片路径查找...")
        test_pdid = "1"  # 使用数字PDID，因为映射中是数字格式
        image_path = replacer.image_mapping.get(test_pdid)
        
        if image_path and os.path.exists(image_path):
            print(f"✅ 成功找到PDID {test_pdid} 的图片: {image_path}")
        else:
            print(f"⚠️  未找到PDID {test_pdid} 的图片")
            # 显示映射中实际存在的PDID示例
            available_pdids = list(replacer.image_mapping.keys())[:3]
            print(f"   映射中存在的PDID示例: {available_pdids}")
        
        # 测试图片替换功能
        print("\n🔄 测试图片替换功能...")
        
        # 首先需要有一个包含DISPIMG公式的Excel文件
        if not os.path.exists("test_dispimg_formulas.xlsx"):
            print("⚠️  未找到测试文件，请先运行采购清单生成器测试")
            return False
        
        output_path = "test_replaced_images.xlsx"
        success = replacer.replace_dispimg_formulas(
            excel_path="test_dispimg_formulas.xlsx",
            output_path=output_path,
            pdid_column="A",
            image_column="I",
            start_row=2
        )
        
        if success:
            print(f"✅ 图片替换成功，生成文件: {output_path}")
        else:
            print("❌ 图片替换失败")
        
        return success
        
    except Exception as e:
        print(f"❌ 测试过程中发生错误: {e}")
        return False

def main():
    """主测试函数"""
    print("=" * 60)
    print("📊 图片嵌入功能测试")
    print("=" * 60)
    
    # 检查必要的文件是否存在
    required_files = [
        "device_statistics_report.json",
        "images/image_mapping.json"
    ]
    
    for file_path in required_files:
        if not os.path.exists(file_path):
            print(f"❌ 缺少必要文件: {file_path}")
            print("请确保在项目根目录下运行此测试")
            return
    
    # 运行测试
    success1 = test_procurement_list_generation()
    success2 = test_excel_image_replacer()
    
    print("\n" + "=" * 60)
    print("📋 测试结果汇总:")
    print(f"   采购清单生成器测试: {'✅ 通过' if success1 else '❌ 失败'}")
    print(f"   Excel图片替换器测试: {'✅ 通过' if success2 else '❌ 失败'}")
    
    if success1 and success2:
        print("\n🎉 所有测试通过！图片嵌入功能正常工作。")
        print("\n📁 生成的文件:")
        for file_name in ["test_dispimg_formulas.xlsx", "test_direct_images.xlsx", "test_replaced_images.xlsx"]:
            if os.path.exists(file_name):
                print(f"   - {file_name}")
    else:
        print("\n⚠️  部分测试失败，请检查错误信息。")
    
    print("=" * 60)

if __name__ == "__main__":
    main()