#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI集成接口模块
为GUI应用提供统一的接口调用
"""

import os
import sys
import json
from datetime import datetime
from typing import Dict, List, Any, Optional
from excel_image_replacer import ExcelImageReplacer
from enhanced_procurement_generator import EnhancedProcurementGenerator


# 添加src目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

class GUIIntegration:
    """GUI集成接口类"""
    
    def __init__(self):
        self.project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
    def generate_mold_library(self, excel_file_path: str, custom_filename: str = None) -> Dict[str, Any]:
        """
        生成模具库PPT文件
        
        Args:
            excel_file_path: Excel模具库文件路径
            custom_filename: 自定义输出文件名（不含扩展名）
            
        Returns:
            Dict[str, Any]: 生成结果信息
        """
        try:
            # 首先保存设备图片到images目录
            images_dir = os.path.join(self.project_root, 'images')
            print(f"开始保存设备图片到目录: {images_dir}")
            
            # 导入图片保存控制器
            from image_save_controller import ImageSaveController
            
            # 创建图片保存控制器实例
            image_controller = ImageSaveController(excel_file_path, images_dir)
            
            # 运行设备图片保存流程
            image_save_success = image_controller.run_complete_workflow()
            
            if image_save_success:
                print("✓ 设备图片保存成功")
                
                # 获取保存结果摘要
                summary = image_controller.get_processing_summary()
                print(f"保存图片数量: {summary['results']['saved_count']}")
                
                # 检查images目录中是否有图片文件
                if os.path.exists(images_dir):
                    image_files = [f for f in os.listdir(images_dir) if f.endswith(('.png', '.jpg', '.jpeg'))]
                    print(f"images目录中的图片文件数量: {len(image_files)}")
            else:
                print("⚠ 设备图片保存失败，但继续生成PPT")
            
            # 导入Excel到PPT转换器
            from excel_to_ppt_converter import ExcelToPPTConverter
            
            # 创建转换器实例，使用统一的/images目录
            converter = ExcelToPPTConverter(image_folder=images_dir)
            
            # 生成输出文件路径
            if custom_filename:
                # 使用自定义文件名，保存在项目根目录
                output_dir = self.project_root
                output_file = os.path.join(output_dir, f"{custom_filename}.pptx")
            else:
                # 使用默认命名规则
                base_name = os.path.splitext(excel_file_path)[0]
                output_file = f"{base_name}_模具库.pptx"
            
            # 执行转换
            success = converter.generate_ppt_from_excel(excel_file_path, output_file)
            
            if success:
                # 返回结果信息
                return {
                    'success': True,
                    'message': 'PPT模具库生成成功',
                    'output_file': output_file,
                    'file_size': os.path.getsize(output_file) if os.path.exists(output_file) else 0,
                    'images_saved': image_save_success,
                    'image_count': summary['results']['saved_count'] if image_save_success else 0
                }
            else:
                return {
                    'success': False,
                    'message': 'PPT模具库生成失败',
                    'output_file': None,
                    'images_saved': image_save_success
                }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'生成PPT模具库失败: {str(e)}',
                'output_file': None,
                'error': str(e)
            }
    
    def generate_procurement_list(self, ppt_file_path: str, template_file_path: str, 
                                 mold_library_file_path: str, custom_filename: str) -> Dict[str, Any]:
        """
        生成采购清单的核心逻辑
        
        Args:
            ppt_file_path: PPT文件路径
            template_file_path: 模板文件路径
            mold_library_file_path: 模具库文件路径
            custom_filename: 自定义输出文件名 (不含.xlsx)
            
        Returns:
            Dict[str, Any]: 包含成功状态和文件路径的字典
        """
        try:
            # 1. 初始化增强型采购清单生成器
            generator = EnhancedProcurementGenerator()
            
            # 确定输出路径
            base_dir = os.path.dirname(ppt_file_path)
            output_filename = f"{custom_filename}.xlsx"
            output_file_path = os.path.join(base_dir, output_filename)
            
            # 2. 调用生成器生成 *包含DISPIMG公式* 的Excel文件
            success, errors = generator.generate_enhanced_procurement_list(
                template_path=template_file_path,
                mold_library_path=mold_library_file_path,
                ppt_file_path=ppt_file_path,
                output_path=output_file_path
            )
            
            if not success:
                error_msg = "\\n".join(errors) if errors else "未知错误"
                return {'success': False, 'message': f"生成采购清单失败: {error_msg}"}
            
            print(f"✅ 成功生成带DISPIMG公式的Excel文件: {output_file_path}")
            
            # --- [!! 关键修复 !!] ---
            # 3. 调用 ExcelImageReplacer 替换 DISPIMG 公式为真实图片
            # -------------------------
            print("🔄 开始替换DISPIMG公式为嵌入图片...")
            replacer = ExcelImageReplacer()
            
            # 定义最终带图片的文件路径
            output_with_images_path = output_file_path.replace('.xlsx', '_with_images.xlsx')
            
            # 使用我们之前确认过的正确列名 "L" 和 "I"
            replace_success = replacer.replace_dispimg_formulas(
                excel_path=output_file_path,
                output_path=output_with_images_path,
                pdid_column="I",    # <-- 必须是 "I"
                image_column="I",  # <-- 图片在 'I' 列，# <-- 必须是 "I"
                start_row=2       
            )
            
            if not replace_success:
                print("⚠️ 图片替换失败，返回原始DISPIMG文件")
                # 即使替换失败，也返回成功（但返回的是原始文件）
                return {
                    'success': True, 
                    'output_file': output_file_path,
                    'message': '采购清单生成成功，但图片替换失败'
                }

            print(f"✅ 图片替换完成，最终文件: {output_with_images_path}")
            
            # 4. 返回 *包含图片* 的最终文件路径
            return {
                'success': True,
                'output_file': output_with_images_path
            }
            
        except Exception as e:
            return {'success': False, 'message': f"生成采购清单时发生严重错误: {e}"}
    def get_template_info(self) -> Dict[str, Any]:
        """
        获取模板文件信息
        
        Returns:
            Dict[str, Any]: 模板文件信息
        """
        try:
            template_file = os.path.join(self.project_root, '采购清单模板.xlsx')
            mold_template = os.path.join(self.project_root, '智能家居模具库.xlsx')
            
            info = {
                'procurement_template': {
                    'exists': os.path.exists(template_file),
                    'path': template_file,
                    'size': os.path.getsize(template_file) if os.path.exists(template_file) else 0
                },
                'mold_template': {
                    'exists': os.path.exists(mold_template),
                    'path': mold_template,
                    'size': os.path.getsize(mold_template) if os.path.exists(mold_template) else 0
                }
            }
            
            return info
            
        except Exception as e:
            return {
                'error': str(e)
            }
    
    def validate_input_file(self, file_path: str, expected_type: str) -> Dict[str, Any]:
        """
        验证输入文件
        
        Args:
            file_path: 文件路径
            expected_type: 期望的文件类型 ('excel' 或 'ppt')
            
        Returns:
            Dict[str, Any]: 验证结果
        """
        try:
            if not os.path.exists(file_path):
                return {
                    'valid': False,
                    'message': '文件不存在'
                }
            
            # 检查文件大小（支持大文件，但给出警告）
            file_size = os.path.getsize(file_path) / (1024 * 1024)  # MB
            
            if file_size > 300:  # 300MB警告
                size_warning = f'文件较大 ({file_size:.1f}MB)，处理可能需要较长时间'
            else:
                size_warning = None
            
            # 检查文件扩展名
            ext = os.path.splitext(file_path)[1].lower()
            
            if expected_type == 'excel':
                valid_extensions = ['.xlsx', '.xls']
                if ext not in valid_extensions:
                    return {
                        'valid': False,
                        'message': f'不支持的文件类型: {ext}，请选择Excel文件(.xlsx, .xls)'
                    }
            elif expected_type == 'ppt':
                valid_extensions = ['.pptx', '.ppt']
                if ext not in valid_extensions:
                    return {
                        'valid': False,
                        'message': f'不支持的文件类型: {ext}，请选择PowerPoint文件(.pptx, .ppt)'
                    }
            
            return {
                'valid': True,
                'file_size_mb': file_size,
                'warning': size_warning,
                'message': '文件验证通过'
            }
            
        except Exception as e:
            return {
                'valid': False,
                'message': f'文件验证失败: {str(e)}'
            }
    
    def get_system_info(self) -> Dict[str, Any]:
        """
        获取系统信息
        
        Returns:
            Dict[str, Any]: 系统信息
        """
        try:
            # 检查关键模块是否存在
            modules = {
                'excel_to_ppt_converter': False,
                'smart_analyze_plan': False,
                'template_based_procurement_generator': False,
                'openpyxl': False,
                'python-pptx': False,
                'PIL': False
            }
            
            for module_name in modules.keys():
                try:
                    if module_name == 'excel_to_ppt_converter':
                        from excel_to_ppt_converter import ExcelToPPTConverter
                    elif module_name == 'smart_analyze_plan':
                        from smart_analyze_plan import smart_analyze_smart_home_plan
                    elif module_name == 'template_based_procurement_generator':
                        from template_based_procurement_generator import generate_procurement_list_with_template
                    elif module_name == 'openpyxl':
                        import openpyxl
                    elif module_name == 'python-pptx':
                        from pptx import Presentation
                    elif module_name == 'PIL':
                        from PIL import Image
                    
                    modules[module_name] = True
                except ImportError:
                    pass
            
            return {
                'modules': modules,
                'project_root': self.project_root,
                'timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            return {
                'error': str(e)
            }


def main():
    """测试函数"""
    integration = GUIIntegration()
    
    # 测试系统信息
    print("系统信息:")
    system_info = integration.get_system_info()
    print(json.dumps(system_info, indent=2, ensure_ascii=False))
    
    # 测试模板信息
    print("\n模板信息:")
    template_info = integration.get_template_info()
    print(json.dumps(template_info, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()