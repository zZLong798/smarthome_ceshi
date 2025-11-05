#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设备图片保存流程主控制器
协调Excel图片提取、重命名和保存的完整流程
"""

import os
import time
import re
from typing import Dict, List, Optional, Any
from excel_image_extractor import ExcelImageExtractor
from image_renamer import ImageRenamer
from image_saver import ImageSaver
from image_mapping_generator import ImageMappingGenerator


class ImageSaveController:
    """设备图片保存流程主控制器"""
    
    def __init__(self, excel_file_path: str, output_dir: str = None):
        """
        初始化控制器
        
        Args:
            excel_file_path: Excel文件路径
            output_dir: 输出目录（默认为项目根目录下的images文件夹）
        """
        self.excel_file_path = excel_file_path
        
        # 如果未指定输出目录，则使用项目根目录下的images文件夹
        if output_dir is None:
            # 获取项目根目录（当前文件所在目录的父目录）
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            self.output_dir = os.path.join(project_root, "images")
        else:
            self.output_dir = output_dir
            
        # 确保输出目录存在
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        self.temp_dir = "temp_image_processing"
        
        # 初始化各模块
        self.extractor = ExcelImageExtractor(excel_file_path, self.temp_dir)
        self.renamer = ImageRenamer(excel_file_path, self.temp_dir)
        self.saver = ImageSaver(self.output_dir)
        self.mapping_generator = ImageMappingGenerator(excel_file_path, self.output_dir)
        
        # 初始化处理状态
        self.processing_status = {
            'extraction': False,
            'renaming': False,
            'saving': False,
            'cleanup': False
        }
        
        # 初始化增强版映射数据
        self.enhanced_mapping_data = None
        
        # 处理结果
        self.processing_results = {
            'extracted_images': {},
            'renamed_images': {},
            'saved_images': {},
            'errors': []
        }
    
    def run_complete_workflow(self) -> bool:
        """
        运行完整的图片保存流程
        
        Returns:
            bool: 是否成功完成
        """
        print("=" * 60)
        print("开始设备图片保存流程")
        print("=" * 60)
        
        start_time = time.time()
        
        try:
            # 1. 提取Excel中的图片
            if not self._extract_images():
                self.processing_results['errors'].append("图片提取失败")
                return False
            
            # 2. 重命名图片文件
            if not self._rename_images():
                self.processing_results['errors'].append("图片重命名失败")
                return False
            
            # 3. 保存图片文件
            if not self._save_images():
                self.processing_results['errors'].append("图片保存失败")
                return False
            
            # 4. 清理临时文件
            self._cleanup()
            
            # 5. 生成处理报告
            self._generate_report()
            
            end_time = time.time()
            processing_time = end_time - start_time
            
            print("=" * 60)
            print(f"设备图片保存流程完成!")
            print(f"处理时间: {processing_time:.2f} 秒")
            print(f"提取图片: {len(self.processing_results['extracted_images'])} 张")
            print(f"重命名图片: {len(self.processing_results['renamed_images'])} 张")
            print(f"保存图片: {len(self.processing_results['saved_images'])} 张")
            print(f"错误数量: {len(self.processing_results['errors'])} 个")
            print("=" * 60)
            
            return len(self.processing_results['errors']) == 0
            
        except Exception as e:
            error_msg = f"流程执行异常: {e}"
            self.processing_results['errors'].append(error_msg)
            print(error_msg)
            return False
    
    def _extract_images(self) -> bool:
        """提取Excel中的图片（使用增强版映射功能）"""
        print("\n步骤1: 提取Excel中的图片")
        print("-" * 40)
        
        try:
            # 创建临时提取目录
            extract_dir = os.path.join(self.temp_dir, "extracted")
            if not os.path.exists(extract_dir):
                os.makedirs(extract_dir)
            
            # 首先尝试使用增强版映射器建立完整映射关系
            from src.enhanced_excel_image_mapper import EnhancedExcelImageMapper
            
            # 创建映射器但不立即清理
            mapper = EnhancedExcelImageMapper(self.excel_file_path)
            
            try:
                # 解析映射关系
                mapping_data = mapper.parse_enhanced_mapping()
                
                if mapping_data:
                    print("✓ 增强版映射解析成功，使用改进的映射功能")
                    # 提取映射链数据供后续使用
                    self.enhanced_mapping_data = mapping_data
                    print(f"增强版映射数据包含 {len(self.enhanced_mapping_data)} 条记录")
                    
                    # 使用增强版映射器提供的图片路径直接复制图片
                    extracted_images = {}
                    for mapping_key, mapping_info in self.enhanced_mapping_data.items():
                        file_path = mapping_info.get('file_path', '')
                        if file_path and os.path.exists(file_path):
                            # 获取原文件名
                            original_filename = os.path.basename(file_path)
                            # 目标文件路径
                            target_path = os.path.join(extract_dir, original_filename)
                            
                            # 复制图片文件
                            import shutil
                            shutil.copy2(file_path, target_path)
                            
                            # 记录提取的图片
                            extracted_images[original_filename] = target_path
                            print(f"✓ 复制图片: {original_filename}")
                        else:
                            print(f"⚠ 图片文件不存在或路径无效: {file_path}")
                    
                    if extracted_images:
                        print(f"✓ 使用增强版映射器提取图片成功: {len(extracted_images)} 张图片")
                        extracted_images = extracted_images
                    else:
                        print("⚠ 增强版映射器未找到有效图片，回退到基础提取功能")
                        # 回退到基础提取功能
                        extracted_images = self.extractor.extract_all_images(extract_dir)
                else:
                    print("⚠ 增强版映射解析失败，回退到基础提取功能")
                    # 回退到基础提取功能
                    extracted_images = self.extractor.extract_all_images(extract_dir)
                    # 清空增强版映射数据
                    self.enhanced_mapping_data = None
            finally:
                # 在复制完图片后再清理临时文件
                mapper.cleanup()
            
            if extracted_images:
                self.processing_results['extracted_images'] = extracted_images
                self.processing_status['extraction'] = True
                print(f"✓ 图片提取成功: {len(extracted_images)} 张图片")
                return True
            else:
                print("✗ 未提取到任何图片")
                return False
                
        except Exception as e:
            error_msg = f"图片提取失败: {e}"
            self.processing_results['errors'].append(error_msg)
            print(error_msg)
            return False
    
    def _rename_images(self) -> bool:
        """重命名图片文件"""
        print("\n步骤2: 重命名图片文件")
        print("-" * 40)
        
        try:
            # 获取提取的图片
            extracted_images = self.processing_results['extracted_images']
            
            if not extracted_images:
                print("✗ 没有可重命名的图片")
                return False
            
            # 如果有增强版映射数据，使用它来指导重命名
            if hasattr(self, 'enhanced_mapping_data') and self.enhanced_mapping_data:
                print("使用增强版映射数据进行智能重命名...")
                renamed_images = {}
                
                # 基于增强版映射数据创建重命名映射
                for mapping_key, mapping_info in self.enhanced_mapping_data.items():
                    pdid = mapping_info.get('pdid', '')
                    actual_file = mapping_info.get('actual_file', '')
                    image_id = mapping_info.get('image_id', '')
                    device_name = mapping_info.get('device_name', '')
                    
                    if pdid and actual_file and image_id:
                        # 生成新的文件名：pdid{id}_{设备简称拼音}.png
                        # 使用控制器内部的拼音转换方法
                        device_pinyin = self._convert_device_name_to_pinyin(device_name) if device_name else 'device'
                        new_filename = f"pdid{pdid}_{device_pinyin}.png"
                        
                        # 在提取的图片中找到对应的文件
                        for original_path in extracted_images.values():
                            if image_id in original_path or actual_file in original_path:
                                renamed_images[new_filename] = original_path
                                print(f"  映射: {actual_file} -> {new_filename} (PDID: {pdid})")
                                break
                
                if renamed_images:
                    self.processing_results['renamed_images'] = renamed_images
                    self.processing_status['renaming'] = True
                    print(f"✓ 基于增强版映射重命名成功: {len(renamed_images)} 张图片")
                    return True
            
            # 如果没有增强版映射数据，使用传统重命名器
            print("使用传统重命名器...")
            # 重命名图片 - 提取正确的映射格式
            if isinstance(extracted_images, dict):
                # 使用增强版映射器返回的格式
                correct_mapping = extracted_images.get('correct_mapping', {})
                if correct_mapping:
                    image_mapping = correct_mapping
                else:
                    # 如果没有correct_mapping，使用extracted_images作为备选
                    image_mapping = extracted_images.get('extracted_images', {})
            else:
                # 基础格式（直接是图片映射）
                image_mapping = extracted_images
            
            renamed_images = self.renamer.rename_images(image_mapping)
            
            if renamed_images:
                self.processing_results['renamed_images'] = renamed_images
                self.processing_status['renaming'] = True
                print(f"✓ 图片重命名成功: {len(renamed_images)} 张图片")
                
                # 显示重命名结果
                print("重命名结果:")
                for new_name, original_path in renamed_images.items():
                    print(f"  {os.path.basename(original_path)} -> {new_name}")
                
                return True
            else:
                print("✗ 图片重命名失败")
                return False
                
        except Exception as e:
            error_msg = f"图片重命名失败: {e}"
            self.processing_results['errors'].append(error_msg)
            print(error_msg)
            return False
    
    def _build_mapping_data(self, renamed_images: Dict[str, str], saved_images: Dict[str, str]) -> Dict[str, Any]:
        """构建完整的映射数据结构"""
        mapping_data = {
            'original_images': [],
            'saved_images': [],
            'mapping_relationships': []
        }
        
        # 尝试获取增强版映射数据
        enhanced_mapping_info = {}
        if hasattr(self, 'enhanced_mapping_data') and self.enhanced_mapping_data is not None:
            # 从增强版映射数据中获取图片ID到行号、设备名称、产品ID的映射
            for mapping_key, mapping_info in self.enhanced_mapping_data.items():
                if 'image_id' in mapping_info:
                    image_id = mapping_info['image_id']
                    enhanced_mapping_info[image_id] = {
                        'row_number': mapping_info.get('row_number', 0),
                        'device_name': mapping_info.get('device_name', ''),
                        'product_id': mapping_info.get('product_id', ''),
                        'dispimg_formula': mapping_info.get('dispimg_formula', ''),
                        'cell_reference': mapping_info.get('cell_reference', '')  # 添加cell_reference字段
                    }
        
        # 构建原图片信息
        for new_filename, original_path in renamed_images.items():
            # 从原文件路径中提取基本信息
            original_filename = os.path.basename(original_path) if original_path else ''
            
            # 尝试从文件名中提取图片ID（假设文件名就是图片ID）
            file_based_image_id = os.path.splitext(original_filename)[0] if original_filename else new_filename
            
            # 从增强版映射数据获取详细信息
            image_id = file_based_image_id
            enhanced_info = {}
            
            # 查找匹配的增强映射信息
            if hasattr(self, 'enhanced_mapping_data') and self.enhanced_mapping_data is not None:
                # enhanced_mapping_data是一个字典，键是pdid，值是包含映射信息的字典
                for pdid, mapping_info in self.enhanced_mapping_data.items():
                    # 检查是否通过图片ID匹配（最可靠的匹配方式）
                    if mapping_info.get('image_id', '') == file_based_image_id:
                        image_id = mapping_info.get('image_id', file_based_image_id)
                        enhanced_info = {
                            'row_number': mapping_info.get('row_number', 0),
                            'device_name': mapping_info.get('device_name', ''),
                            'product_id': mapping_info.get('pdid', ''),  # 使用pdid作为product_id
                            'dispimg_formula': mapping_info.get('dispimg_formula', ''),
                            'cell_reference': mapping_info.get('cell_reference', '')
                        }
                        print(f"✓ 通过图片ID匹配: {file_based_image_id} -> PDID={mapping_info.get('pdid', '')}")
                        break
                    # 检查是否通过文件路径匹配（包含图片ID）
                    elif original_path and mapping_info.get('image_id', '') in original_path:
                        image_id = mapping_info.get('image_id', file_based_image_id)
                        enhanced_info = {
                            'row_number': mapping_info.get('row_number', 0),
                            'device_name': mapping_info.get('device_name', ''),
                            'product_id': mapping_info.get('pdid', ''),  # 使用pdid作为product_id
                            'dispimg_formula': mapping_info.get('dispimg_formula', ''),
                            'cell_reference': mapping_info.get('cell_reference', '')
                        }
                        print(f"✓ 通过路径匹配: {original_path} -> PDID={mapping_info.get('pdid', '')}")
                        break
                    # 检查是否通过文件名匹配（实际文件名） - 这个匹配方式可能不可靠
                    elif mapping_info.get('actual_file', '') == original_filename:
                        image_id = mapping_info.get('image_id', file_based_image_id)
                        enhanced_info = {
                            'row_number': mapping_info.get('row_number', 0),
                            'device_name': mapping_info.get('device_name', ''),
                            'product_id': mapping_info.get('pdid', ''),  # 使用pdid作为product_id
                            'dispimg_formula': mapping_info.get('dispimg_formula', ''),
                            'cell_reference': mapping_info.get('cell_reference', '')
                        }
                        print(f"✓ 通过文件名匹配: {original_filename} -> PDID={mapping_info.get('pdid', '')}")
                        break
            
            # 如果文件名以ID_开头，直接使用它作为image_id
            if file_based_image_id.startswith('ID_'):
                image_id = file_based_image_id
            
            # 如果没有找到匹配信息，记录警告
            if not enhanced_info and hasattr(self, 'enhanced_mapping_data') and self.enhanced_mapping_data:
                print(f"⚠ 未找到匹配的增强映射信息: {original_filename} (ID: {file_based_image_id})")
            
            # 如果没有从enhanced_mapping_data获取到信息，尝试从文件名中提取
            if not enhanced_info:
                # 尝试从文件名中提取product_id
                if new_filename.startswith('pdid'):
                    parts = new_filename.split('_')
                    if len(parts) > 0 and parts[0].startswith('pdid'):
                        product_id = parts[0][4:]  # 去掉"pdid"前缀
                        enhanced_info = {
                            'row_number': 0,
                            'device_name': '',
                            'product_id': product_id,
                            'dispimg_formula': '',
                            'cell_reference': ''
                        }
            
            original_image = {
                'image_id': image_id,  # 使用从enhanced_mapping_data获取的正确image_id
                'original_filename': original_filename,
                'original_path': original_path,
                'row_number': enhanced_info.get('row_number', 0) if enhanced_info else 0,
                'device_name': enhanced_info.get('device_name', '') if enhanced_info else '',
                'product_id': enhanced_info.get('product_id', '') if enhanced_info else '',
                'dispimg_formula': enhanced_info.get('dispimg_formula', '') if enhanced_info else '',
                'cell_reference': enhanced_info.get('cell_reference', '') if enhanced_info else ''
            }
            mapping_data['original_images'].append(original_image)
        
        # 构建保存图片信息
        for filename, file_path in saved_images.items():
            # 获取文件大小和哈希值
            file_size = 0
            file_hash = ''
            
            if os.path.exists(file_path):
                try:
                    file_size = os.path.getsize(file_path)
                    # 计算文件哈希
                    import hashlib
                    with open(file_path, 'rb') as f:
                        file_hash = hashlib.md5(f.read()).hexdigest()
                except Exception as e:
                    print(f"获取文件信息失败 {file_path}: {e}")
            
            saved_image = {
                'filename': filename,
                'file_path': file_path,
                'file_size': file_size,
                'file_hash': file_hash,
                'mapping_source': 'auto_rename'
            }
            mapping_data['saved_images'].append(saved_image)
        
        # 构建映射关系
        for new_filename, original_path in renamed_images.items():
            # 直接使用新文件名作为映射关系
            if new_filename in saved_images:
                # 获取正确的image_id，优先使用从enhanced_mapping_data获取的值
                original_filename = os.path.basename(original_path) if original_path else ''
                file_based_image_id = os.path.splitext(original_filename)[0] if original_filename else new_filename
                
                # 初始化映射关系所需的字段
                product_id = ''
                device_name = ''
                cell_reference = ''
                dispimg_formula = ''
                rid = ''
                
                # 如果文件名以ID_开头，直接使用它作为image_id
                if file_based_image_id.startswith('ID_'):
                    original_image_id = file_based_image_id
                else:
                    # 尝试从enhanced_mapping_data中获取正确的image_id
                    original_image_id = file_based_image_id
                    if hasattr(self, 'enhanced_mapping_data') and self.enhanced_mapping_data is not None:
                        for mapping_key, mapping_info in self.enhanced_mapping_data.items():
                            if mapping_info.get('actual_file', '') == original_filename:
                                original_image_id = mapping_info.get('image_id', file_based_image_id)
                                break
                
                # 从enhanced_mapping_data中获取额外的映射信息
                if hasattr(self, 'enhanced_mapping_data') and self.enhanced_mapping_data is not None:
                    # enhanced_mapping_data是一个字典，键是pdid，值是包含映射信息的字典
                    for pdid, mapping_info in self.enhanced_mapping_data.items():
                        # 检查image_id或actual_file是否匹配
                        if mapping_info.get('image_id', '') == original_image_id or mapping_info.get('actual_file', '') == original_filename:
                            product_id = mapping_info.get('pdid', '')  # 使用pdid作为product_id
                            device_name = mapping_info.get('device_name', '')
                            cell_reference = mapping_info.get('cell_reference', '')
                            dispimg_formula = mapping_info.get('dispimg_formula', '')
                            rid = mapping_info.get('r_id', '')  # 使用r_id作为rid
                            break
                
                # 如果没有从enhanced_mapping_data中获取到信息，尝试从文件名中提取product_id
                if not product_id and new_filename.startswith('pdid'):
                    # 从文件名中提取PDID，例如从"pdid5_yijian.png"中提取"5"
                    parts = new_filename.split('_')
                    if len(parts) > 0 and parts[0].startswith('pdid'):
                        product_id = parts[0][4:]  # 去掉"pdid"前缀
                
                # 确保real_image_file使用正确的路径格式和文件名
                # 使用renamed_images中的键（已转换为拼音的文件名）而不是saved_images中的值
                filename = os.path.basename(saved_images[new_filename])  # 获取实际保存的文件名
                # 检查实际保存的文件名是否包含汉字，如果是，则使用转换后的拼音文件名
                if any('\u4e00' <= char <= '\u9fff' for char in filename):
                    # 使用转换后的拼音文件名
                    real_image_file = f"images/{new_filename}"
                else:
                    # 使用实际保存的文件名
                    real_image_file = saved_images[new_filename]
                
                # 确保路径以images/开头
                if real_image_file and real_image_file.startswith('renamed_device_images/'):
                    # 替换路径前缀
                    real_image_file = real_image_file.replace('renamed_device_images/', 'images/')
                elif real_image_file and not real_image_file.startswith('images/'):
                    # 确保路径以images/开头
                    filename = os.path.basename(real_image_file)
                    real_image_file = f"images/{filename}"
                
                mapping_relationship = {
                    'product_id': product_id,
                    'device_name': device_name,
                    'cell_reference': cell_reference,
                    'dispimg_formula': dispimg_formula,
                    'image_id': original_image_id,  # 使用正确的image_id
                    'rid': rid,
                    'original_image_file': original_path,
                    'real_image_file': real_image_file,
                    'mapping_type': 'dispimg_formula' if dispimg_formula else 'direct',
                    'confidence': 1.0
                }
                mapping_data['mapping_relationships'].append(mapping_relationship)
        
        return mapping_data

    def _save_images(self) -> bool:
        """保存图片文件"""
        print("\n步骤3: 保存图片文件")
        print("-" * 40)
        
        try:
            # 获取重命名后的图片
            renamed_images = self.processing_results['renamed_images']
            
            if not renamed_images:
                print("✗ 没有可保存的图片")
                return False
            
            # 检查重复图片
            duplicates = self.saver.check_duplicate_images(renamed_images)
            
            if duplicates:
                print(f"发现 {len(duplicates)} 组重复图片，将进行去重处理")
            
            # 保存图片
            saved_images = self.saver.save_multiple_images(renamed_images, overwrite=True)
            
            if saved_images:
                self.processing_results['saved_images'] = saved_images
                self.processing_status['saving'] = True
                print(f"✓ 图片保存成功: {len(saved_images)} 张图片")
                
                # 构建完整的映射数据结构
                mapping_data = self._build_mapping_data(renamed_images, saved_images)
                
                # 生成图片映射关系文件
                print("生成图片映射关系文件...")
                mapping_file = self.mapping_generator.generate_mapping_file(mapping_data)
                print(f"映射关系文件已生成: {mapping_file}")
                
                return True
            else:
                print("✗ 图片保存失败")
                return False
                
        except Exception as e:
            error_msg = f"图片保存失败: {e}"
            self.processing_results['errors'].append(error_msg)
            print(error_msg)
            return False
    
    def _cleanup(self):
        """清理临时文件"""
        print("\n步骤4: 清理临时文件")
        print("-" * 40)
        
        try:
            # 清理提取器的临时文件
            self.extractor.cleanup()
            
            # 清理临时目录
            if os.path.exists(self.temp_dir):
                import shutil
                shutil.rmtree(self.temp_dir)
                print(f"✓ 临时文件已清理: {self.temp_dir}")
            
            self.processing_status['cleanup'] = True
            
        except Exception as e:
            error_msg = f"清理临时文件失败: {e}"
            self.processing_results['errors'].append(error_msg)
            print(error_msg)
    
    def _generate_report(self):
        """生成处理报告"""
        print("\n步骤5: 生成处理报告")
        print("-" * 40)
        
        try:
            # 创建报告目录
            report_dir = "processing_reports"
            if not os.path.exists(report_dir):
                os.makedirs(report_dir)
            
            # 生成报告文件名
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            report_file = os.path.join(report_dir, f"image_save_report_{timestamp}.txt")
            
            # 写入报告内容
            with open(report_file, 'w', encoding='utf-8') as f:
                f.write("设备图片保存流程处理报告\n")
                f.write("=" * 50 + "\n\n")
                
                f.write("处理状态:\n")
                for step, status in self.processing_status.items():
                    f.write(f"  {step}: {'✓' if status else '✗'}\n")
                f.write("\n")
                
                f.write("处理结果:\n")
                f.write(f"  提取图片: {len(self.processing_results['extracted_images'])} 张\n")
                f.write(f"  重命名图片: {len(self.processing_results['renamed_images'])} 张\n")
                f.write(f"  保存图片: {len(self.processing_results['saved_images'])} 张\n")
                f.write(f"  错误数量: {len(self.processing_results['errors'])} 个\n")
                f.write("\n")
                
                if self.processing_results['errors']:
                    f.write("错误详情:\n")
                    for error in self.processing_results['errors']:
                        f.write(f"  - {error}\n")
                f.write("\n")
                
                f.write("保存的图片文件:\n")
                for filename in self.processing_results['saved_images'].keys():
                    f.write(f"  - {filename}\n")
            
            print(f"✓ 处理报告已生成: {report_file}")
            
        except Exception as e:
            error_msg = f"生成处理报告失败: {e}"
            self.processing_results['errors'].append(error_msg)
            print(error_msg)
    
    def _convert_device_name_to_pinyin(self, device_name: str) -> str:
        """
        将设备名称转换为拼音
        
        Args:
            device_name: 设备名称
            
        Returns:
            str: 拼音字符串
        """
        if not device_name:
            return "device"
        
        try:
            # 扩展的拼音映射表（包含更多常见汉字）
            pinyin_map = {
                '智': 'zhi', '能': 'neng', '开': 'kai', '关': 'guan', '键': 'jian',
                '灯': 'deng', '具': 'ju', '窗': 'chuang', '帘': 'lian', '传': 'chuan',
                '感': 'gan', '器': 'qi', '家': 'jia', '电': 'dian', '门': 'men',
                '锁': 'suo', '中': 'zhong', '控': 'kong', '音': 'yin', '箱': 'xiang',
                '幕': 'mu', '墙': 'qiang', '壁': 'bi', '顶': 'ding', '外': 'wai',
                '人': 'ren', '造': 'zao', '贴': 'tie', '板': 'ban', '空': 'kong',
                '调': 'tiao', '插': 'cha', '座': 'zuo', '全': 'quan', '屋': 'wu',
                'WiFi': 'WiFi', 'wifi': 'wifi', 'WIFI': 'WIFI',
                '摄': 'she', '像': 'xiang', '头': 'tou', '监': 'jian', '视': 'shi',
                '报': 'bao', '警': 'jing', '烟': 'yan', '雾': 'wu', '气': 'qi',
                '温': 'wen', '度': 'du', '湿': 'shi', '度': 'du', '光': 'guang',
                '照': 'zhao', '明': 'ming', '风': 'feng', '扇': 'shan', '净': 'jing',
                '化': 'hua', '器': 'qi', '新': 'xin', '风': 'feng', '系': 'xi',
                '统': 'tong', '热': 're', '水': 'shui', '器': 'qi', '燃': 'ran',
                '气': 'qi', '阀': 'fa', '门': 'men', '机': 'ji', '器': 'qi',
                '人': 'ren', '扫': 'sao', '地': 'di', '机': 'ji', '器': 'qi',
                '晾': 'liang', '衣': 'yi', '架': 'jia', '电': 'dian', '动': 'dong',
                '窗': 'chuang', '帘': 'lian', '电': 'dian', '视': 'shi', '盒': 'he',
                '投': 'tou', '影': 'ying', '仪': 'yi', '音': 'yin', '响': 'xiang',
                '功': 'gong', '放': 'fang', '蓝': 'lan', '牙': 'ya', '音': 'yin',
                '箱': 'xiang', '智': 'zhi', '能': 'neng', '马': 'ma', '桶': 'tong',
                '智': 'zhi', '能': 'neng', '镜': 'jing', '子': 'zi', '智': 'zhi',
                '能': 'neng', '体': 'ti', '重': 'zhong', '秤': 'cheng', '智': 'zhi',
                '能': 'neng', '手': 'shou', '环': 'huan', '智': 'zhi', '能': 'neng',
                '门': 'men', '铃': 'ling', '智': 'zhi', '能': 'neng', '床': 'chuang',
                '智': 'zhi', '能': 'neng', '枕': 'zhen', '智': 'zhi', '能': 'neng',
                '花': 'hua', '洒': 'sa', '智': 'zhi', '能': 'neng', '喂': 'wei',
                '食': 'shi', '器': 'qi', '智': 'zhi', '能': 'neng', '垃': 'la',
                '圾': 'ji', '桶': 'tong', '智': 'zhi', '能': 'neng', '鞋': 'xie',
                '柜': 'gui', '智': 'zhi', '能': 'neng', '衣': 'yi', '柜': 'gui',
                '智': 'zhi', '能': 'neng', '茶': 'cha', '几': 'ji', '智': 'zhi',
                '能': 'neng', '座': 'zuo', '椅': 'yi', '智': 'zhi', '能': 'neng',
                '浴': 'yu', '室': 'shi', '镜': 'jing', '智': 'zhi', '能': 'neng',
                '花': 'hua', '洒': 'sa', '智': 'zhi', '能': 'neng', '马': 'ma',
                '桶': 'tong', '智': 'zhi', '能': 'neng', '浴': 'yu', '霸': 'ba',
                '智': 'zhi', '能': 'neng', '热': 're', '水': 'shui', '器': 'qi',
                '智': 'zhi', '能': 'neng', '空': 'kong', '气': 'qi', '净': 'jing',
                '化': 'hua', '器': 'qi', '智': 'zhi', '能': 'neng', '新': 'xin',
                '风': 'feng', '系': 'xi', '统': 'tong', '智': 'zhi', '能': 'neng',
                '除': 'chu', '湿': 'shi', '机': 'ji', '智': 'zhi', '能': 'neng',
                '加': 'jia', '湿': 'shi', '器': 'qi', '智': 'zhi', '能': 'neng',
                '空': 'kong', '调': 'tiao', '智': 'zhi', '能': 'neng', '地': 'di',
                '暖': 'nuan', '智': 'zhi', '能': 'neng', '开': 'kai', '关': 'guan',
                '智': 'zhi', '能': 'neng', '插': 'cha', '座': 'zuo', '智': 'zhi',
                '能': 'neng', '灯': 'deng', '智': 'zhi', '能': 'neng', '门': 'men',
                '锁': 'suo', '智': 'zhi', '能': 'neng', '窗': 'chuang', '帘': 'lian',
                '智': 'zhi', '能': 'neng', '晾': 'liang', '衣': 'yi', '架': 'jia',
                '智': 'zhi', '能': 'neng', '扫': 'sao', '地': 'di', '机': 'ji',
                '器': 'qi', '智': 'zhi', '能': 'neng', '拖': 'tuo', '把': 'ba',
                '智': 'zhi', '能': 'neng', '吸': 'xi', '尘': 'chen', '器': 'qi',
                '智': 'zhi', '能': 'neng', '擦': 'ca', '窗': 'chuang', '机': 'ji',
                '器': 'qi', '智': 'zhi', '能': 'neng', '洗': 'xi', '碗': 'wan',
                '机': 'ji', '智': 'zhi', '能': 'neng', '烘': 'hong', '干': 'gan',
                '机': 'ji', '智': 'zhi', '能': 'neng', '消': 'xiao', '毒': 'du',
                '柜': 'gui', '智': 'zhi', '能': 'neng', '煮': 'zhu', '饭': 'fan',
                '煲': 'bao', '智': 'zhi', '能': 'neng', '电': 'dian', '饭': 'fan',
                '煲': 'bao', '智': 'zhi', '能': 'neng', '压': 'ya', '力': 'li',
                '锅': 'guo', '智': 'zhi', '能': 'neng', '豆': 'dou', '浆': 'jiang',
                '机': 'ji', '智': 'zhi', '能': 'neng', '榨': 'zha', '汁': 'zhi',
                '机': 'ji', '智': 'zhi', '能': 'neng', '烤': 'kao', '箱': 'xiang',
                '智': 'zhi', '能': 'neng', '微': 'wei', '波': 'bo', '炉': 'lu',
                '智': 'zhi', '能': 'neng', '消': 'xiao', '毒': 'du', '柜': 'gui',
                '智': 'zhi', '能': 'neng', '洗': 'xi', '衣': 'yi', '机': 'ji',
                '智': 'zhi', '能': 'neng', '烘': 'hong', '干': 'gan', '机': 'ji',
                '智': 'zhi', '能': 'neng', '冰': 'bing', '箱': 'xiang', '智': 'zhi',
                '能': 'neng', '酒': 'jiu', '柜': 'gui', '智': 'zhi', '能': 'neng',
                '饮': 'yin', '水': 'shui', '机': 'ji', '智': 'zhi', '能': 'neng',
                '咖': 'ka', '啡': 'fei', '机': 'ji', '智': 'zhi', '能': 'neng',
                '面': 'mian', '包': 'bao', '机': 'ji', '智': 'zhi', '能': 'neng',
                '电': 'dian', '磁': 'ci', '炉': 'lu', '智': 'zhi', '能': 'neng',
                '电': 'dian', '饭': 'fan', '煲': 'bao', '智': 'zhi', '能': 'neng',
                '电': 'dian', '压': 'ya', '力': 'li', '锅': 'guo', '智': 'zhi',
                '能': 'neng', '电': 'dian', '烤': 'kao', '箱': 'xiang', '智': 'zhi',
                '能': 'neng', '电': 'dian', '热': 're', '水': 'shui', '器': 'qi',
                '智': 'zhi', '能': 'neng', '电': 'dian', '茶': 'cha', '壶': 'hu',
                '智': 'zhi', '能': 'neng', '电': 'dian', '炖': 'dun', '锅': 'guo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '煮': 'zhu', '锅': 'guo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '炒': 'chao', '锅': 'guo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '煎': 'jian', '锅': 'guo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '蒸': 'zheng', '锅': 'guo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '烧': 'shao', '烤': 'kao',
                '炉': 'lu', '智': 'zhi', '能': 'neng', '电': 'dian', '火': 'huo',
                '锅': 'guo', '智': 'zhi', '能': 'neng', '电': 'dian', '热': 're',
                '锅': 'guo', '智': 'zhi', '能': 'neng', '电': 'dian', '煎': 'jian',
                '烤': 'kao', '盘': 'pan', '智': 'zhi', '能': 'neng', '电': 'dian',
                '煮': 'zhu', '蛋': 'dan', '器': 'qi', '智': 'zhi', '能': 'neng',
                '电': 'dian', '炖': 'dun', '盅': 'zhong', '智': 'zhi', '能': 'neng',
                '电': 'dian', '蒸': 'zheng', '笼': 'long', '智': 'zhi', '能': 'neng',
                '电': 'dian', '暖': 'nuan', '锅': 'guo', '智': 'zhi', '能': 'neng',
                '电': 'dian', '热': 're', '杯': 'bei', '智': 'zhi', '能': 'neng',
                '电': 'dian', '搅': 'jiao', '拌': 'ban', '杯': 'bei', '智': 'zhi',
                '能': 'neng', '电': 'dian', '磨': 'mo', '豆': 'dou', '机': 'ji',
                '智': 'zhi', '能': 'neng', '电': 'dian', '榨': 'zha', '汁': 'zhi',
                '机': 'ji', '智': 'zhi', '能': 'neng', '电': 'dian', '切': 'qie',
                '菜': 'cai', '机': 'ji', '智': 'zhi', '能': 'neng', '电': 'dian',
                '和': 'he', '面': 'mian', '机': 'ji', '智': 'zhi', '能': 'neng',
                '电': 'dian', '打': 'da', '蛋': 'dan', '器': 'qi', '智': 'zhi',
                '能': 'neng', '电': 'dian', '绞': 'jiao', '肉': 'rou', '机': 'ji',
                '智': 'zhi', '能': 'neng', '电': 'dian', '切': 'qie', '肉': 'rou',
                '机': 'ji', '智': 'zhi', '能': 'neng', '电': 'dian', '切': 'qie',
                '菜': 'cai', '机': 'ji', '智': 'zhi', '能': 'neng', '电': 'dian',
                '洗': 'xi', '菜': 'cai', '机': 'ji', '智': 'zhi', '能': 'neng',
                '电': 'dian', '消': 'xiao', '毒': 'du', '柜': 'gui', '智': 'zhi',
                '能': 'neng', '电': 'dian', '烘': 'hong', '手': 'shou', '器': 'qi',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '餐': 'can', '盒': 'he', '智': 'zhi', '能': 'neng', '电': 'dian',
                '热': 're', '饭': 'fan', '盒': 'he', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '杯': 'bei', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '壶': 'hu',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '箱': 'xiang', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '袋': 'dai', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '坐': 'zuo', '垫': 'dian', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '鞋': 'xie',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '腰': 'yao', '托': 'tuo', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '手': 'shou', '套': 'tao', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '脚': 'jiao',
                '垫': 'dian', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '背': 'bei', '心': 'xin', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '膝': 'xi', '盖': 'gai',
                '垫': 'dian', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '颈': 'jing', '托': 'tuo', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '眼': 'yan', '罩': 'zhao',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '面': 'mian', '膜': 'mo', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '手': 'shou', '机': 'ji', '套': 'tao',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '平': 'ping', '板': 'ban', '套': 'tao', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '笔': 'bi', '记': 'ji',
                '本': 'ben', '套': 'tao', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '相': 'xiang', '机': 'ji', '套': 'tao',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '耳': 'er', '机': 'ji', '套': 'tao', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '移': 'yi', '动': 'dong',
                '电': 'dian', '源': 'yuan', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '移': 'yi', '动': 'dong', '电': 'dian',
                '源': 'yuan', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '充': 'chong', '电': 'dian', '宝': 'bao', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '充': 'chong',
                '电': 'dian', '宝': 'bao', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '手': 'shou', '机': 'ji', '支': 'zhi',
                '架': 'jia', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '平': 'ping', '板': 'ban', '支': 'zhi', '架': 'jia',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '笔': 'bi', '记': 'ji', '本': 'ben', '支': 'zhi', '架': 'jia',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '相': 'xiang', '机': 'ji', '支': 'zhi', '架': 'jia', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '耳': 'er',
                '机': 'ji', '支': 'zhi', '架': 'jia', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '手': 'shou', '机': 'ji',
                '座': 'zuo', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '平': 'ping', '板': 'ban', '座': 'zuo', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '笔': 'bi',
                '记': 'ji', '本': 'ben', '座': 'zuo', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '相': 'xiang', '机': 'ji',
                '座': 'zuo', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '耳': 'er', '机': 'ji', '座': 'zuo', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '手': 'shou',
                '机': 'ji', '夹': 'jia', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '平': 'ping', '板': 'ban', '夹': 'jia',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '笔': 'bi', '记': 'ji', '本': 'ben', '夹': 'jia', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '相': 'xiang',
                '机': 'ji', '夹': 'jia', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '耳': 'er', '机': 'ji', '夹': 'jia',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '手': 'shou', '机': 'ji', '套': 'tao', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '平': 'ping', '板': 'ban',
                '套': 'tao', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '笔': 'bi', '记': 'ji', '本': 'ben', '套': 'tao',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '相': 'xiang', '机': 'ji', '套': 'tao', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '耳': 'er', '机': 'ji',
                '套': 'tao', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '手': 'shou', '机': 'ji', '支': 'zhi', '架': 'jia',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '平': 'ping', '板': 'ban', '支': 'zhi', '架': 'jia', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '笔': 'bi',
                '记': 'ji', '本': 'ben', '支': 'zhi', '架': 'jia', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '相': 'xiang',
                '机': 'ji', '支': 'zhi', '架': 'jia', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '耳': 'er', '机': 'ji',
                '支': 'zhi', '架': 'jia', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '手': 'shou', '机': 'ji', '座': 'zuo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '平': 'ping', '板': 'ban', '座': 'zuo', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '笔': 'bi', '记': 'ji',
                '本': 'ben', '座': 'zuo', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '相': 'xiang', '机': 'ji', '座': 'zuo',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '耳': 'er', '机': 'ji', '座': 'zuo', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '手': 'shou', '机': 'ji',
                '夹': 'jia', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '平': 'ping', '板': 'ban', '夹': 'jia', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '笔': 'bi',
                '记': 'ji', '本': 'ben', '夹': 'jia', '智': 'zhi', '能': 'neng',
                '电': 'dian', '保': 'bao', '温': 'wen', '相': 'xiang', '机': 'ji',
                '夹': 'jia', '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao',
                '温': 'wen', '耳': 'er', '机': 'ji', '夹': 'jia', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '手': 'shou',
                '机': 'ji', '套': 'tao', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '平': 'ping', '板': 'ban', '套': 'tao',
                '智': 'zhi', '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen',
                '笔': 'bi', '记': 'ji', '本': 'ben', '套': 'tao', '智': 'zhi',
                '能': 'neng', '电': 'dian', '保': 'bao', '温': 'wen', '相': 'xiang',
                '机': 'ji', '套': 'tao', '智': 'zhi', '能': 'neng', '电': 'dian',
                '保': 'bao', '温': 'wen', '耳': 'er', '机': 'ji', '套': 'tao'
            }
            
            # 使用pypinyin库进行准确的拼音转换
            from pypinyin import pinyin, Style
            
            # 处理混合文本（中文+数字+字母）
            pinyin_result = []
            
            # 逐个字符处理
            for char in device_name:
                if '\u4e00' <= char <= '\u9fff':  # 中文字符
                    # 使用pypinyin转换为拼音
                    char_pinyin = pinyin(char, style=Style.NORMAL)
                    if char_pinyin:
                        pinyin_result.append(char_pinyin[0][0])
                elif char.isalnum():  # 字母或数字
                    pinyin_result.append(char.lower())
                elif char in (' ', '-', '_'):  # 分隔符
                    pinyin_result.append('_')
                else:
                    # 其他字符跳过
                    continue
            
            # 合并拼音结果
            pinyin_str = ''.join(pinyin_result)
            
            # 如果转换失败或结果为空，使用备选方案
            if not pinyin_str:
                simplified = re.sub(r'[^a-zA-Z0-9]', '_', device_name)
                pinyin_str = simplified.lower()
            
            # 清理结果：移除多余的下划线，确保格式规范
            pinyin_str = re.sub(r'_+', '_', pinyin_str)
            pinyin_str = pinyin_str.strip('_')
            
            # 如果结果仍然包含中文字符，使用更严格的转换
            if re.search(r'[\u4e00-\u9fff]', pinyin_str):
                print(f"警告：拼音转换结果仍包含中文字符: {device_name} -> {pinyin_str}")
                # 使用备选方案：移除所有非ASCII字符
                pinyin_str = re.sub(r'[^a-zA-Z0-9_]', '_', device_name)
                pinyin_str = re.sub(r'_+', '_', pinyin_str)
                pinyin_str = pinyin_str.strip('_').lower()
            
            return pinyin_str
            
        except Exception as e:
            print(f"设备名称转拼音失败 {device_name}: {e}")
            # 返回简化版本作为备选
            safe_name = "".join(c for c in device_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            return safe_name.replace(' ', '_').lower()

    def get_processing_summary(self) -> Dict:
        """
        获取处理摘要
        
        Returns:
            Dict: 处理摘要信息
        """
        return {
            'status': self.processing_status,
            'results': {
                'extracted_count': len(self.processing_results['extracted_images']),
                'renamed_count': len(self.processing_results['renamed_images']),
                'saved_count': len(self.processing_results['saved_images']),
                'error_count': len(self.processing_results['errors'])
            },
            'errors': self.processing_results['errors']
        }


def main():
    """主函数 - 测试设备图片保存流程"""
    excel_file = "../智能家居模具库.xlsx"
    
    if os.path.exists(excel_file):
        print("开始测试设备图片保存流程...")
        
        # 创建控制器
        controller = ImageSaveController(excel_file, "test_device_images")
        
        # 运行完整流程
        success = controller.run_complete_workflow()
        
        # 输出处理摘要
        summary = controller.get_processing_summary()
        print("\n处理摘要:")
        print(f"  成功: {'是' if success else '否'}")
        print(f"  提取图片: {summary['results']['extracted_count']}")
        print(f"  重命名图片: {summary['results']['renamed_count']}")
        print(f"  保存图片: {summary['results']['saved_count']}")
        print(f"  错误数量: {summary['results']['error_count']}")
        
        if not success and summary['errors']:
            print("\n错误详情:")
            for error in summary['errors']:
                print(f"  - {error}")
    else:
        print(f"Excel文件不存在: {excel_file}")


if __name__ == "__main__":
    main()