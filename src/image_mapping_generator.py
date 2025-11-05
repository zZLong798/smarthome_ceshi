#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片映射关系文件生成器
生成原Excel图片信息与保存图片的对应关系映射文件
"""

import os
import json
import csv
from typing import Dict, List, Any
from datetime import datetime


class ImageMappingGenerator:
    """图片映射关系文件生成器"""
    
    def __init__(self, excel_file_path: str, images_dir: str):
        """
        初始化映射生成器
        
        Args:
            excel_file_path: Excel文件路径
            images_dir: 图片保存目录
        """
        self.excel_file_path = excel_file_path
        self.images_dir = images_dir
        self.mapping_file_path = os.path.join(images_dir, "image_mapping.json")
        self.csv_mapping_file_path = os.path.join(images_dir, "image_mapping.csv")
        
    def generate_mapping_file(self, image_data: Dict[str, Any]) -> bool:
        """
        生成图片映射关系文件
        
        Args:
            image_data: 图片数据字典，包含以下结构：
                {
                    'original_images': [
                        {
                            'image_id': '图片ID',
                            'original_filename': '原文件名',
                            'row_number': 行号,
                            'device_name': '设备名称',
                            'product_id': '产品ID',
                            'dispimg_formula': 'DISPIMG公式'
                        }
                    ],
                    'saved_images': [
                        {
                            'filename': '保存的文件名',
                            'file_path': '完整文件路径',
                            'file_size': 文件大小,
                            'file_hash': '文件哈希值',
                            'mapping_source': '映射来源'
                        }
                    ],
                    'mapping_relationships': [
                        {
                            'original_image_id': '原图片ID',
                            'saved_filename': '保存的文件名',
                            'mapping_type': '映射类型',
                            'confidence': '映射置信度'
                        }
                    ]
                }
                
        Returns:
            bool: 是否成功生成
        """
        try:
            # 确保images目录存在
            if not os.path.exists(self.images_dir):
                os.makedirs(self.images_dir)
            
            # 添加元数据
            mapping_data = {
                'metadata': {
                    'generated_at': datetime.now().isoformat(),
                    'excel_file': os.path.basename(self.excel_file_path),
                    'images_directory': self.images_dir,
                    'total_original_images': len(image_data.get('original_images', [])),
                    'total_saved_images': len(image_data.get('saved_images', [])),
                    'total_mappings': len(image_data.get('mapping_relationships', []))
                },
                **image_data
            }
            
            # 保存JSON格式的映射文件
            with open(self.mapping_file_path, 'w', encoding='utf-8') as f:
                json.dump(mapping_data, f, ensure_ascii=False, indent=2)
            
            print(f"✓ JSON映射文件已生成: {self.mapping_file_path}")
            
            # 同时生成CSV格式的映射文件（便于查看）
            self._generate_csv_mapping(image_data)
            
            return True
            
        except Exception as e:
            print(f"✗ 生成映射文件失败: {e}")
            return False
    
    def _generate_csv_mapping(self, image_data: Dict[str, Any]) -> bool:
        """生成CSV格式的映射文件"""
        try:
            with open(self.csv_mapping_file_path, 'w', encoding='utf-8', newline='') as csvfile:
                writer = csv.writer(csvfile)
                
                # 写入CSV头部，添加cell_reference字段
                writer.writerow([
                    '产品ID', '设备名称', '原图片ID', '原文件名', '单元格引用',
                    '保存文件名', '文件路径', '文件大小', '映射类型', '置信度'
                ])
                
                # 写入映射关系数据
                for mapping in image_data.get('mapping_relationships', []):
                    # 查找对应的原图片信息
                    original_image = None
                    for orig_img in image_data.get('original_images', []):
                        if orig_img.get('image_id') == mapping.get('image_id', ''):
                            original_image = orig_img
                            break
                    
                    # 查找对应的保存图片信息
                    saved_image = None
                    saved_filename = os.path.basename(mapping.get('real_image_file', ''))
                    for saved_img in image_data.get('saved_images', []):
                        if saved_img.get('filename', '') == saved_filename:
                            saved_image = saved_img
                            break
                    
                    writer.writerow([
                        mapping.get('product_id', ''),
                        mapping.get('device_name', ''),
                        mapping.get('image_id', ''),
                        original_image.get('original_filename', '') if original_image else '',
                        mapping.get('cell_reference', ''),
                        saved_image.get('filename', '') if saved_image else '',
                        saved_image.get('file_path', '') if saved_image else '',
                        saved_image.get('file_size', '') if saved_image else '',
                        mapping.get('mapping_type', ''),
                        mapping.get('confidence', '')
                    ])
            
            print(f"✓ CSV映射文件已生成: {self.csv_mapping_file_path}")
            return True
            
        except Exception as e:
            print(f"✗ 生成CSV映射文件失败: {e}")
            return False
    
    def load_mapping_file(self) -> Dict[str, Any]:
        """
        加载图片映射关系文件
        
        Returns:
            Dict[str, Any]: 映射数据
        """
        try:
            if not os.path.exists(self.mapping_file_path):
                print(f"映射文件不存在: {self.mapping_file_path}")
                return {}
            
            with open(self.mapping_file_path, 'r', encoding='utf-8') as f:
                mapping_data = json.load(f)
            
            print(f"✓ 映射文件已加载: {self.mapping_file_path}")
            return mapping_data
            
        except Exception as e:
            print(f"✗ 加载映射文件失败: {e}")
            return {}
    
    def get_image_by_pdid(self, pdid: str) -> str:
        """
        根据PDID获取对应的图片路径
        
        Args:
            pdid: 产品ID
            
        Returns:
            str: 图片文件路径，如果未找到返回空字符串
        """
        mapping_data = self.load_mapping_file()
        
        if not mapping_data:
            return ""
        
        # 首先尝试在映射关系中查找
        for mapping in mapping_data.get('mapping_relationships', []):
            original_image_id = mapping.get('original_image_id', '')
            
            # 查找对应的原图片信息
            for orig_img in mapping_data.get('original_images', []):
                if (orig_img.get('image_id') == original_image_id and 
                    orig_img.get('product_id') == pdid):
                    
                    # 返回保存的图片路径
                    saved_filename = mapping.get('saved_filename', '')
                    for saved_img in mapping_data.get('saved_images', []):
                        if saved_img.get('filename') == saved_filename:
                            return saved_img.get('file_path', '')
        
        # 如果在映射关系中找不到，直接在saved_images中查找文件名匹配的图片
        for saved_img in mapping_data.get('saved_images', []):
            filename = saved_img.get('filename', '')
            # 检查文件名是否以pdid{pdid}_开头
            if filename.startswith(f'pdid{pdid}_'):
                return saved_img.get('file_path', '')
        
        return ""
    
    def get_all_mappings(self) -> List[Dict[str, str]]:
        """
        获取所有映射关系
        
        Returns:
            List[Dict[str, str]]: 映射关系列表
        """
        mapping_data = self.load_mapping_file()
        
        if not mapping_data:
            return []
        
        mappings = []
        for mapping in mapping_data.get('mapping_relationships', []):
            original_image_id = mapping.get('original_image_id', '')
            saved_filename = mapping.get('saved_filename', '')
            
            # 查找原图片信息
            original_info = {}
            for orig_img in mapping_data.get('original_images', []):
                if orig_img.get('image_id') == original_image_id:
                    original_info = orig_img
                    break
            
            # 查找保存图片信息
            saved_info = {}
            for saved_img in mapping_data.get('saved_images', []):
                if saved_img.get('filename') == saved_filename:
                    saved_info = saved_img
                    break
            
            mappings.append({
                'pdid': original_info.get('product_id', ''),
                'device_name': original_info.get('device_name', ''),
                'original_image_id': original_image_id,
                'original_filename': original_info.get('original_filename', ''),
                'cell_reference': original_info.get('cell_reference', ''),  # 添加cell_reference字段
                'saved_filename': saved_filename,
                'file_path': saved_info.get('file_path', ''),
                'mapping_type': mapping.get('mapping_type', ''),
                'confidence': mapping.get('confidence', '')
            })
        
        return mappings


def create_sample_mapping_data() -> Dict[str, Any]:
    """创建示例映射数据"""
    return {
        'original_images': [
            {
                'image_id': 'image1',
                'original_filename': 'image1.png',
                'row_number': 2,
                'device_name': '一键智能开关',
                'product_id': '1',
                'dispimg_formula': '=_xlfn.DISPIMG("ID_165225A72C6A443A9B253A3B8E11BA45",1)',
                'cell_reference': 'L37'  # 添加cell_reference字段
            },
            {
                'image_id': 'image2', 
                'original_filename': 'image2.png',
                'row_number': 3,
                'device_name': '二键智能开关',
                'product_id': '2',
                'dispimg_formula': '=_xlfn.DISPIMG("ID_265225A72C6A443A9B253A3B8E11BA45",1)',
                'cell_reference': 'L38'  # 添加cell_reference字段
            }
        ],
        'saved_images': [
            {
                'filename': 'pdid1_yijian.png',
                'file_path': '/path/to/images/pdid1_yijian.png',
                'file_size': 10240,
                'file_hash': 'abc123',
                'mapping_source': 'auto_rename'
            },
            {
                'filename': 'pdid2_erjian.png',
                'file_path': '/path/to/images/pdid2_erjian.png',
                'file_size': 15360,
                'file_hash': 'def456',
                'mapping_source': 'auto_rename'
            }
        ],
        'mapping_relationships': [
            {
                'original_image_id': 'image1',
                'saved_filename': 'pdid1_yijian.png',
                'mapping_type': 'direct',
                'confidence': 1.0
            },
            {
                'original_image_id': 'image2',
                'saved_filename': 'pdid2_erjian.png',
                'mapping_type': 'direct',
                'confidence': 1.0
            }
        ]
    }


if __name__ == "__main__":
    # 测试映射生成器
    excel_file = "智能家居模具库.xlsx"
    images_dir = "images"
    
    generator = ImageMappingGenerator(excel_file, images_dir)
    
    # 生成示例映射文件
    sample_data = create_sample_mapping_data()
    success = generator.generate_mapping_file(sample_data)
    
    if success:
        print("映射文件生成成功!")
        
        # 测试加载和查询功能
        mappings = generator.get_all_mappings()
        print(f"加载到 {len(mappings)} 个映射关系")
        
        for mapping in mappings:
            print(f"PDID {mapping['pdid']}: {mapping['device_name']} -> {mapping['saved_filename']}")
    else:
        print("映射文件生成失败!")