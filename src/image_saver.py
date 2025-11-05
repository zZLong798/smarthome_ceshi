#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片保存器模块
管理图片的保存、验证和清理
"""

import os
import shutil
from typing import Dict, List, Optional
from PIL import Image
import hashlib


class ImageSaver:
    """图片保存器类"""
    
    def __init__(self, base_dir: str = "images"):
        """
        初始化图片保存器
        
        Args:
            base_dir: 基础保存目录
        """
        self.base_dir = base_dir
        self.saved_images: Dict[str, str] = {}  # 文件名到文件路径的映射
        
        # 确保基础目录存在
        if not os.path.exists(self.base_dir):
            os.makedirs(self.base_dir)
    
    def validate_image(self, image_path: str) -> bool:
        """
        验证图片文件是否有效
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            bool: 是否有效
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(image_path):
                print(f"图片文件不存在: {image_path}")
                return False
            
            # 检查文件大小
            file_size = os.path.getsize(image_path)
            if file_size == 0:
                print(f"图片文件为空: {image_path}")
                return False
            
            # 尝试用PIL打开图片
            with Image.open(image_path) as img:
                # 验证图片格式
                img.verify()
            
            print(f"图片验证通过: {image_path} ({file_size} bytes)")
            return True
            
        except Exception as e:
            print(f"图片验证失败 {image_path}: {e}")
            return False
    
    def calculate_image_hash(self, image_path: str) -> Optional[str]:
        """
        计算图片文件的哈希值
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            str: 图片哈希值，失败返回None
        """
        try:
            with open(image_path, 'rb') as f:
                file_hash = hashlib.md5()
                while chunk := f.read(8192):
                    file_hash.update(chunk)
                return file_hash.hexdigest()
        except Exception as e:
            print(f"计算图片哈希失败 {image_path}: {e}")
            return None
    
    def save_image(self, source_path: str, target_filename: str, 
                   overwrite: bool = False) -> Optional[str]:
        """
        保存图片文件
        
        Args:
            source_path: 源文件路径
            target_filename: 目标文件名
            overwrite: 是否覆盖已存在的文件
            
        Returns:
            str: 保存后的文件路径，失败返回None
        """
        try:
            # 验证源图片
            if not self.validate_image(source_path):
                print(f"源图片验证失败，无法保存: {source_path}")
                return None
            
            # 构建目标路径
            target_path = os.path.join(self.base_dir, target_filename)
            
            # 检查目标文件是否已存在
            if os.path.exists(target_path) and not overwrite:
                print(f"目标文件已存在，跳过保存: {target_path}")
                return target_path
            
            # 复制文件
            shutil.copy2(source_path, target_path)
            
            # 验证目标文件
            if not self.validate_image(target_path):
                print(f"目标图片验证失败，删除文件: {target_path}")
                os.remove(target_path)
                return None
            
            # 记录保存的图片
            self.saved_images[target_filename] = target_path
            
            print(f"图片保存成功: {source_path} -> {target_path}")
            return target_path
            
        except Exception as e:
            print(f"保存图片失败: {source_path} -> {target_filename}, 错误: {e}")
            return None
    
    def save_multiple_images(self, image_mapping: Dict[str, str], 
                            overwrite: bool = False) -> Dict[str, str]:
        """
        批量保存图片
        
        Args:
            image_mapping: 文件名到源文件路径的映射
            overwrite: 是否覆盖已存在的文件
            
        Returns:
            Dict[str, str]: 成功保存的文件名到路径的映射
        """
        saved_files = {}
        
        try:
            print(f"开始批量保存 {len(image_mapping)} 张图片...")
            
            for filename, source_path in image_mapping.items():
                saved_path = self.save_image(source_path, filename, overwrite)
                if saved_path:
                    saved_files[filename] = saved_path
            
            print(f"批量保存完成: 成功 {len(saved_files)}/{len(image_mapping)} 张图片")
            return saved_files
            
        except Exception as e:
            print(f"批量保存图片失败: {e}")
            return {}
    
    def check_duplicate_images(self, image_mapping: Dict[str, str]) -> Dict[str, List[str]]:
        """
        检查重复图片
        
        Args:
            image_mapping: 文件名到源文件路径的映射
            
        Returns:
            Dict[str, List[str]]: 哈希值到文件列表的映射
        """
        hash_mapping = {}
        
        try:
            print("开始检查重复图片...")
            
            for filename, source_path in image_mapping.items():
                # 计算图片哈希
                file_hash = self.calculate_image_hash(source_path)
                if file_hash:
                    if file_hash not in hash_mapping:
                        hash_mapping[file_hash] = []
                    hash_mapping[file_hash].append(filename)
            
            # 找出重复的图片
            duplicates = {}
            for file_hash, filenames in hash_mapping.items():
                if len(filenames) > 1:
                    duplicates[file_hash] = filenames
            
            if duplicates:
                print(f"发现 {len(duplicates)} 组重复图片")
                for file_hash, filenames in duplicates.items():
                    print(f"  哈希 {file_hash[:8]}...: {filenames}")
            else:
                print("未发现重复图片")
            
            return duplicates
            
        except Exception as e:
            print(f"检查重复图片失败: {e}")
            return {}
    
    def cleanup_old_images(self, keep_filenames: List[str]) -> int:
        """
        清理旧图片文件
        
        Args:
            keep_filenames: 需要保留的文件名列表
            
        Returns:
            int: 删除的文件数量
        """
        try:
            deleted_count = 0
            
            # 获取当前目录中的所有图片文件
            image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
            
            for filename in os.listdir(self.base_dir):
                file_path = os.path.join(self.base_dir, filename)
                
                # 检查是否是图片文件
                if os.path.isfile(file_path) and any(filename.lower().endswith(ext) for ext in image_extensions):
                    # 检查是否需要保留
                    if filename not in keep_filenames:
                        try:
                            os.remove(file_path)
                            deleted_count += 1
                            print(f"删除旧图片: {filename}")
                        except Exception as e:
                            print(f"删除文件失败 {filename}: {e}")
            
            print(f"清理完成: 删除了 {deleted_count} 个旧图片文件")
            return deleted_count
            
        except Exception as e:
            print(f"清理旧图片失败: {e}")
            return 0
    
    def get_saved_images_info(self) -> Dict[str, Dict]:
        """
        获取已保存图片的详细信息
        
        Returns:
            Dict[str, Dict]: 文件名到图片信息的映射
        """
        image_info = {}
        
        try:
            for filename, file_path in self.saved_images.items():
                if os.path.exists(file_path):
                    file_size = os.path.getsize(file_path)
                    file_hash = self.calculate_image_hash(file_path)
                    
                    image_info[filename] = {
                        'path': file_path,
                        'size': file_size,
                        'hash': file_hash,
                        'valid': self.validate_image(file_path)
                    }
            
            return image_info
            
        except Exception as e:
            print(f"获取图片信息失败: {e}")
            return {}
    
    def backup_images(self, backup_dir: str = "backup_images") -> bool:
        """
        备份图片文件
        
        Args:
            backup_dir: 备份目录
            
        Returns:
            bool: 是否备份成功
        """
        try:
            # 确保备份目录存在
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            # 备份所有已保存的图片
            backup_count = 0
            for filename, file_path in self.saved_images.items():
                if os.path.exists(file_path):
                    backup_path = os.path.join(backup_dir, filename)
                    shutil.copy2(file_path, backup_path)
                    backup_count += 1
            
            print(f"图片备份完成: 备份了 {backup_count} 张图片到 {backup_dir}")
            return True
            
        except Exception as e:
            print(f"图片备份失败: {e}")
            return False


def main():
    """测试函数"""
    # 测试图片保存器
    saver = ImageSaver("test_images")
    
    # 测试图片验证
    test_image = "test.png"
    if os.path.exists(test_image):
        is_valid = saver.validate_image(test_image)
        print(f"图片验证结果: {is_valid}")
        
        # 测试保存图片
        saved_path = saver.save_image(test_image, "test_saved.png")
        print(f"图片保存结果: {saved_path}")
    else:
        print("测试图片不存在")


if __name__ == "__main__":
    main()