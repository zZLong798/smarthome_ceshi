"""
图片路径解析器
根据PDID从映像关系表中查找真实图片路径
"""

import json
import os
from typing import Optional


class ImagePathResolver:
    """图片路径解析器类"""
    
    def __init__(self, mapping_file_path: str = "../images/image_mapping.json"):
        """
        初始化图片路径解析器
        
        Args:
            mapping_file_path: 映像关系表文件路径
        """
        self.mapping_file_path = mapping_file_path
        self.mapping_data = None
        self._load_mapping_data()
    
    def _load_mapping_data(self):
        """加载映像关系表数据"""
        try:
            if os.path.exists(self.mapping_file_path):
                with open(self.mapping_file_path, 'r', encoding='utf-8') as f:
                    self.mapping_data = json.load(f)
                print(f"[OK] 成功加载映像关系表: {self.mapping_file_path}")
            else:
                print(f"[WARN] 映像关系表文件不存在: {self.mapping_file_path}")
                self.mapping_data = None
        except Exception as e:
            print(f"[ERROR] 加载映像关系表失败: {e}")
            self.mapping_data = None
    
    def get_image_path_by_pdid(self, pdid: str) -> Optional[str]:
        """
        根据PDID获取真实图片路径
        
        Args:
            pdid: 产品ID
            
        Returns:
            真实图片路径，如果找不到则返回None
        """
        if not self.mapping_data:
            print(f"⚠️ 映像关系表未加载，无法查找PDID: {pdid}")
            return None
        
        # 在mapping_relationships中查找
        if "mapping_relationships" in self.mapping_data:
            for mapping in self.mapping_data["mapping_relationships"]:
                if mapping.get("product_id") == pdid:
                    real_path = mapping.get("real_image_file")
                    if real_path:
                        # 确保路径是绝对路径
                        if not os.path.isabs(real_path):
                            # 映像关系表文件在images目录下，real_image_file已经是相对images目录的路径
                            # 所以需要从项目根目录开始构建路径
                            project_root = os.path.dirname(os.path.dirname(self.mapping_file_path))
                            real_path = os.path.join(project_root, real_path)
                        
                        if os.path.exists(real_path):
                            print(f"✅ 找到PDID {pdid} 对应的图片: {real_path}")
                            return real_path
                        else:
                            print(f"⚠️ PDID {pdid} 对应的图片路径不存在: {real_path}")
                            return None
        
        print(f"❌ 未找到PDID {pdid} 对应的图片")
        return None
    
    def get_image_path_by_device_name(self, device_name: str) -> Optional[str]:
        """
        根据设备名称获取真实图片路径
        
        Args:
            device_name: 设备名称
            
        Returns:
            真实图片路径，如果找不到则返回None
        """
        if not self.mapping_data:
            print(f"[WARN] 映像关系表未加载，无法查找设备: {device_name}")
            return None
        
        # 在mapping_relationships中查找
        if "mapping_relationships" in self.mapping_data:
            for mapping in self.mapping_data["mapping_relationships"]:
                if mapping.get("device_name") == device_name:
                    real_path = mapping.get("real_image_file")
                    if real_path:
                        # 确保路径是绝对路径
                        if not os.path.isabs(real_path):
                            # 映像关系表文件在images目录下，real_image_file已经是相对images目录的路径
                            # 所以需要从项目根目录开始构建路径
                            project_root = os.path.dirname(os.path.dirname(self.mapping_file_path))
                            real_path = os.path.join(project_root, real_path)
                        
                        if os.path.exists(real_path):
                            print(f"[OK] 找到设备 {device_name} 对应的图片: {real_path}")
                            return real_path
                        else:
                            print(f"[WARN] 设备 {device_name} 对应的图片路径不存在: {real_path}")
                            return None
        
        print(f"[ERROR] 未找到设备 {device_name} 对应的图片")
        return None


def get_image_path(pdid: str = None, device_name: str = None) -> Optional[str]:
    """
    便捷函数：根据PDID或设备名称获取图片路径
    
    Args:
        pdid: 产品ID（优先使用）
        device_name: 设备名称
        
    Returns:
        真实图片路径，如果找不到则返回None
    """
    resolver = ImagePathResolver()
    
    if pdid:
        return resolver.get_image_path_by_pdid(pdid)
    elif device_name:
        return resolver.get_image_path_by_device_name(device_name)
    else:
        print("[ERROR] 必须提供PDID或设备名称")
        return None


# 测试函数
if __name__ == "__main__":
    # 测试PDID查找
    test_pdid = "1"
    path = get_image_path(pdid=test_pdid)
    print(f"PDID {test_pdid} 对应的图片路径: {path}")
    
    # 测试设备名称查找
    test_device = "一键"
    path = get_image_path(device_name=test_device)
    print(f"设备 {test_device} 对应的图片路径: {path}")