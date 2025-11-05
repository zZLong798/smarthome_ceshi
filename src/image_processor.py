"""
图片处理模块 - 负责设备图片的提取和尺寸调整
"""

import os
from PIL import Image
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU

class ImageProcessor:
    """图片处理器类"""
    
    def __init__(self, image_base_path="../assets/images"):
        """
        初始化图片处理器
        
        Args:
            image_base_path: 图片资源基础路径
        """
        self.image_base_path = image_base_path
        self.target_width_cm = 0.9  # 目标宽度（厘米）
        self.target_height_cm = 0.9  # 目标高度（厘米）
        
        # 设备名称到图片文件的映射
        self.device_image_mapping = {
            "一键智能开关": "switches/一键.png",
            "二键智能开关": "switches/二键.png", 
            "三键智能开关": "switches/三键.png",
            "四键智能开关": "switches/四键.png",
            "领普二键智能开关": "switches/二键.png",
            "易来四键智能开关": "switches/四键.png"
        }
    
    def get_device_image_path(self, device_name):
        """
        根据设备名称获取图片路径
        
        Args:
            device_name: 设备名称
            
        Returns:
            str: 图片文件路径，如果不存在返回None
        """
        # 尝试精确匹配
        if device_name in self.device_image_mapping:
            image_relative_path = self.device_image_mapping[device_name]
            image_path = os.path.join(self.image_base_path, image_relative_path)
            if os.path.exists(image_path):
                return image_path
        
        # 尝试模糊匹配
        for key, value in self.device_image_mapping.items():
            if key in device_name:
                image_relative_path = value
                image_path = os.path.join(self.image_base_path, image_relative_path)
                if os.path.exists(image_path):
                    return image_path
        
        return None
    
    def resize_image_to_cm(self, image_path, target_width_cm=None, target_height_cm=None):
        """
        将图片调整为指定厘米尺寸
        
        Args:
            image_path: 原始图片路径
            target_width_cm: 目标宽度（厘米）
            target_height_cm: 目标高度（厘米）
            
        Returns:
            PIL.Image: 调整后的图片对象
        """
        if target_width_cm is None:
            target_width_cm = self.target_width_cm
        if target_height_cm is None:
            target_height_cm = self.target_height_cm
        
        # 打开原始图片
        original_image = Image.open(image_path)
        
        # 计算目标像素尺寸（假设96 DPI）
        dpi = 96
        target_width_px = int(target_width_cm * dpi / 2.54)
        target_height_px = int(target_height_cm * dpi / 2.54)
        
        # 调整图片尺寸
        resized_image = original_image.resize((target_width_px, target_height_px), Image.Resampling.LANCZOS)
        
        return resized_image
    
    def create_excel_image(self, device_name, temp_dir="temp_images"):
        """
        为Excel创建图片对象
        
        Args:
            device_name: 设备名称
            temp_dir: 临时文件目录
            
        Returns:
            ExcelImage: Excel图片对象，如果图片不存在返回None
        """
        # 获取图片路径
        image_path = self.get_device_image_path(device_name)
        if not image_path:
            return None
        
        # 创建临时目录
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        # 调整图片尺寸
        resized_image = self.resize_image_to_cm(image_path)
        
        # 保存临时文件
        temp_image_path = os.path.join(temp_dir, f"{device_name.replace(' ', '_')}.png")
        resized_image.save(temp_image_path)
        
        # 创建Excel图片对象
        excel_image = ExcelImage(temp_image_path)
        
        # 设置图片尺寸（转换为EMU单位）
        excel_image.width = cm_to_EMU(self.target_width_cm)
        excel_image.height = cm_to_EMU(self.target_height_cm)
        
        return excel_image
    
    def cleanup_temp_files(self, temp_dir="../temp_images"):
        """
        清理临时图片文件
        
        Args:
            temp_dir: 临时文件目录
        """
        if os.path.exists(temp_dir):
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)
            os.rmdir(temp_dir)


def test_image_processor():
    """测试图片处理器"""
    processor = ImageProcessor()
    
    # 测试设备图片路径获取
    test_devices = ["一键智能开关", "二键智能开关", "三键智能开关", "四键智能开关", "领普二键智能开关", "易来四键智能开关"]
    
    print("设备图片路径测试:")
    for device in test_devices:
        path = processor.get_device_image_path(device)
        if path:
            print(f"✓ {device}: {path}")
        else:
            print(f"✗ {device}: 图片不存在")
    
    # 测试图片尺寸调整
    print("\\n图片尺寸调整测试:")
    test_device = "二键智能开关"
    path = processor.get_device_image_path(test_device)
    if path:
        resized_image = processor.resize_image_to_cm(path)
        print(f"原始尺寸: {Image.open(path).size}")
        print(f"调整后尺寸: {resized_image.size}")
        print(f"目标尺寸: {processor.target_width_cm}cm × {processor.target_height_cm}cm")
    
    # 测试Excel图片创建
    print("\\nExcel图片创建测试:")
    excel_image = processor.create_excel_image(test_device)
    if excel_image:
        print(f"Excel图片创建成功: {excel_image.width} × {excel_image.height} EMU")
    
    # 清理临时文件
    processor.cleanup_temp_files()


if __name__ == "__main__":
    test_image_processor()