"""
采购清单生成器模块 - 根据设备统计报告生成采购清单
支持所有设备类型，包括智能开关、中控屏、智能窗帘等
集成图片处理和格式美化功能
支持生成包含DISPIMG公式的Excel文件
"""

import json
import pandas as pd
from typing import Dict, List, Any
from datetime import datetime
import sys
import os
# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__)))
from image_processor import ImageProcessor
from excel_formatter import ExcelFormatter
from excel_image_replacer import ExcelImageReplacer


class ProcurementListGenerator:
    """采购清单生成器"""
    
    def __init__(self):
        # 初始化图片处理器和格式美化器
        self.image_processor = ImageProcessor()
        self.excel_formatter = ExcelFormatter()
        # 设备价格参考数据库（单位：元）
        self.device_prices = {
            # 智能开关
            "领普": {
                "二键智能开关": 89.0,
                "四键智能开关": 109.0,
                "人体存在传感器": 199.0
            },
            "易来": {
                "二键智能开关": 95.0,
                "四键智能开关": 115.0
            },
            # 中控屏
            "小米": {
                "中控屏": 1299.0
            },
            "华为": {
                "中控屏": 1599.0
            },
            # 智能窗帘
            "Aqara": {
                "智能窗帘": 899.0
            },
            # 智能灯具
            "Yeelight": {
                "智能灯具": 299.0
            },
            # 全屋WiFi
            "TP-Link": {
                "全屋WiFi": 699.0
            }
        }
        
        # 设备产品链接参考数据库
        self.device_links = {
            # 智能开关
            "领普": {
                "二键智能开关": "https://item.taobao.com/item.htm?abbucket=9&fpChannel=101&fpChannelSig=e5df04843b998062633bcc1c5e31365aa19861de&id=847484751320",
                "四键智能开关": "https://item.taobao.com/item.htm?abbucket=9&fpChannel=101&fpChannelSig=e5df04843b998062633bcc1c5e31365aa19861de&id=847484751320",
                "人体存在传感器": "https://item.taobao.com/item.htm?abbucket=9&id=673456793"
            },
            "易来": {
                "二键智能开关": "https://detail.tmall.com/item.htm?abbucket=15&id=857377043",
                "四键智能开关": "https://detail.tmall.com/item.htm?abbucket=15&id=857377043"
            },
            # 中控屏
            "小米": {
                "中控屏": "https://detail.tmall.com/item.htm?abbucket=2&id=673456789"
            },
            "华为": {
                "中控屏": "https://detail.tmall.com/item.htm?abbucket=2&id=673456790"
            },
            # 智能窗帘
            "Aqara": {
                "智能窗帘": "https://detail.tmall.com/item.htm?abbucket=15&id=673456791"
            },
            # 智能灯具
            "Yeelight": {
                "智能灯具": "https://detail.tmall.com/item.htm?abbucket=15&id=673456792"
            },
            # 全屋WiFi
            "TP-Link": {
                "全屋WiFi": "https://detail.tmall.com/item.htm?id=857377043"
            }
        }
        
        # 设备品类映射
        self.category_mapping = {
            "智能开关": "智能开关",
            "中控屏": "中控屏",
            "智能窗帘": "智能窗帘",
            "智能灯具": "智能灯具",
            "人体存在传感器": "人体存在传感器",
            "全屋WiFi": "全屋WiFi"
        }
    
    def load_statistics_data(self, statistics_report_path: str = "device_statistics_report.json") -> Dict[str, Any]:
        """
        加载设备统计报告数据
        
        Args:
            statistics_report_path: 设备统计报告文件路径
            
        Returns:
            Dict[str, Any]: 设备统计数据
        """
        try:
            with open(statistics_report_path, 'r', encoding='utf-8') as f:
                statistics_data = json.load(f)
            print(f"✅ 成功加载设备统计报告")
            return statistics_data
        except Exception as e:
            print(f"❌ 加载设备统计报告失败: {e}")
            return {}
    
    def generate_device_procurement_list(self, statistics_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        生成所有设备类型的采购清单
        
        Args:
            statistics_data: 设备统计数据
            
        Returns:
            List[Dict[str, Any]]: 采购清单数据
        """
        procurement_list = []
        
        # 获取所有设备统计数据
        category_stats = statistics_data.get('category_stats', {})
        
        if not category_stats:
            print("⚠️ 未找到设备统计数据")
            return procurement_list
        
        print(f"📋 开始生成设备采购清单，共 {len(category_stats)} 个设备类别")
        
        # 处理每个设备类别
        for category, devices in category_stats.items():
            if not devices:
                continue
                
            print(f"   📊 处理设备类别: {category}")
            
            for device in devices:
                brand = device.get('brand', '')
                device_name = device.get('device_name', '')
                specification = device.get('specification', '')
                count = device.get('count', 0)
                
                if count <= 0:
                    continue
                
                # 确定设备品类
                device_category = self.category_mapping.get(category, category)
                
                # 获取价格
                unit_price = self.device_prices.get(brand, {}).get(device_name, 0)
                if unit_price == 0:
                    # 如果找不到精确匹配，尝试通用匹配
                    for device_key in self.device_prices.get(brand, {}).keys():
                        if device_name in device_key or device_key in device_name:
                            unit_price = self.device_prices[brand][device_key]
                            break
                
                # 计算小计
                subtotal = count * unit_price
                
                # 获取产品链接
                product_link = self.device_links.get(brand, {}).get(device_name, '')
                if not product_link:
                    # 如果找不到精确匹配，尝试通用匹配
                    for device_key in self.device_links.get(brand, {}).keys():
                        if device_name in device_key or device_key in device_name:
                            product_link = self.device_links[brand][device_key]
                            break
                
                # 获取设备的PDID（需要从原始统计数据中查找）
                device_pdid = self._find_device_pdid(statistics_data, brand, device_name, specification)
                
                # 构建采购清单项
                procurement_item = {
                    '设备品类': device_category,
                    '设备': device_name,
                    '品牌': brand,
                    '型号': specification,
                    '数量': count,
                    '单位': '个',
                    '单价': unit_price,
                    '小计': subtotal,
                    '产品图片': '',  # 图片将在保存时动态添加
                    '备注': specification,
                    '产品链接': product_link,
                    'pdid': device_pdid  # 添加PDID字段
                }
                
                procurement_list.append(procurement_item)
                print(f"      ✅ 添加设备: {brand} {device_name} x {count}个 (单价: {unit_price}元)")
        
        return procurement_list
    
    def _find_device_pdid(self, statistics_data: Dict[str, Any], brand: str, device_name: str, specification: str) -> str:
        """
        根据品牌、设备名称和规格查找设备的PDID
        
        Args:
            statistics_data: 设备统计数据
            brand: 品牌
            device_name: 设备名称
            specification: 规格
            
        Returns:
            str: 设备的PDID，如果未找到返回空字符串
        """
        # 从设备统计报告中查找PDID
        device_count_data = statistics_data.get('device_count', {})
        
        # 遍历所有设备，查找匹配的设备
        for pdid, device_info in device_count_data.items():
            if (device_info.get('品牌') == brand and 
                device_info.get('设备名称') == device_name and 
                device_info.get('主规格') == specification):
                return str(pdid)
        
        # 如果没有精确匹配，尝试部分匹配
        for pdid, device_info in device_count_data.items():
            if (device_info.get('品牌') == brand and 
                device_info.get('设备名称') == device_name):
                return str(pdid)
        
        return ""
    
    def add_summary_rows(self, procurement_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        添加汇总行
        
        Args:
            procurement_list: 采购清单数据
            
        Returns:
            List[Dict[str, Any]]: 包含汇总行的采购清单
        """
        # 计算智能设备总计
        smart_device_total = sum(item['小计'] for item in procurement_list if item['设备品类'] == '智能开关')
        
        # 添加智能设备总计行
        if smart_device_total > 0:
            summary_row = {
                '设备品类': '智能设备总计',
                '设备': '',
                '品牌': '',
                '型号': '',
                '数量': '',
                '单位': '',
                '单价': '',
                '小计': smart_device_total,
                '产品图片': '',
                '备注': '',
                '产品链接': ''
            }
            procurement_list.append(summary_row)
        
        # 添加总计行
        total_row = {
            '设备品类': '总计',
            '设备': '',
            '品牌': '',
            '型号': '',
            '数量': '',
            '单位': '',
            '单价': '',
            '小计': smart_device_total,
            '产品图片': '',
            '备注': '',
            '产品链接': ''
        }
        procurement_list.append(total_row)
        
        return procurement_list
    
    def save_procurement_list(self, procurement_list: List[Dict[str, Any]], 
                            output_path: str = "智能开关采购清单.xlsx",
                            use_dispimg_formulas: bool = False) -> bool:
        """
        保存采购清单到Excel文件，集成图片插入和格式美化
        
        Args:
            procurement_list: 采购清单数据
            output_path: 输出文件路径
            use_dispimg_formulas: 是否使用DISPIMG公式而不是直接嵌入图片
            
        Returns:
            bool: 是否保存成功
        """
        try:
            # 转换为DataFrame
            df = pd.DataFrame(procurement_list)
            
            # 保存到Excel
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='采购清单', index=False)
                
                # 获取工作表
                worksheet = writer.sheets['采购清单']
                
                if use_dispimg_formulas:
                    # 使用DISPIMG公式
                    self.insert_dispimg_formulas(worksheet, procurement_list)
                else:
                    # 直接插入设备图片
                    self.insert_device_images(worksheet, procurement_list)
                
                # 应用格式美化
                self.excel_formatter.format_worksheet(worksheet)
                
                # 设置超链接格式
                self.excel_formatter.format_hyperlink_cells(worksheet)
            
            print(f"💾 采购清单已保存至: {output_path}")
            
            # 如果使用了DISPIMG公式，使用图片替换器进行替换
            if use_dispimg_formulas:
                print("🔄 开始替换DISPIMG公式为嵌入图片...")
                replacer = ExcelImageReplacer()
                output_with_images = output_path.replace('.xlsx', '_with_images.xlsx')
                
                # --- [!! 修正 Bug 2 !!] ---
                # 'pdid' 列在 'L' 列 (第12列), 不是 'A' 列
                # '产品图片' 列在 'I' 列 (第9列)
                # -------------------------
                success = replacer.replace_dispimg_formulas(
                    excel_path=output_path,
                    output_path=output_with_images,
                    pdid_column="L",  # <-- 修正于此 (A -> L)
                    image_column="I",  # 图片在I列
                    start_row=2       # 从第2行开始（第1行是标题）
                )
                
                if success:
                    print(f"✅ 图片替换完成，最终文件: {output_with_images}")
                else:
                    print("⚠️ 图片替换失败，保留原始DISPIMG公式文件")
            
            return True
            
        except Exception as e:
            print(f"❌ 保存采购清单失败: {e}")
            return False
    def insert_device_images(self, worksheet, procurement_list):
        """
        插入设备图片到Excel工作表 (直接嵌入模式)
        
        Args:
            worksheet: Excel工作表对象
            procurement_list: 采购清单数据
        """
        print("🖼️  开始插入设备图片...")
        
        # 图片插入起始位置（第2行开始，I列 - 产品图片列）
        image_row = 2
        image_col_letter = "I" # I列
        
        # --- [!! 修正 Bug 3 (Part 1) !!] ---
        # 1. 必须设置列宽以容纳图片
        # (您的 excel_image_replacer.py 中有此设置, 但直接嵌入模式没有)
        # -----------------------------------
        worksheet.column_dimensions[image_col_letter].width = 25  # 约180像素

        for i, item in enumerate(procurement_list):
            # 跳过汇总行
            if item['设备品类'] in ['智能设备总计', '总计']:
                image_row += 1
                continue
            
            device_name = item['设备']
            brand = item['品牌']
            pdid = item.get('pdid', '')  # 获取PDID (小写, 正确)
            
            # 创建Excel图片对象，传递PDID参数
            excel_image = self.image_processor.create_excel_image(device_name, pdid)
            
            if excel_image:
                # 设置图片位置
                cell_ref = f"{image_col_letter}{image_row}"  # I2, I3, etc.
                excel_image.anchor = cell_ref
                
                # --- [!! 修正 Bug 3 (Part 2) !!] ---
                # 2. 必须设置行高以容纳图片
                # -----------------------------------
                worksheet.row_dimensions[image_row].height = 80  # 约106像素
                
                # (可选) 调整图片大小以适应单元格
                # image_processor 似乎已经处理了尺寸, 但我们以防万一
                try:
                    target_height_px = 80 * (96/72) # 转换为像素
                    scale = target_height_px / excel_image.height 
                    excel_image.height = target_height_px
                    excel_image.width = excel_image.width * scale
                except Exception:
                    # 如果 image_processor 返回的不是 openpyxl Image 对象，
                    # 而是 PIL Image，这里的逻辑会失败，但 image_processor 内部似乎已经处理了
                    pass

                # 添加到工作表
                worksheet.add_image(excel_image)
                print(f"   ✅ 插入图片: {brand} {device_name} (PDID: {pdid}) 到 {cell_ref}")
            else:
                print(f"   ⚠️  未找到图片: {brand} {device_name} (PDID: {pdid})")
            
            image_row += 1
        
        # 清理临时文件
        self.image_processor.cleanup_temp_files()
    def insert_dispimg_formulas(self, worksheet, procurement_list):
        """
        插入DISPIMG公式到Excel工作表
        
        Args:
            worksheet: Excel工作表对象
            procurement_list: 采购清单数据
        """
        print("📝 开始插入DISPIMG公式...")
        
        # 从第2行开始（第1行是标题）
        for row_idx, device_data in enumerate(procurement_list, start=2):
            # 跳过汇总行 (修正逻辑以匹配您的汇总行)
            if device_data.get("设备品类", "") in ['智能设备总计', '总计']:
                continue
                
            # 获取PDID
            # --- [!! 修正 Bug 1 !!] ---
            # 键名是 'pdid' (小写), 不是 'PDID' (大写)
            # -------------------------
            pdid = device_data.get("pdid", "") # <-- 修正于此 (PDID -> pdid)
            if not pdid:
                print(f"   ⚠️  第 {row_idx} 行: PDID为空，跳过DISPIMG公式插入")
                continue
                
            # 在I列插入DISPIMG公式
            try:
                cell_ref = f"I{row_idx}"
                # 创建WPS DISPIMG公式
                # 注意: 这里插入的pdid (如 '13') 将在 Bug 2 修复后被替换器 (L列) 正确找到
                dispimg_formula = f'=DISPIMG("{pdid}", 1)' 
                worksheet[cell_ref] = dispimg_formula
                print(f"   ✅ 已插入DISPIMG公式到 {cell_ref}: {dispimg_formula}")
                    
            except Exception as e:
                print(f"   ❌ 插入DISPIMG公式失败 (PDID: {pdid}): {e}")
        
        print("✅ DISPIMG公式插入完成")
    def generate_procurement_report(self, statistics_report_path: str = "device_statistics_report.json",
                                 output_path: str = "智能设备采购清单.xlsx") -> bool:
        """
        生成完整的采购清单报告
        
        Args:
            statistics_report_path: 设备统计报告文件路径
            output_path: 输出文件路径
            
        Returns:
            bool: 是否生成成功
        """
        print("[START] 开始生成智能设备采购清单...")
        
        # 1. 加载设备统计数据
        statistics_data = self.load_statistics_data(statistics_report_path)
        if not statistics_data:
            print("❌ 无法加载设备统计数据，采购清单生成终止")
            return False
        
        # 2. 生成所有设备采购清单
        procurement_list = self.generate_device_procurement_list(statistics_data)
        if not procurement_list:
            print("[WARN] 未生成任何采购清单项")
            return False
        
        # 3. 添加汇总行
        procurement_list = self.add_summary_rows(procurement_list)
        
        # 4. 保存采购清单
        success = self.save_procurement_list(procurement_list, output_path)
        
        if success:
            print("[SUCCESS] 智能设备采购清单生成完成！")
            print(f"[INFO] 生成采购清单项: {len(procurement_list) - 2} 个设备")
            total_amount = procurement_list[-1]['小计'] if procurement_list else 0
            print(f"[INFO] 采购总金额: {total_amount:.2f} 元")
        
        return success


def test_device_procurement():
    """测试所有设备类型采购清单生成功能"""
    print("🧪 测试智能设备采购清单生成...")
    
    generator = ProcurementListGenerator()
    
    # 测试生成所有设备采购清单
    success = generator.generate_procurement_report(
        statistics_report_path="device_statistics_report.json",
        output_path="test_device_procurement_list.xlsx"
    )
    
    if success:
        print("✅ 智能设备采购清单测试成功")
        
        # 读取并显示生成的采购清单内容
        try:
            df = pd.read_excel("test_device_procurement_list.xlsx")
            print("\n📋 生成的采购清单内容:")
            print(df.to_string(index=False))
        except Exception as e:
            print(f"❌ 读取采购清单失败: {e}")
    else:
        print("❌ 智能设备采购清单测试失败")
    
    return success


if __name__ == "__main__":
    test_device_procurement()