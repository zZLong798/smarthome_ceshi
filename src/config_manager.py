#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件管理器
用于存储和加载用户选择的文件路径配置
"""

import os
import json
from datetime import datetime


class ConfigManager:
    """配置文件管理器类"""
    
    def __init__(self, config_file=None):
        """
        初始化配置管理器
        
        Args:
            config_file: 配置文件路径，默认为程序目录下的 config.json
        """
        if config_file is None:
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        
        self.config_file = config_file
        self._load_config()
    
    def _load_config(self):
        """加载配置文件"""
        # 先获取默认配置
        default_config = self._get_default_config()
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    
                    # 合并配置，保留新字段
                    self.config = self._merge_config(default_config, loaded_config)
                    
                print(f"配置文件加载成功: {self.config_file}")
            else:
                print("配置文件不存在，使用默认配置")
                self.config = default_config
                self.save_config()  # 创建默认配置文件
                
        except Exception as e:
            print(f"加载配置文件时出错: {e}")
            # 出错时使用默认配置
            self.config = default_config
            self.save_config()
    
    def _get_default_config(self):
        """获取默认配置"""
        return {
            "version": "1.0",
            "last_updated": datetime.now().isoformat(),
            "mold_generation": {
                "excel_file_path": "",
                "mold_library_filename": "智能家居模具库"
            },
            "procurement_generation": {
                "ppt_file_path": "",
                "template_file_path": "",
                "mold_library_file_path": "",
                "procurement_filename": "采购清单"
            },
            "recent_files": {
                "excel_files": [],
                "ppt_files": [],
                "template_files": []
            }
        }
    
    def _merge_config(self, default_config, current_config):
        """合并配置，确保新字段存在"""
        for key, value in default_config.items():
            if key not in current_config:
                current_config[key] = value
            elif isinstance(value, dict) and isinstance(current_config[key], dict):
                self._merge_config(value, current_config[key])
        return current_config
    
    def save_config(self):
        """保存配置文件"""
        try:
            self.config["last_updated"] = datetime.now().isoformat()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def set_mold_generation_config(self, excel_file_path="", mold_library_filename=""):
        """设置模具生成配置"""
        if excel_file_path:
            self.config["mold_generation"]["excel_file_path"] = excel_file_path
            self._add_recent_file("excel_files", excel_file_path)
        
        if mold_library_filename:
            self.config["mold_generation"]["mold_library_filename"] = mold_library_filename
        
        return self.save_config()
    
    def set_procurement_generation_config(self, ppt_file_path="", template_file_path="", 
                                        mold_library_file_path="", procurement_filename=""):
        """设置采购清单生成配置"""
        if ppt_file_path:
            self.config["procurement_generation"]["ppt_file_path"] = ppt_file_path
            self._add_recent_file("ppt_files", ppt_file_path)
        
        if template_file_path:
            self.config["procurement_generation"]["template_file_path"] = template_file_path
            self._add_recent_file("template_files", template_file_path)
        
        if mold_library_file_path:
            self.config["procurement_generation"]["mold_library_file_path"] = mold_library_file_path
        
        if procurement_filename:
            self.config["procurement_generation"]["procurement_filename"] = procurement_filename
        
        return self.save_config()
    
    def _add_recent_file(self, file_type, file_path):
        """添加最近使用的文件"""
        if file_path not in self.config["recent_files"][file_type]:
            self.config["recent_files"][file_type].insert(0, file_path)
            # 最多保留10个最近文件
            self.config["recent_files"][file_type] = self.config["recent_files"][file_type][:10]
    
    def get_mold_generation_config(self):
        """获取模具生成配置"""
        return self.config["mold_generation"]
    
    def get_procurement_generation_config(self):
        """获取采购清单生成配置"""
        return self.config["procurement_generation"]
    
    def get_recent_files(self, file_type):
        """获取最近使用的文件列表"""
        return self.config["recent_files"].get(file_type, [])
    
    def clear_config(self):
        """清除所有配置"""
        self.config = {
            "version": "1.0",
            "last_updated": datetime.now().isoformat(),
            "mold_generation": {
                "excel_file_path": "",
                "mold_library_filename": "智能家居模具库"
            },
            "procurement_generation": {
                "ppt_file_path": "",
                "template_file_path": "",
                "mold_library_file_path": "",
                "procurement_filename": "采购清单"
            },
            "recent_files": {
                "excel_files": [],
                "ppt_files": [],
                "template_files": []
            }
        }
        return self.save_config()


def test_config_manager():
    """测试配置管理器"""
    config_manager = ConfigManager()
    
    # 测试设置配置
    config_manager.set_mold_generation_config(
        excel_file_path="C:/test/excel.xlsx",
        mold_library_filename="测试模具库"
    )
    
    config_manager.set_procurement_generation_config(
        ppt_file_path="C:/test/ppt.pptx",
        template_file_path="C:/test/template.xlsx",
        mold_library_file_path="C:/test/mold_library.xlsx",
        procurement_filename="测试采购清单"
    )
    
    # 测试获取配置
    mold_config = config_manager.get_mold_generation_config()
    procurement_config = config_manager.get_procurement_generation_config()
    
    print("模具生成配置:", mold_config)
    print("采购清单生成配置:", procurement_config)
    
    # 测试清除配置
    # config_manager.clear_config()


if __name__ == "__main__":
    test_config_manager()