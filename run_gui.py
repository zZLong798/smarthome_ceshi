#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能家居方案生成系统 - GUI启动脚本
"""

import os
import sys
import subprocess

def check_dependencies():
    """检查依赖包"""
    required_packages = [
        'tkinter',  # 通常Python自带
        'openpyxl',
        'python-pptx',
        'Pillow'
    ]
    
    missing_packages = []
    for package in required_packages:
        try:
            if package == 'tkinter':
                import tkinter
            elif package == 'openpyxl':
                import openpyxl
            elif package == 'python-pptx':
                from pptx import Presentation
            elif package == 'Pillow':
                from PIL import Image
        except ImportError:
            missing_packages.append(package)
    
    return missing_packages

def install_dependencies(missing_packages):
    """安装缺失的依赖包"""
    if not missing_packages:
        return True
        
    print("正在安装缺失的依赖包...")
    
    package_mapping = {
        'tkinter': None,  # 系统自带，无需安装
        'openpyxl': 'openpyxl',
        'python-pptx': 'python-pptx',
        'Pillow': 'Pillow'
    }
    
    for package in missing_packages:
        if package == 'tkinter':
            print("错误: tkinter 是Python标准库，但当前环境可能不支持GUI")
            return False
        
        pip_package = package_mapping.get(package)
        if pip_package:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_package])
                print(f"✓ 已安装 {package}")
            except subprocess.CalledProcessError:
                print(f"✗ 安装 {package} 失败")
                return False
    
    return True

def main():
    """主函数"""
    print("=" * 50)
    print("智能家居方案生成系统 - GUI启动器")
    print("=" * 50)
    
    # 检查依赖
    print("检查依赖包...")
    missing_packages = check_dependencies()
    
    if missing_packages:
        print(f"发现缺失的包: {', '.join(missing_packages)}")
        
        # 询问是否安装
        response = input("是否自动安装缺失的依赖包? (y/n): ").lower().strip()
        if response in ['y', 'yes', '是']:
            if not install_dependencies(missing_packages):
                print("依赖包安装失败，请手动安装后重试")
                input("按回车键退出...")
                return
        else:
            print("请手动安装缺失的依赖包后重试")
            input("按回车键退出...")
            return
    else:
        print("✓ 所有依赖包已就绪")
    
    # 启动GUI应用
    print("启动GUI应用...")
    try:
        sys.path.insert(0, 'src')
        from smart_home_gui import main as gui_main
        gui_main()
    except Exception as e:
        print(f"启动GUI应用时发生错误: {e}")
        print("请确保所有依赖包已正确安装")
        input("按回车键退出...")


if __name__ == "__main__":
    main()