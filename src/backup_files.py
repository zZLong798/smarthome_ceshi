#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件备份模块 - 任务0：环境准备和备份
为产品ID标准化和PPT模具库改进项目创建文件备份
"""

import os
import shutil
from datetime import datetime

def backup_file(file_path, backup_dir="backup"):
    """备份单个文件"""
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在: {file_path}")
        return False
    
    # 创建备份目录
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    
    # 生成备份文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = os.path.basename(file_path)
    name_part, ext_part = os.path.splitext(file_name)
    backup_name = f"{name_part}_备份_{timestamp}{ext_part}"
    backup_path = os.path.join(backup_dir, backup_name)
    
    # 执行备份
    try:
        shutil.copy2(file_path, backup_path)
        print(f"✅ 备份成功: {file_path} -> {backup_path}")
        return True
    except Exception as e:
        print(f"❌ 备份失败: {file_path} - {e}")
        return False

def main():
    """主函数 - 备份所有相关文件"""
    print("=" * 60)
    print("📁 开始文件备份 - 任务0：环境准备和备份")
    print("=" * 60)
    
    # 需要备份的文件列表
    files_to_backup = [
        "E:\\Programs\\smarthome\\智能家居模具库.xlsx",
        "E:\\Programs\\smarthome\\智能家居模具库.pptx"
    ]
    
    backup_dir = "E:\\Programs\\smarthome\\backup"
    
    # 创建备份目录
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    
    success_count = 0
    total_count = len(files_to_backup)
    
    for file_path in files_to_backup:
        if backup_file(file_path, backup_dir):
            success_count += 1
    
    print("-" * 60)
    print(f"📊 备份完成: {success_count}/{total_count} 个文件备份成功")
    
    # 检查备份结果
    if success_count == total_count:
        print("✅ 所有文件备份成功，环境准备就绪")
        return True
    else:
        print("⚠️  部分文件备份失败，请检查文件路径")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)