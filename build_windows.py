#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能家居方案生成系统 - Windows打包脚本
使用PyInstaller创建独立的可执行文件
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def clean_build_dirs():
    """清理构建目录"""
    build_dirs = ['build', 'dist', 'SmartHomeGenerator.spec']
    for dir_name in build_dirs:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}")
            shutil.rmtree(dir_name) if os.path.isdir(dir_name) else os.remove(dir_name)

def create_spec_file():
    """创建PyInstaller spec文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

block_cipher = None

a = Analysis(
    ['run_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 模板文件
        ('采购清单模板.xlsx', '.'),
        ('智能家居产品库示例.xlsx', '.'),
        # 资源文件
        ('assets', 'assets'),
        ('docs', 'docs'),
        # 源代码
        ('src', 'src'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk', 
        'tkinter.filedialog',
        'openpyxl',
        'python_pptx',
        'PIL',
        'requests',
        'os',
        'sys',
        'json',
        'logging',
        'src.smart_home_gui',
        'src.gui_integration',
        'src.smart_home_mold_library',
        'src.template_based_procurement_generator',
        'src.smart_analyze_plan',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 添加图标
icon_file = None
if os.path.exists('assets/images/icon.ico'):
    icon_file = 'assets/images/icon.ico'

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SmartHomeGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为False以隐藏控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_file,
)
'''
    
    with open('SmartHomeGenerator.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    print("✓ 创建spec文件完成")

def build_executable():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # 使用PyInstaller构建
    cmd = [
        'pyinstaller',
        '--noconfirm',
        '--clean',
        '--onefile',
        '--windowed',  # 窗口模式，不显示控制台
        '--name', 'SmartHomeGenerator',
        '--add-data', '采购清单模板.xlsx;.',
        '--add-data', 'assets/智能家居产品库示例.xlsx;assets',
        '--add-data', 'src;src',
        '--hidden-import', 'tkinter',
        '--hidden-import', 'tkinter.ttk',
        '--hidden-import', 'tkinter.filedialog',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'python_pptx',
        '--hidden-import', 'PIL',
        '--hidden-import', 'requests',
        'run_gui.py'
    ]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✓ 构建成功完成")
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print("✗ 构建失败")
        print(f"错误信息: {e.stderr}")
        sys.exit(1)

def create_installer_script():
    """创建安装脚本"""
    installer_content = '''@echo off
chcp 65001 >nul

echo ========================================
echo  智能家居方案生成系统 - 安装程序
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python环境
    echo 请先安装Python 3.8或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [信息] 检测到Python环境
echo.

REM 安装依赖包
echo [信息] 正在安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 依赖包安装失败
    pause
    exit /b 1
)

echo [成功] 依赖包安装完成
echo.

REM 创建桌面快捷方式
echo [信息] 正在创建桌面快捷方式...
set "SHORTCUT_PATH=%USERPROFILE%\\Desktop\\智能家居方案生成系统.lnk"
set "TARGET_PATH=%~dp0run_gui.py"

powershell -Command "
$WshShell = New-Object -comObject WScript.Shell; 
$Shortcut = $WshShell.CreateShortcut('%SHORTCUT_PATH%'); 
$Shortcut.TargetPath = 'python'; 
$Shortcut.Arguments = '\"%TARGET_PATH%\"'; 
$Shortcut.WorkingDirectory = '%~dp0'; 
$Shortcut.Description = '智能家居方案生成系统'; 
$Shortcut.Save()"

if errorlevel 1 (
    echo [警告] 桌面快捷方式创建失败，但程序仍可运行
) else (
    echo [成功] 桌面快捷方式创建完成
)

echo.
echo ========================================
echo  安装完成！
echo ========================================
echo.
echo 使用方法：
echo 1. 双击桌面快捷方式启动程序
echo 2. 或运行: python run_gui.py
echo.
echo 按任意键退出...
pause >nul
'''
    
    with open('install.bat', 'w', encoding='utf-8') as f:
        f.write(installer_content)
    print("✓ 安装脚本创建完成")

def create_portable_package():
    """创建便携版打包脚本"""
    portable_content = '''@echo off
chcp 65001 >nul

echo ========================================
echo  智能家居方案生成系统 - 便携版
echo ========================================
echo.

echo [信息] 启动智能家居方案生成系统...
python run_gui.py

echo.
echo [信息] 程序已退出
echo 按任意键关闭窗口...
pause >nul
'''
    
    with open('启动程序.bat', 'w', encoding='utf-8') as f:
        f.write(portable_content)
    print("✓ 便携版启动脚本创建完成")

def main():
    """主函数"""
    print("=" * 60)
    print("智能家居方案生成系统 - Windows打包工具")
    print("=" * 60)
    
    # 检查PyInstaller是否安装
    try:
        import PyInstaller
        print("✓ PyInstaller已安装")
    except ImportError:
        print("✗ PyInstaller未安装，正在安装...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        print("✓ PyInstaller安装完成")
    
    # 清理构建目录
    clean_build_dirs()
    
    # 创建spec文件
    create_spec_file()
    
    # 构建可执行文件
    build_executable()
    
    # 创建安装脚本
    create_installer_script()
    
    # 创建便携版脚本
    create_portable_package()
    
    print("\n" + "=" * 60)
    print("打包完成！")
    print("生成的文件：")
    print("- dist/SmartHomeGenerator.exe (可执行文件)")
    print("- install.bat (安装脚本)")
    print("- 启动程序.bat (便携版启动脚本)")
    print("=" * 60)

if __name__ == '__main__':
    main()