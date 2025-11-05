#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能家居方案生成系统 - Windows部署打包配置
"""

import os
import sys
from setuptools import setup, find_packages

# 项目基本信息
PROJECT_NAME = "SmartHomeGenerator"
VERSION = "1.0.0"
DESCRIPTION = "智能家居方案生成系统 - 自动化生成智能家居方案PPT和采购清单"
AUTHOR = "智能家居方案生成团队"
AUTHOR_EMAIL = "support@smarthome-generator.com"
URL = "https://github.com/smarthome-generator/smart-home-generator"

# 读取requirements.txt
with open('requirements.txt', 'r', encoding='utf-8') as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

# 读取README.md
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

# 数据文件配置
data_files = [
    # 模板文件
    ('templates', [
        '采购清单模板.xlsx',
        '智能家居产品库示例.xlsx'
    ]),
    # 资源文件
    ('assets', [
        'assets/ppt_mold_usage_guide.md',
        'assets/模具使用文档.md',
        'assets/目录结构规划.md'
    ]),
    # 文档
    ('docs', [
        'docs/部署和配置指南.md'
    ])
]

# 打包配置
setup(
    name=PROJECT_NAME,
    version=VERSION,
    description=DESCRIPTION,
    long_description=long_description,
    long_description_content_type='text/markdown',
    author=AUTHOR,
    author_email=AUTHOR_EMAIL,
    url=URL,
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    include_package_data=True,
    install_requires=requirements,
    python_requires='>=3.8',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: End Users/Desktop',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Topic :: Office/Business',
        'Topic :: Multimedia :: Graphics :: Presentation',
    ],
    keywords='smart home, ppt generation, procurement list, automation',
    entry_points={
        'gui_scripts': [
            'smarthome-generator = smart_home_gui:main',
        ],
        'console_scripts': [
            'smarthome-cli = gui_integration:main',
        ],
    },
    data_files=data_files,
    options={
        'build_exe': {
            'includes': [
                'tkinter', 'tkinter.ttk', 'tkinter.filedialog',
                'openpyxl', 'python-pptx', 'PIL', 'requests',
                'os', 'sys', 'json', 'logging'
            ],
            'excludes': ['test', 'tests', 'unittest'],
            'optimize': 1,
        }
    }
)

if __name__ == '__main__':
    print(f"{PROJECT_NAME} v{VERSION} 打包配置已加载")