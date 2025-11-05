# 智能家居方案生成系统

## 项目概述
专业的智能家居设计工具，提供标准化的智能家居产品模具库，支持自动识别和报价功能。

**新增功能：PDID设备识别分析系统**
- 从PPT文件中自动提取PDID标签
- 智能设备信息查询和统计分析
- 生成详细的设备清单报告
- 完整的端到端测试框架

## 目录结构

```
smarthome/
├── src/                    # 核心源代码
│   ├── smart_home_mold_library.py     # 模具库核心类
│   ├── smart_home_product_library.py  # 产品库
│   ├── ppt_to_excel_bridge.py         # PPT到Excel桥梁
│   ├── ppt_template_generator.py      # 模板生成器
│   ├── pdid_extractor.py              # PDID标签提取模块
│   ├── device_info_query.py           # 设备信息查询模块
│   ├── device_statistics.py           # 设备统计模块
│   ├── brief_report_generator.py      # 简要报告生成模块
│   ├── device_inventory_report.py     # 设备清单报告模块
│   ├── pdid_analysis_main.py          # 主分析程序
│   ├── pdid_integration_test.py       # PDID集成测试
│   └── integration_test.py            # 系统集成测试
├── tools/                  # 工具脚本
│   ├── check_mold_gallery.py          # 模具检查
│   ├── check_template.py              # 模板检查
│   ├── debug_mold_names.py           # 调试工具
│   ├── read_ppt.py                    # PPT读取
│   ├── read_excel.py                  # Excel读取
│   ├── demo_smart_switch_object.py    # 演示脚本
│   ├── smart_switch_marker.py         # 开关标记
│   └── ppt_shape_gallery.py           # 形状库
├── output/                 # 生成的文件
│   ├── smart_home_mold_gallery.pptx   # 模具库PPT
│   ├── smart_home_shape_gallery.pptx  # 形状库PPT
│   ├── smart_home_template.pptx       # 模板文件
│   ├── smart_home_template.potx       # 模板文件
│   ├── smart_home_shape_gallery_采购清单.xlsx  # 采购清单
│   ├── 采购清单.xlsx                  # 采购清单
│   ├── 全屋智能方案.pptx              # 方案文件
│   ├── device_statistics_report.json  # 设备统计报告
│   ├── brief_device_report.json       # 简要设备报告
│   ├── device_inventory_report.json   # 设备清单报告
│   └── pdid_analysis_summary.json     # PDID分析总结
├── assets/                # 资源文件
│   ├── ppt_mold_usage_guide.md        # 使用指南
│   ├── 模具使用文档.md                 # 模具文档
│   ├── 目录结构规划.md                # 目录规划
│   └── device_database.json           # 设备数据库
├── docs/                  # 文档
│   └── 智能家居方案生成/   # 项目文档
└── README.md              # 项目说明
```

## 快速开始

### PDID设备识别分析系统

#### 1. 运行设备识别分析
```bash
cd src
python pdid_analysis_main.py
```

#### 2. 生成设备清单报告
```bash
python device_inventory_report.py
```

#### 3. 运行集成测试
```bash
python pdid_integration_test.py
```

### 原有功能

#### 1. Excel到PPT模具库转换
```bash
cd src
python excel_to_ppt_converter.py
```

#### 2. 生成模具库
```bash
cd src
python smart_home_mold_library.py
```

#### 3. 查看模具库
```bash
cd tools
python check_mold_gallery.py
```

#### 4. 使用自动识别
```bash
cd src
python smart_switch_marker.py```-c "from ppt_to_excel_bridge import PPTtoExcelBridge; bridge = PPTtoExcelBridge(); result = bridge.scan_ppt_file('../output/smart_home_mold_gallery.pptx'); print(f'识别到 {len(result)} 个产品')"
```

## 功能特性

### PDID设备识别分析系统
- **PDID标签提取**: 从PPT文件中自动提取PDID标签信息
- **设备信息查询**: 基于PDID查询详细的设备信息
- **设备统计分析**: 生成设备统计报告和分布分析
- **设备清单生成**: 创建详细的设备清单报告
- **集成测试**: 完整的端到端测试框架

### 原有功能
- **智能模具库**: 9种标准智能家居产品模具
- **自动识别**: 基于智能标记的产品识别
- **统一设计**: 标准化颜色、尺寸和布局
- **报价生成**: 自动生成采购清单

## 模具内容

- **智能开关**: 1-4键开关
- **传感器**: 人体感应、门窗感应、温湿度
- **控制器**: 智能网关、场景控制器

## 依赖要求

- Python 3.8+
- python-pptx
- openpyxl

## 使用说明

详细使用说明请参考 `assets/模具使用文档.md`

## 更新日志

- 2024年: 项目重构，目录结构优化
- 初始版本: 智能家居模具库开发完成
