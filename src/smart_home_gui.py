#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能家居方案生成系统 - GUI桌面应用
主应用程序文件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import threading
import json
from datetime import datetime

# 添加项目路径到系统路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入集成接口
from gui_integration import GUIIntegration
from config_manager import ConfigManager

class SmartHomeGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("智能家居方案生成系统")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        # 设置应用图标
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
        
        # 创建集成接口实例
        self.integration = GUIIntegration()
        
        # 创建配置管理器实例
        self.config_manager = ConfigManager()
        
        # 当前处理状态
        self.processing = False
        
        # 创建简约风格样式
        self._create_styles()
        
        # 创建主界面
        self._create_main_interface()
        
        # 设置窗口居中
        self._center_window()
        
        # 加载历史记录和配置文件
        self._load_and_display_history()
        self._load_configuration()

    def _create_styles(self):
        """创建现代化简约风格样式"""
        # 现代化配色方案：浅色主题
        self.colors = {
            'primary': '#2563eb',      # 主色调蓝
            'secondary': '#64748b',    # 辅助色灰蓝
            'accent': '#3b82f6',       # 强调色蓝
            'success': '#10b981',      # 成功色绿
            'warning': '#f59e0b',      # 警告色橙
            'error': '#ef4444',        # 错误色红
            'light': '#f8fafc',        # 浅色背景
            'dark': '#1e293b',         # 深色文字
            'background': '#ffffff',    # 主背景白
            'card': '#f1f5f9',         # 卡片背景
            'border': '#e2e8f0'        # 边框色
        }
        
        # 配置样式
        style = ttk.Style()
        
        # 基础样式
        style.configure('TFrame', background=self.colors['background'])
        style.configure('TLabel', background=self.colors['background'], 
                       foreground=self.colors['dark'], font=('微软雅黑', 10))
        style.configure('TButton', font=('微软雅黑', 10, 'normal'), padding='10 8')
        
        # 标题样式
        style.configure('Title.TLabel', font=('微软雅黑', 16, 'bold'), 
                       foreground=self.colors['primary'])
        style.configure('Subtitle.TLabel', font=('微软雅黑', 12, 'normal'),
                       foreground=self.colors['secondary'])
        
        # 自定义按钮样式
        style.configure('Primary.TButton', 
                       background=self.colors['primary'], 
                       foreground='black',
                       borderwidth=0,
                       focuscolor='none')
        style.configure('Secondary.TButton',
                       background=self.colors['card'],
                       foreground=self.colors['dark'],
                       borderwidth=1,
                       bordercolor=self.colors['border'])
        
        # 标签框架样式
        style.configure('TLabelframe', background=self.colors['background'],
                       bordercolor=self.colors['border'])
        style.configure('TLabelframe.Label', background=self.colors['card'],
                       foreground=self.colors['dark'], font=('微软雅黑', 10, 'bold'))
        
        # 选项卡样式
        style.configure('TNotebook', background=self.colors['background'])
        style.configure('TNotebook.Tab', background=self.colors['card'],
                       foreground=self.colors['dark'], padding='10 5')
        style.map('TNotebook.Tab', background=[('selected', self.colors['primary'])],
                 foreground=[('selected', self.colors['dark'])])

    def _create_main_interface(self):
        """创建现代化主界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题区域 - 现代化设计
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 25))
        
        # 主标题
        title_label = ttk.Label(title_frame, text="智能家居方案生成系统", 
                               style='Title.TLabel')
        title_label.pack(pady=(0, 5))
        
        # 副标题
        subtitle_label = ttk.Label(title_frame, 
                                  text="现代化简约设计，高效处理智能家居方案",
                                  style='Subtitle.TLabel')
        subtitle_label.pack()
        
        # 分隔线
        separator = ttk.Separator(title_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=15)
        
        # 创建选项卡容器
        tab_container = ttk.Frame(main_frame)
        tab_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建选项卡
        self.notebook = ttk.Notebook(tab_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 模具生成选项卡
        self.mold_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.mold_frame, text="📊 模具生成")
        
        # 采购清单选项卡
        self.procurement_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.procurement_frame, text="📋 采购清单")
        
        # 创建模具生成界面
        self._create_mold_interface()
        
        # 创建采购清单界面
        self._create_procurement_interface()
        
        # 状态栏
        self._create_status_bar()

    def _create_mold_interface(self):
        """创建现代化模具生成界面"""
        # 文件选择区域 - 现代化卡片设计
        file_frame = ttk.LabelFrame(self.mold_frame, text="📁 Excel文件选择", padding="15")
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 文件信息显示 - 更清晰的视觉层次
        info_frame = ttk.Frame(file_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.mold_file_info = tk.StringVar(value="📄 未选择文件")
        file_info_label = ttk.Label(info_frame, textvariable=self.mold_file_info,
                                   font=('微软雅黑', 10), foreground=self.colors['secondary'])
        file_info_label.pack(anchor=tk.W)
        
        # 文件路径显示 - 现代化输入框
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 15))
        
        path_label = ttk.Label(path_frame, text="文件路径：", font=('微软雅黑', 9))
        path_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.mold_file_path = tk.StringVar()
        file_entry = ttk.Entry(path_frame, textvariable=self.mold_file_path, 
                              state='readonly', font=('微软雅黑', 9), width=50)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 按钮区域 - 现代化按钮布局
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(fill=tk.X)
        
        # 第一行按钮
        button_row1 = ttk.Frame(button_frame)
        button_row1.pack(fill=tk.X, pady=(0, 10))
        
        browse_btn = ttk.Button(button_row1, text="📂 选择Excel文件",
                               command=self.select_mold_file, style='Secondary.TButton')
        browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(button_row1, text="🗑️ 清除选择",
                              command=self.clear_mold_file, style='Secondary.TButton')
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 输出文件名设置区域
        filename_frame = ttk.LabelFrame(self.mold_frame, text="📝 输出文件设置", padding="15")
        filename_frame.pack(fill=tk.X, pady=(10, 20))
        
        # 文件名输入框
        name_label = ttk.Label(filename_frame, text="模具库文件名：", font=('微软雅黑', 9))
        name_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 文件名输入说明
        hint_label = ttk.Label(filename_frame, text="（仅输入文件名，后缀自动设置为.pptx）", 
                              font=('微软雅黑', 8), foreground=self.colors['secondary'])
        hint_label.pack(anchor=tk.W, pady=(0, 5))
        
        name_input_frame = ttk.Frame(filename_frame)
        name_input_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.mold_output_name = tk.StringVar(value="智能家居模具库")
        name_entry = ttk.Entry(name_input_frame, textvariable=self.mold_output_name,
                              font=('微软雅黑', 9), width=40)
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # 文件名验证
        name_entry.bind('<KeyRelease>', self._validate_filename)
        # 绑定文件名变更事件
        self.mold_output_name.trace('w', self._on_mold_filename_change)
        
        # 默认文件名按钮
        default_btn = ttk.Button(name_input_frame, text="恢复默认",
                               command=self._reset_mold_filename, style='Secondary.TButton')
        default_btn.pack(side=tk.LEFT)
        
        # 生成按钮区域 - 放在输出文件设置区域内
        generate_frame = ttk.Frame(filename_frame)
        generate_frame.pack(fill=tk.X, pady=(15, 10))
        
        generate_btn = ttk.Button(generate_frame, text="🚀 生成模具库",
                                 command=self.generate_mold_library, style='Primary.TButton')
        generate_btn.pack(side=tk.LEFT, padx=(0, 10), ipady=5)
        
        # 打开模具库文件按钮
        open_btn = ttk.Button(generate_frame, text="📂 打开模具库",
                             command=self.open_mold_library, style='Secondary.TButton')
        open_btn.pack(side=tk.LEFT, ipady=5)
        
        # 历史记录显示
        history_label = ttk.Label(filename_frame, text="历史记录：", font=('微软雅黑', 9))
        history_label.pack(anchor=tk.W, pady=(5, 0))
        
        self.mold_history_text = tk.Text(filename_frame, height=3, font=('微软雅黑', 8),
                                         bg=self.colors['light'], relief='flat',
                                         borderwidth=1, padx=5, pady=5)
        self.mold_history_text.pack(fill=tk.X, pady=(5, 0))
        self.mold_history_text.insert(tk.END, "暂无历史记录")
        self.mold_history_text.config(state='disabled')
        
        # 结果区域 - 现代化结果展示
        result_frame = ttk.LabelFrame(self.mold_frame, text="📊 生成结果", padding="15")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 结果文本框 - 现代化文本区域
        text_container = ttk.Frame(result_frame)
        text_container.pack(fill=tk.BOTH, expand=True)
        
        self.mold_result_text = tk.Text(text_container, height=12, font=('微软雅黑', 10),
                                       bg=self.colors['light'], relief='flat',
                                       borderwidth=1, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, 
                                 command=self.mold_result_text.yview)
        self.mold_result_text.configure(yscrollcommand=scrollbar.set)
        
        self.mold_result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 设置初始提示文本
        self.mold_result_text.insert(tk.END, "等待生成模具库...\n\n")
        self.mold_result_text.insert(tk.END, "请先选择Excel文件，然后点击'生成模具库'按钮。")
        self.mold_result_text.config(state='disabled')

    def _create_procurement_interface(self):
        """创建现代化采购清单界面"""
        # 文件选择区域 - 现代化卡片设计
        file_frame = ttk.LabelFrame(self.procurement_frame, text="📁 PPT文件选择", padding="15")
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 文件信息显示 - 更清晰的视觉层次
        info_frame = ttk.Frame(file_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.procurement_file_info = tk.StringVar(value="📄 未选择文件")
        file_info_label = ttk.Label(info_frame, textvariable=self.procurement_file_info,
                                   font=('微软雅黑', 10), foreground=self.colors['secondary'])
        file_info_label.pack(anchor=tk.W)
        
        # 文件路径显示 - 现代化输入框
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 15))
        
        path_label = ttk.Label(path_frame, text="文件路径：", font=('微软雅黑', 9))
        path_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.procurement_file_path = tk.StringVar()
        file_entry = ttk.Entry(path_frame, textvariable=self.procurement_file_path,
                              state='readonly', font=('微软雅黑', 9), width=50)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 按钮区域 - 现代化按钮布局
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(fill=tk.X)
        
        # 第一行按钮
        button_row1 = ttk.Frame(button_frame)
        button_row1.pack(fill=tk.X, pady=(0, 10))
        
        browse_btn = ttk.Button(button_row1, text="📂 选择PPT文件",
                               command=self.select_procurement_file, style='Secondary.TButton')
        browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(button_row1, text="🗑️ 清除选择",
                              command=self.clear_procurement_file, style='Secondary.TButton')
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 第二行按钮 - 模板和模具库选择框
        button_row2 = ttk.Frame(button_frame)
        button_row2.pack(fill=tk.X, pady=(0, 10))
        
        # 模板文件选择
        template_label = ttk.Label(button_row2, text="模板文件：", font=('微软雅黑', 9))
        template_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.template_file_path = tk.StringVar()
        template_entry = ttk.Entry(button_row2, textvariable=self.template_file_path,
                                  state='readonly', font=('微软雅黑', 9), width=30)
        template_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        template_browse_btn = ttk.Button(button_row2, text="📂 选择模板",
                                        command=self.select_template_file, style='Secondary.TButton')
        template_browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        template_clear_btn = ttk.Button(button_row2, text="🗑️ 清除",
                                       command=self.clear_template_file, style='Secondary.TButton')
        template_clear_btn.pack(side=tk.LEFT)
        
        # 第三行按钮 - 模具库文件选择
        button_row3 = ttk.Frame(button_frame)
        button_row3.pack(fill=tk.X)
        
        mold_library_label = ttk.Label(button_row3, text="模具库文件：", font=('微软雅黑', 9))
        mold_library_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.mold_library_file_path = tk.StringVar()
        mold_library_entry = ttk.Entry(button_row3, textvariable=self.mold_library_file_path,
                                       state='readonly', font=('微软雅黑', 9), width=30)
        mold_library_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        mold_library_browse_btn = ttk.Button(button_row3, text="📂 选择模具库",
                                            command=self.select_mold_library_file, style='Secondary.TButton')
        mold_library_browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        mold_library_clear_btn = ttk.Button(button_row3, text="🗑️ 清除",
                                           command=self.clear_mold_library_file, style='Secondary.TButton')
        mold_library_clear_btn.pack(side=tk.LEFT)
        
        # 结果区域 - 现代化结果展示
        result_frame = ttk.LabelFrame(self.procurement_frame, text="📋 生成采购清单", padding="15")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 文件名设置区域
        filename_frame = ttk.Frame(result_frame)
        filename_frame.pack(fill=tk.X, pady=(0, 15))
        
        filename_label = ttk.Label(filename_frame, text="采购清单文件名：", font=('微软雅黑', 9))
        filename_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.procurement_output_name = tk.StringVar(value="采购清单")
        filename_entry = ttk.Entry(filename_frame, textvariable=self.procurement_output_name,
                                  font=('微软雅黑', 9), width=30)
        filename_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # 绑定文件名验证
        filename_entry.bind('<KeyRelease>', self._validate_procurement_filename)
        
        reset_btn = ttk.Button(filename_frame, text="恢复默认",
                              command=self._reset_procurement_filename, style='Secondary.TButton')
        reset_btn.pack(side=tk.LEFT)
        
        # 按钮区域 - 生成和打开按钮
        button_frame = ttk.Frame(result_frame)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 生成采购清单按钮
        generate_btn = ttk.Button(button_frame, text="🚀 生成采购清单",
                                 command=self.generate_procurement_list, style='Primary.TButton')
        generate_btn.pack(side=tk.LEFT, padx=(0, 10), ipady=5)
        
        # 打开文件按钮
        self.open_procurement_btn = ttk.Button(button_frame, text="📂 打开文件",
                                              command=self.open_procurement_file, style='Secondary.TButton')
        self.open_procurement_btn.pack(side=tk.LEFT, ipady=5)
        
        # 初始状态下禁用打开文件按钮
        self.open_procurement_btn.config(state='disabled')
        
        # 结果文本框 - 现代化文本区域
        text_container = ttk.Frame(result_frame)
        text_container.pack(fill=tk.BOTH, expand=True)
        
        self.procurement_result_text = tk.Text(text_container, height=10, font=('微软雅黑', 10),
                                               bg=self.colors['light'], relief='flat',
                                               borderwidth=1, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, 
                                 command=self.procurement_result_text.yview)
        self.procurement_result_text.configure(yscrollcommand=scrollbar.set)
        
        self.procurement_result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 设置初始提示文本
        self.procurement_result_text.insert(tk.END, "等待生成采购清单...\n\n")
        self.procurement_result_text.insert(tk.END, "请先选择PPT文件，然后点击'生成采购清单'按钮。")
        self.procurement_result_text.config(state='disabled')

    def _create_status_bar(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_text = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_text,
                                font=('微软雅黑', 8))
        status_label.pack(side=tk.LEFT, padx=5)
        
        # 系统信息
        sys_info = tk.StringVar(value=f"系统版本: 1.0 | 运行环境: Windows")
        sys_label = ttk.Label(status_frame, textvariable=sys_info,
                             font=('微软雅黑', 8))
        sys_label.pack(side=tk.RIGHT, padx=5)

    def _center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def select_mold_file(self):
        """选择模具库文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel模具库文件",
            filetypes=[("Excel文件", "*.xlsx"), ("Excel文件", "*.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.mold_file_path.set(filename)
            self._update_file_info(filename, 'excel')
            self.update_status(f"已选择Excel文件: {os.path.basename(filename)}")
            
            # 保存配置
            self.config_manager.set_mold_generation_config(
                excel_file_path=filename,
                mold_library_filename=self.mold_output_name.get()
            )
            
    def select_procurement_file(self):
        """选择采购清单文件"""
        filename = filedialog.askopenfilename(
            title="选择PPT智能家居方案",
            filetypes=[("PowerPoint文件", "*.pptx"), ("PowerPoint文件", "*.ppt"), ("所有文件", "*.*")]
        )
        if filename:
            self.procurement_file_path.set(filename)
            self._update_file_info(filename, 'ppt')
            self.update_status(f"已选择PPT文件: {os.path.basename(filename)}")
            
            # 保存配置
            self.config_manager.set_procurement_generation_config(
                ppt_file_path=filename,
                template_file_path=self.template_file_path.get(),
                mold_library_file_path=self.mold_library_file_path.get(),
                procurement_filename=self.procurement_output_name.get()
            )
            
    def clear_mold_file(self):
        """清除模具文件选择"""
        self.mold_file_path.set("")
        self.mold_file_info.set("未选择文件")
        self.update_status("已清除Excel文件选择")
        
    def clear_procurement_file(self):
        """清除采购文件选择"""
        self.procurement_file_path.set("")
        self.procurement_file_info.set("未选择文件")
        self.update_status("已清除PPT文件选择")
        
    def select_template_file(self):
        """选择模板文件"""
        filename = filedialog.askopenfilename(
            title="选择采购清单模板文件",
            filetypes=[("Excel文件", "*.xlsx"), ("Excel文件", "*.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.template_file_path.set(filename)
            self.update_status(f"已选择模板文件: {os.path.basename(filename)}")
            
            # 保存配置
            self.config_manager.set_procurement_generation_config(
                ppt_file_path=self.procurement_file_path.get(),
                template_file_path=filename,
                mold_library_file_path=self.mold_library_file_path.get(),
                procurement_filename=self.procurement_output_name.get()
            )
            
    def select_mold_library_file(self):
        """选择模具库文件"""
        filename = filedialog.askopenfilename(
            title="选择模具库Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("Excel文件", "*.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.mold_library_file_path.set(filename)
            self.update_status(f"已选择模具库文件: {os.path.basename(filename)}")
            
            # 保存配置
            self.config_manager.set_procurement_generation_config(
                ppt_file_path=self.procurement_file_path.get(),
                template_file_path=self.template_file_path.get(),
                mold_library_file_path=filename,
                procurement_filename=self.procurement_output_name.get()
            )
            
    def clear_template_file(self):
        """清除模板文件选择"""
        self.template_file_path.set("")
        self.update_status("已清除模板文件选择")
        
    def clear_mold_library_file(self):
        """清除模具库文件选择"""
        self.mold_library_file_path.set("")
        self.update_status("已清除模具库文件选择")
        
    def _update_file_info(self, file_path: str, file_type: str):
        """更新文件信息显示"""
        try:
            file_size = os.path.getsize(file_path) / (1024 * 1024)  # MB
            file_name = os.path.basename(file_path)
            
            if file_type == 'excel':
                info_text = f"文件: {file_name} | 大小: {file_size:.1f}MB | 类型: Excel"
                if file_size > 300:
                    info_text += " ⚠️ 大文件"
                self.mold_file_info.set(info_text)
            else:  # ppt
                info_text = f"文件: {file_name} | 大小: {file_size:.1f}MB | 类型: PowerPoint"
                if file_size > 300:
                    info_text += " ⚠️ 大文件"
                self.procurement_file_info.set(info_text)
                
        except Exception as e:
            if file_type == 'excel':
                self.mold_file_info.set(f"文件: {os.path.basename(file_path)} | 无法获取文件信息")
            else:
                self.procurement_file_info.set(f"文件: {os.path.basename(file_path)} | 无法获取文件信息")

    def generate_mold_library(self):
        """生成模具库"""
        if not self.mold_file_path.get():
            messagebox.showwarning("警告", "请先选择Excel模具库文件")
            return
            
        # 验证文件
        validation = self.integration.validate_input_file(self.mold_file_path.get(), 'excel')
        if not validation.get('valid', False):
            messagebox.showwarning("文件验证失败", validation.get('message', '未知错误'))
            return
            
        if validation.get('warning'):
            if not messagebox.askyesno("文件较大", f"{validation.get('warning')}\n是否继续处理？"):
                return
            
        if self.processing:
            messagebox.showwarning("警告", "当前有任务正在处理中")
            return
            
        # 保存配置
        self.config_manager.set_mold_generation_config(
            excel_file_path=self.mold_file_path.get(),
            mold_library_filename=self.mold_output_name.get()
        )
            
        # 开始处理
        self.processing = True
        self.update_status("正在生成PPT模具库...")
        
        # 在新线程中处理
        thread = threading.Thread(target=self._generate_mold_thread)
        thread.daemon = True
        thread.start()
        
    def _generate_mold_thread(self):
        """模具生成线程"""
        try:
            # 获取用户输入的文件名
            custom_filename = self.mold_output_name.get().strip()
            if not custom_filename:
                custom_filename = "智能家居模具库"
            
            # 使用集成接口生成模具库，传递自定义文件名
            result = self.integration.generate_mold_library(
                self.mold_file_path.get(), 
                custom_filename
            )
            
            # 在主线程中更新UI
            if result.get('success', False):
                self.root.after(0, self._mold_generation_complete, result)
            else:
                self.root.after(0, self._mold_generation_error, result.get('message', '未知错误'))
            
        except Exception as e:
            self.root.after(0, self._mold_generation_error, str(e))
            
    def _mold_generation_complete(self, result):
        """模具生成完成"""
        self.processing = False
        
        # 显示结果
        self._show_mold_result(result.get('output_file'))
        self.update_status("PPT模具库生成完成")
        messagebox.showinfo("完成", f"PPT模具库已生成: {os.path.basename(result.get('output_file'))}")
        
    def _mold_generation_error(self, error_msg):
        """模具生成错误"""
        self.processing = False
        
        self.update_status("生成失败")
        messagebox.showerror("错误", f"生成PPT模具库时发生错误:\n{error_msg}")

    def generate_procurement_list(self):
        """生成采购清单"""
        if not self.procurement_file_path.get():
            messagebox.showwarning("警告", "请先选择PPT智能家居方案")
            return
            
        # 检查模板文件是否选择
        if not self.template_file_path.get():
            messagebox.showwarning("警告", "请先选择采购清单模板文件")
            return
            
        # 检查模具库文件是否选择
        if not self.mold_library_file_path.get():
            messagebox.showwarning("警告", "请先选择模具库Excel文件")
            return
            
        # 验证PPT文件
        validation = self.integration.validate_input_file(self.procurement_file_path.get(), 'ppt')
        if not validation.get('valid', False):
            messagebox.showwarning("文件验证失败", validation.get('message', '未知错误'))
            return
            
        # 验证模板文件
        template_validation = self.integration.validate_input_file(self.template_file_path.get(), 'excel')
        if not template_validation.get('valid', False):
            messagebox.showwarning("模板文件验证失败", template_validation.get('message', '未知错误'))
            return
            
        # 验证模具库文件
        mold_validation = self.integration.validate_input_file(self.mold_library_file_path.get(), 'excel')
        if not mold_validation.get('valid', False):
            messagebox.showwarning("模具库文件验证失败", mold_validation.get('message', '未知错误'))
            return
            
        if validation.get('warning'):
            if not messagebox.askyesno("文件较大", f"{validation.get('warning')}\n是否继续处理？"):
                return
            
        if self.processing:
            messagebox.showwarning("警告", "当前有任务正在处理中")
            return
            
        # 保存配置
        self.config_manager.set_procurement_generation_config(
            ppt_file_path=self.procurement_file_path.get(),
            template_file_path=self.template_file_path.get(),
            mold_library_file_path=self.mold_library_file_path.get(),
            procurement_filename=self.procurement_output_name.get()
        )
            
        # 开始处理
        self.processing = True
        self.update_status("正在生成采购清单...")
        
        # 在新线程中处理
        thread = threading.Thread(target=self._generate_procurement_thread)
        thread.daemon = True
        thread.start()
        
    def _generate_procurement_thread(self):
        """采购清单生成线程"""
        try:
            # 获取用户输入的文件名
            custom_filename = self.procurement_output_name.get().strip()
            if not custom_filename:
                custom_filename = "采购清单"
            
            # 使用集成接口生成采购清单，传递模板和模具库文件路径
            result = self.integration.generate_procurement_list(
                self.procurement_file_path.get(),
                self.template_file_path.get(),
                self.mold_library_file_path.get(),
                custom_filename
            )
            
            # 在主线程中更新UI
            if result.get('success', False):
                self.root.after(0, self._procurement_generation_complete, result)
            else:
                self.root.after(0, self._procurement_generation_error, result.get('message', '未知错误'))
            
        except Exception as e:
            self.root.after(0, self._procurement_generation_error, str(e))
            
    def _procurement_generation_complete(self, result):
        """采购清单生成完成"""
        self.processing = False
        
        # 显示结果
        self._show_procurement_result(result.get('output_file'))
        self.update_status("采购清单生成完成")
        messagebox.showinfo("完成", f"采购清单已生成: {os.path.basename(result.get('output_file'))}")
        
    def _procurement_generation_error(self, error_msg):
        """采购清单生成错误"""
        self.processing = False
        
        self.update_status("生成失败")
        messagebox.showerror("错误", f"生成采购清单时发生错误:\n{error_msg}")

    def _validate_filename(self, event):
        """验证文件名输入，确保不包含扩展名"""
        current_value = self.mold_output_name.get()
        
        # 检查是否包含扩展名（只处理常见的文件扩展名）
        common_extensions = ['.pptx', '.ppt', '.xlsx', '.xls', '.docx', '.doc', '.pdf', '.txt']
        
        for ext in common_extensions:
            if current_value.lower().endswith(ext):
                # 移除扩展名部分
                base_name = current_value[:-len(ext)]
                self.mold_output_name.set(base_name)
                
                # 显示提示信息
                self.update_status(f"文件名已自动移除扩展名{ext}，后缀固定为.pptx")
                return
        
        # 如果文件名以点结尾，可能是用户正在输入扩展名
        if current_value.endswith('.'):
            # 移除末尾的点
            base_name = current_value.rstrip('.')
            self.mold_output_name.set(base_name)
            self.update_status("文件名已自动移除末尾的点，后缀固定为.pptx")
    
    def _on_mold_filename_change(self, *args):
        """模具文件名变更事件"""
        filename = self.mold_output_name.get()
        if filename:
            self.update_status(f"模具库文件名已更新: {filename}")
            
            # 保存配置
            self.config_manager.set_mold_generation_config(
                excel_file_path=self.mold_file_path.get(),
                mold_library_filename=filename
            )
        else:
            self.update_status("模具库文件名不能为空")
    
    def _reset_mold_filename(self):
        """恢复默认文件名"""
        self.mold_output_name.set("智能家居模具库")
    
    def _validate_procurement_filename(self, event):
        """验证采购清单文件名输入，确保不包含扩展名"""
        current_value = self.procurement_output_name.get()
        
        # 检查是否包含扩展名（只处理常见的文件扩展名）
        common_extensions = ['.xlsx', '.xls', '.pptx', '.ppt', '.docx', '.doc', '.pdf', '.txt']
        
        for ext in common_extensions:
            if current_value.lower().endswith(ext):
                # 移除扩展名部分
                base_name = current_value[:-len(ext)]
                self.procurement_output_name.set(base_name)
                
                # 显示提示信息
                self.update_status(f"采购清单文件名已自动移除扩展名{ext}，后缀固定为.xlsx")
                
                # 保存配置
                self.config_manager.set_procurement_generation_config(
                    ppt_file_path=self.procurement_file_path.get(),
                    template_file_path=self.template_file_path.get(),
                    mold_library_file_path=self.mold_library_file_path.get(),
                    procurement_filename=base_name
                )
                return
        
        # 如果文件名以点结尾，可能是用户正在输入扩展名
        if current_value.endswith('.'):
            # 移除末尾的点
            base_name = current_value.rstrip('.')
            self.procurement_output_name.set(base_name)
            self.update_status("采购清单文件名已自动移除末尾的点，后缀固定为.xlsx")
            
            # 保存配置
            self.config_manager.set_procurement_generation_config(
                ppt_file_path=self.procurement_file_path.get(),
                template_file_path=self.template_file_path.get(),
                mold_library_file_path=self.mold_library_file_path.get(),
                procurement_filename=base_name
            )
        else:
            # 保存配置
            self.config_manager.set_procurement_generation_config(
                ppt_file_path=self.procurement_file_path.get(),
                template_file_path=self.template_file_path.get(),
                mold_library_file_path=self.mold_library_file_path.get(),
                procurement_filename=current_value
            )
    
    def _reset_procurement_filename(self):
        """恢复采购清单默认文件名"""
        self.procurement_output_name.set("采购清单")
    
    def _load_configuration(self):
        """加载配置文件"""
        try:
            # 加载模具生成配置
            mold_config = self.config_manager.get_mold_generation_config()
            if mold_config.get('excel_file_path') and os.path.exists(mold_config['excel_file_path']):
                self.mold_file_path.set(mold_config['excel_file_path'])
                self._update_file_info(mold_config['excel_file_path'], 'excel')
            
            if mold_config.get('mold_library_filename'):
                self.mold_output_name.set(mold_config['mold_library_filename'])
            
            # 加载采购清单生成配置
            procurement_config = self.config_manager.get_procurement_generation_config()
            if procurement_config.get('ppt_file_path') and os.path.exists(procurement_config['ppt_file_path']):
                self.procurement_file_path.set(procurement_config['ppt_file_path'])
                self._update_file_info(procurement_config['ppt_file_path'], 'ppt')
            
            if procurement_config.get('template_file_path') and os.path.exists(procurement_config['template_file_path']):
                self.template_file_path.set(procurement_config['template_file_path'])
            
            if procurement_config.get('mold_library_file_path') and os.path.exists(procurement_config['mold_library_file_path']):
                self.mold_library_file_path.set(procurement_config['mold_library_file_path'])
            
            if procurement_config.get('procurement_filename'):
                self.procurement_output_name.set(procurement_config['procurement_filename'])
                
        except Exception as e:
            print(f"加载配置文件失败: {e}")
    
    def _load_and_display_history(self):
        """加载并显示历史记录"""
        # 加载历史记录
        history_data = self._load_mold_history()
        
        # 如果有历史记录，设置默认文件名
        if history_data:
            latest_record = history_data[0]
            latest_filename = latest_record.get('filename', '智能家居模具库')
            self.mold_output_name.set(latest_filename)
            
        # 更新历史记录显示
        self._update_mold_history_display()
    
    def _save_mold_history(self, excel_file, output_file, timestamp):
        """保存模具生成历史记录"""
        history_file = os.path.join(os.path.dirname(__file__), "mold_history.json")
        history_data = {
            "excel_file": excel_file,
            "output_file": output_file,
            "timestamp": timestamp,
            "filename": self.mold_output_name.get()
        }
        
        # 读取现有历史记录
        existing_history = []
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    existing_history = json.load(f)
            except:
                existing_history = []
        
        # 添加新记录到开头，最多保留10条
        existing_history.insert(0, history_data)
        existing_history = existing_history[:10]
        
        # 保存历史记录
        try:
            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(existing_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存历史记录失败：{e}")
    
    def _load_mold_history(self):
        """加载模具生成历史记录"""
        history_file = os.path.join(os.path.dirname(__file__), "mold_history.json")
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    history_data = json.load(f)
                    return history_data
            except:
                return []
        return []
    
    def _update_mold_history_display(self):
        """更新历史记录显示"""
        history_data = self._load_mold_history()
        self.mold_history_text.config(state='normal')
        self.mold_history_text.delete(1.0, tk.END)
        
        if not history_data:
            self.mold_history_text.insert(tk.END, "暂无历史记录")
        else:
            for i, record in enumerate(history_data):
                timestamp = record.get('timestamp', '未知时间')
                filename = record.get('filename', '未知文件')
                excel_file = os.path.basename(record.get('excel_file', '未知Excel'))
                output_file = os.path.basename(record.get('output_file', '未知输出'))
                
                self.mold_history_text.insert(tk.END, f"{i+1}. {timestamp} - {filename}\\n")
                self.mold_history_text.insert(tk.END, f"   源文件：{excel_file} → 输出：{output_file}\\n")
                if i < len(history_data) - 1:
                    self.mold_history_text.insert(tk.END, "\\n")
        
        self.mold_history_text.config(state='disabled')
    
    def open_mold_library(self):
        """打开模具库文件"""
        # 首先检查当前是否选择了Excel文件
        current_excel_file = self.mold_file_path.get()
        
        if not current_excel_file:
            messagebox.showinfo("提示", "请先选择Excel模具库文件")
            return
        
        # 在Excel文件同文件夹内查找模具库文件
        excel_dir = os.path.dirname(current_excel_file)
        filename = f"{self.mold_output_name.get()}.pptx"
        mold_file_path = os.path.join(excel_dir, filename)
        
        if not os.path.exists(mold_file_path):
            # 如果找不到，尝试使用历史记录中的文件名
            history_data = self._load_mold_history()
            if history_data:
                latest_record = history_data[0]
                history_filename = latest_record.get('filename', '智能家居模具库')
                mold_file_path = os.path.join(excel_dir, f"{history_filename}.pptx")
                
                if not os.path.exists(mold_file_path):
                    # 最后尝试默认文件名
                    mold_file_path = os.path.join(excel_dir, "智能家居模具库.pptx")
        
        if not os.path.exists(mold_file_path):
            messagebox.showwarning("警告", 
                f"在Excel文件所在文件夹中找不到模具库文件：\n"
                f"文件夹：{excel_dir}\n"
                f"期望文件名：{filename}\n"
                f"请先生成模具库或检查文件是否存在")
            return
        
        try:
            # 使用系统默认程序打开文件
            os.startfile(mold_file_path)
            self.update_status(f"已打开模具库文件：{os.path.basename(mold_file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"打开文件失败：{str(e)}")
            self.update_status("打开文件失败")
    
    def open_procurement_file(self):
        """打开采购清单文件"""
        # 首先检查当前是否选择了PPT文件
        current_ppt_file = self.procurement_file_path.get()
        
        if not current_ppt_file:
            messagebox.showinfo("提示", "请先选择PPT智能家居方案文件")
            return
        
        # 在PPT文件同文件夹内查找采购清单文件
        ppt_dir = os.path.dirname(current_ppt_file)
        filename = f"{self.procurement_output_name.get()}.xlsx"
        procurement_file_path = os.path.join(ppt_dir, filename)
        
        if not os.path.exists(procurement_file_path):
            # 如果找不到，尝试默认文件名
            procurement_file_path = os.path.join(ppt_dir, "采购清单.xlsx")
        
        if not os.path.exists(procurement_file_path):
            messagebox.showwarning("警告", 
                f"在PPT文件所在文件夹中找不到采购清单文件：\n"
                f"文件夹：{ppt_dir}\n"
                f"期望文件名：{filename}\n"
                f"请先生成采购清单或检查文件是否存在")
            return
        
        try:
            # 使用系统默认程序打开文件
            os.startfile(procurement_file_path)
            self.update_status(f"已打开采购清单文件：{os.path.basename(procurement_file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"打开文件失败：{str(e)}")
            self.update_status("打开文件失败")
    
    def _show_mold_result(self, result_file):
        """显示模具生成结果"""
        # 清空结果文本框
        self.mold_result_text.delete(1.0, tk.END)
        
        # 显示结果信息
        result_info = f"生成文件: {os.path.basename(result_file)}\n"
        result_info += f"文件路径: {result_file}\n"
        result_info += f"文件大小: {os.path.getsize(result_file) / 1024:.1f} KB\n"
        result_info += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        result_info += "操作说明:\n"
        result_info += "• 点击下方按钮打开文件\n"
        result_info += "• 或手动在文件管理器中查看"
        
        self.mold_result_text.insert(1.0, result_info)
        
        # 保存历史记录
        if hasattr(self, 'current_excel_file') and self.current_excel_file:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._save_mold_history(self.current_excel_file, result_file, timestamp)
            self._update_mold_history_display()
        
        # 添加打开文件按钮
        open_btn = ttk.Button(self.mold_frame, text="打开文件",
                             command=lambda: os.startfile(result_file))
        open_btn.pack(pady=10)

    def _show_procurement_result(self, result_file):
        """显示采购清单生成结果"""
        # 清空结果文本框
        self.procurement_result_text.delete(1.0, tk.END)
        
        # 显示结果信息
        result_info = f"生成文件: {os.path.basename(result_file)}\n"
        result_info += f"文件路径: {result_file}\n"
        result_info += f"文件大小: {os.path.getsize(result_file) / 1024:.1f} KB\n"
        result_info += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        result_info += "操作说明:\n"
        result_info += "• 点击上方'打开文件'按钮打开文件\n"
        result_info += "• 或手动在文件管理器中查看"
        
        self.procurement_result_text.insert(1.0, result_info)
        
        # 启用打开文件按钮
        self.open_procurement_btn.config(state='normal')

    def update_status(self, message):
        """更新状态栏"""
        self.status_text.set(message)
        
    def run(self):
        """运行应用"""
        self.root.mainloop()


def main():
    """主函数"""
    app = SmartHomeGUI()
    app.run()


if __name__ == "__main__":
    main()