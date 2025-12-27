#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 日期分类工具
按日期自动分类Excel数据，支持多种格式
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import time
import threading
import xlwings as xw
import openpyxl
from openpyxl import load_workbook
from datetime import datetime


class DateFilterTool:
    """Excel 日期分类工具类"""
    
    def __init__(self, parent):
        self.parent = parent
        self.window = parent
        
        # 初始化变量
        self.file_path = tk.StringVar()
        self.date_column = tk.StringVar()
        self.processing = False
        self.output_file_path = None
        self.xl_app = None
        
        # 创建主界面
        self.create_main_interface()
    
    def create_main_interface(self):
        """创建主界面"""
        # 主容器
        self.main_container = tk.Frame(self.window, bg='#f5f8ff')
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        title_label = tk.Label(self.main_container, text="📊 Excel 日期分类工具", 
                              font=("微软雅黑", 18, "bold"), bg='#f5f8ff', fg='#2c7be5')
        title_label.pack(pady=(0, 20))
        
        # 文件选择区域
        self.create_file_selection()
        
        # 日期列选择区域
        self.create_column_selection()
        
        # 处理选项区域
        self.create_options_section()
        
        # 按钮区域
        self.create_button_section()
        
        # 状态显示区域
        self.create_status_section()
    
    def create_file_selection(self):
        """创建文件选择区域"""
        file_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(file_frame, text="📄 选择Excel文件:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        file_entry_frame = tk.Frame(file_frame, bg='#f5f8ff')
        file_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        file_entry = tk.Entry(file_entry_frame, textvariable=self.file_path, 
                             font=("微软雅黑", 10), width=50)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        file_button = tk.Button(file_entry_frame, text="浏览", 
                               command=self.select_file, font=("微软雅黑", 9))
        file_button.pack(side=tk.RIGHT, padx=(10, 0))
    
    def create_column_selection(self):
        """创建日期列选择区域"""
        column_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        column_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(column_frame, text="📅 选择日期列:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        column_entry_frame = tk.Frame(column_frame, bg='#f5f8ff')
        column_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        column_entry = tk.Entry(column_entry_frame, textvariable=self.date_column, 
                               font=("微软雅黑", 10), width=20)
        column_entry.pack(side=tk.LEFT)
        
        # 自动检测按钮
        detect_button = tk.Button(column_entry_frame, text="自动检测", 
                                 command=self.auto_detect_columns, 
                                 font=("微软雅黑", 9))
        detect_button.pack(side=tk.LEFT, padx=(10, 0))
    
    def create_options_section(self):
        """创建处理选项区域"""
        options_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(options_frame, text="⚙️ 处理选项:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        # 保留原数据选项
        self.keep_original = tk.BooleanVar(value=True)
        keep_check = tk.Checkbutton(options_frame, text="保留原数据", 
                                   variable=self.keep_original,
                                   font=("微软雅黑", 10), bg='#f5f8ff')
        keep_check.pack(anchor=tk.W, pady=(5, 0))
        
        # 日期格式选项
        format_frame = tk.Frame(options_frame, bg='#f5f8ff')
        format_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Label(format_frame, text="日期格式:", 
                font=("微软雅黑", 10), bg='#f5f8ff').pack(side=tk.LEFT)
        
        self.date_format = tk.StringVar(value="YYYY-MM-DD")
        format_combo = ttk.Combobox(format_frame, textvariable=self.date_format,
                                   values=["YYYY-MM-DD", "YYYY/MM/DD", "MM-DD-YYYY", "DD/MM/YYYY"],
                                   state="readonly", width=15)
        format_combo.pack(side=tk.LEFT, padx=(10, 0))
    
    def create_button_section(self):
        """创建按钮区域"""
        button_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 开始处理按钮
        self.process_button = tk.Button(button_frame, text="🚀 开始日期分类", 
                                       command=self.start_processing, 
                                       font=("微软雅黑", 12, "bold"), 
                                       bg='#007bff', fg='white',
                                       width=20, height=2)
        self.process_button.pack(pady=10)
    
    def create_status_section(self):
        """创建状态显示区域"""
        status_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(status_frame, text="📊 处理状态:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        # 状态文本框
        self.status_text = tk.Text(status_frame, height=6, font=("微软雅黑", 9),
                                  bg='#f8f9fa', fg='#495057', wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(self.status_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)
    
    def select_file(self):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            # 自动检测列
            self.auto_detect_columns()
    
    def auto_detect_columns(self):
        """自动检测日期列"""
        if not self.file_path.get():
            return
        
        try:
            # 读取Excel文件
            df = pd.read_excel(self.file_path.get())
            
            # 查找包含日期的列
            date_columns = []
            for col in df.columns:
                # 检查列名是否包含日期相关关键词
                if any(keyword in str(col).lower() for keyword in ['date', '时间', '日期']):
                    date_columns.append(col)
            
            if date_columns:
                self.date_column.set(date_columns[0])
                self.update_status(f"✅ 自动检测到日期列: {date_columns[0]}")
            else:
                self.update_status("⚠️ 未检测到明显的日期列，请手动指定")
                
        except Exception as e:
            self.update_status(f"❌ 自动检测失败: {str(e)}")
    
    def update_status(self, message):
        """更新状态显示"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.update()
    
    def start_processing(self):
        """开始处理"""
        if self.processing:
            return
        
        # 检查文件是否选择
        if not self.file_path.get():
            messagebox.showwarning("警告", "请先选择Excel文件！")
            return
        
        # 检查日期列是否指定
        if not self.date_column.get():
            messagebox.showwarning("警告", "请指定日期列！")
            return
        
        # 启动处理线程
        self.processing = True
        self.process_button.config(state=tk.DISABLED, text="处理中...")
        
        thread = threading.Thread(target=self.process_date_filter)
        thread.daemon = True
        thread.start()
    
    def process_date_filter(self):
        """执行日期分类处理"""
        try:
            self.update_status("🔧 开始日期分类处理...")
            
            # 读取文件
            self.update_status("📖 正在读取Excel文件...")
            df = pd.read_excel(self.file_path.get())
            
            # 检查日期列是否存在
            if self.date_column.get() not in df.columns:
                raise ValueError(f"日期列 '{self.date_column.get()}' 不存在于文件中")
            
            # 日期分类处理
            self.update_status("📅 正在按日期分类数据...")
            result_df = self.classify_by_date(df)
            
            # 保存结果
            self.update_status("💾 正在保存结果文件...")
            output_path = self.get_output_path()
            result_df.to_excel(output_path, index=False)
            
            self.update_status(f"✅ 处理完成！结果已保存至: {output_path}")
            
        except Exception as e:
            self.update_status(f"❌ 处理失败: {str(e)}")
        finally:
            self.processing = False
            self.window.after(0, self.enable_process_button)
    
    def classify_by_date(self, df):
        """按日期分类数据"""
        # 这里实现具体的日期分类逻辑
        # 示例：按年-月分组
        df['日期'] = pd.to_datetime(df[self.date_column.get()])
        df['年份'] = df['日期'].dt.year
        df['月份'] = df['日期'].dt.month
        
        return df
    
    def get_output_path(self):
        """生成输出文件路径"""
        base_name = os.path.splitext(self.file_path.get())[0]
        return f"{base_name}_date_classified.xlsx"
    
    def enable_process_button(self):
        """启用处理按钮"""
        self.process_button.config(state=tk.NORMAL, text="🚀 开始日期分类")


if __name__ == "__main__":
    # 独立运行时的测试代码
    root = tk.Tk()
    app = DateFilterTool(root)
    root.mainloop()