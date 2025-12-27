#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel VLOOKUP 工具
强大的Excel数据查找和匹配工具
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import time
import threading
import queue
import xlwings as xw
import sys
import subprocess
import openpyxl
from openpyxl import load_workbook
from concurrent.futures import ThreadPoolExecutor, as_completed


class VlookupTool:
    """Excel VLOOKUP工具类"""
    
    def __init__(self, parent):
        self.parent = parent
        self.window = parent
        
        # 初始化变量
        self.setup_variables()
        # 创建主界面
        self.create_main_interface()
        # 设置消息队列
        self.setup_message_queue()
    
    def setup_variables(self):
        """初始化变量"""
        self.file_a_path = tk.StringVar()
        self.file_b_path = tk.StringVar()
        self.output_file_path = None
        self.processing = False
        
    def create_main_interface(self):
        """创建主界面"""
        # 主容器
        self.main_container = tk.Frame(self.window, bg='#f5f8ff')
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        title_label = tk.Label(self.main_container, text="🔍 Excel VLOOKUP 工具", 
                              font=("微软雅黑", 18, "bold"), bg='#f5f8ff', fg='#2c7be5')
        title_label.pack(pady=(0, 20))
        
        # 文件选择区域
        self.create_file_selection()
        
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
        
        # 文件A选择
        file_a_frame = tk.Frame(file_frame, bg='#f5f8ff')
        file_a_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(file_a_frame, text="📄 源文件 (包含查找值):", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        file_a_entry_frame = tk.Frame(file_a_frame, bg='#f5f8ff')
        file_a_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        file_a_entry = tk.Entry(file_a_entry_frame, textvariable=self.file_a_path, 
                               font=("微软雅黑", 10), width=50)
        file_a_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        file_a_button = tk.Button(file_a_entry_frame, text="浏览", 
                                 command=self.select_file_a, font=("微软雅黑", 9))
        file_a_button.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 文件B选择
        file_b_frame = tk.Frame(file_frame, bg='#f5f8ff')
        file_b_frame.pack(fill=tk.X)
        
        tk.Label(file_b_frame, text="📋 查找表文件:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        file_b_entry_frame = tk.Frame(file_b_frame, bg='#f5f8ff')
        file_b_entry_frame.pack(fill=tk.X, pady=(5, 0))
        
        file_b_entry = tk.Entry(file_b_entry_frame, textvariable=self.file_b_path, 
                               font=("微软雅黑", 10), width=50)
        file_b_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        file_b_button = tk.Button(file_b_entry_frame, text="浏览", 
                                 command=self.select_file_b, font=("微软雅黑", 9))
        file_b_button.pack(side=tk.RIGHT, padx=(10, 0))
    
    def create_options_section(self):
        """创建处理选项区域"""
        options_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(options_frame, text="⚙️ 处理选项:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        # 处理模式选择
        mode_frame = tk.Frame(options_frame, bg='#f5f8ff')
        mode_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.processing_mode = tk.StringVar(value="fast")
        
        fast_radio = tk.Radiobutton(mode_frame, text="快速处理模式", 
                                   variable=self.processing_mode, value="fast",
                                   font=("微软雅黑", 10), bg='#f5f8ff')
        fast_radio.pack(side=tk.LEFT)
        
        standard_radio = tk.Radiobutton(mode_frame, text="标准处理模式", 
                                       variable=self.processing_mode, value="standard",
                                       font=("微软雅黑", 10), bg='#f5f8ff')
        standard_radio.pack(side=tk.LEFT, padx=(20, 0))
    
    def create_button_section(self):
        """创建按钮区域"""
        button_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 开始处理按钮
        self.process_button = tk.Button(button_frame, text="🚀 开始VLOOKUP处理", 
                                       command=self.start_processing, 
                                       font=("微软雅黑", 12, "bold"), 
                                       bg='#28a745', fg='white',
                                       width=20, height=2)
        self.process_button.pack(pady=10)
    
    def create_status_section(self):
        """创建状态显示区域"""
        status_frame = tk.Frame(self.main_container, bg='#f5f8ff')
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(status_frame, text="📊 处理状态:", 
                font=("微软雅黑", 11, "bold"), bg='#f5f8ff').pack(anchor=tk.W)
        
        # 状态文本框
        self.status_text = tk.Text(status_frame, height=8, font=("微软雅黑", 9),
                                  bg='#f8f9fa', fg='#495057', wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(self.status_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)
    
    def select_file_a(self):
        """选择文件A"""
        file_path = filedialog.askopenfilename(
            title="选择源文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_a_path.set(file_path)
    
    def select_file_b(self):
        """选择文件B"""
        file_path = filedialog.askopenfilename(
            title="选择查找表文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_b_path.set(file_path)
    
    def setup_message_queue(self):
        """设置消息队列"""
        self.message_queue = queue.Queue()
        self.check_queue()
    
    def check_queue(self):
        """检查消息队列"""
        try:
            while True:
                message = self.message_queue.get_nowait()
                self.update_status(message)
        except queue.Empty:
            pass
        
        # 继续检查队列
        self.window.after(100, self.check_queue)
    
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
        if not self.file_a_path.get() or not self.file_b_path.get():
            messagebox.showwarning("警告", "请先选择源文件和查找表文件！")
            return
        
        # 启动处理线程
        self.processing = True
        self.process_button.config(state=tk.DISABLED, text="处理中...")
        
        thread = threading.Thread(target=self.process_vlookup)
        thread.daemon = True
        thread.start()
    
    def process_vlookup(self):
        """执行VLOOKUP处理"""
        try:
            self.message_queue.put("🔧 开始VLOOKUP处理...")
            
            # 读取文件
            self.message_queue.put("📖 正在读取源文件...")
            df_a = pd.read_excel(self.file_a_path.get())
            
            self.message_queue.put("📖 正在读取查找表文件...")
            df_b = pd.read_excel(self.file_b_path.get())
            
            # 执行VLOOKUP逻辑
            self.message_queue.put("🔍 正在执行VLOOKUP匹配...")
            
            # 这里添加具体的VLOOKUP逻辑
            # 示例：简单的列匹配
            result_df = self.perform_vlookup(df_a, df_b)
            
            # 保存结果
            self.message_queue.put("💾 正在保存结果文件...")
            output_path = self.get_output_path()
            result_df.to_excel(output_path, index=False)
            
            self.message_queue.put(f"✅ 处理完成！结果已保存至: {output_path}")
            
        except Exception as e:
            self.message_queue.put(f"❌ 处理失败: {str(e)}")
        finally:
            self.processing = False
            self.window.after(0, self.enable_process_button)
    
    def perform_vlookup(self, df_a, df_b):
        """执行VLOOKUP操作"""
        # 这里实现具体的VLOOKUP逻辑
        # 示例：简单的合并操作
        return df_a.merge(df_b, how='left')
    
    def get_output_path(self):
        """生成输出文件路径"""
        base_name = os.path.splitext(self.file_a_path.get())[0]
        return f"{base_name}_vlookup_result.xlsx"
    
    def enable_process_button(self):
        """启用处理按钮"""
        self.process_button.config(state=tk.NORMAL, text="🚀 开始VLOOKUP处理")


if __name__ == "__main__":
    # 独立运行时的测试代码
    root = tk.Tk()
    app = VlookupTool(root)
    root.mainloop()