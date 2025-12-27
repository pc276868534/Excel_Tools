"""
Excel工具公共模块
包含重复使用的工具函数和类
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import time
from datetime import datetime


class ExcelUtils:
    """Excel工具公共类"""
    
    @staticmethod
    def format_sheet_name(date, date_format):
        """格式化工作表名称 - 通用版本"""
        format_map = {
            "YYYY-MM-DD": date.strftime("%Y-%m-%d"),
            "YYYY/MM/DD": date.strftime("%Y/%m/%d"),
            "YYYY年MM月DD日": date.strftime("%Y年%m月%d日"),
            "MM-DD-YYYY": date.strftime("%m-%d-%Y"),
            "DD/MM/YYYY": date.strftime("%d/%m/%Y")
        }
        return format_map.get(date_format, date.strftime("%Y-%m-%d"))
    
    @staticmethod
    def validate_excel_file(file_path):
        """验证Excel文件是否存在且有效"""
        if not file_path:
            return False, "请选择Excel文件"
        
        if not os.path.exists(file_path):
            return False, "选择的文件不存在"
        
        if not file_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
            return False, "请选择有效的Excel文件"
        
        return True, "文件验证通过"
    
    @staticmethod
    def get_excel_columns(file_path):
        """获取Excel文件的列名"""
        try:
            df = pd.read_excel(file_path, nrows=0)
            return list(df.columns)
        except Exception as e:
            raise ValueError(f"读取Excel文件列名失败: {str(e)}")
    
    @staticmethod
    def get_save_location(default_name, title="保存文件"):
        """获取保存位置"""
        output_path = filedialog.asksaveasfilename(
            title=title,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=default_name
        )
        return output_path if output_path else None
    
    @staticmethod
    def parse_date_value(date_value):
        """解析日期值，支持多种格式"""
        if not date_value:
            return None
            
        try:
            # 尝试pandas的日期解析
            return pd.to_datetime(date_value).date()
        except:
            # 如果pandas解析失败，尝试手动解析
            if isinstance(date_value, str):
                for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日', '%m-%d-%Y', '%d/%m/%Y']:
                    try:
                        return datetime.strptime(str(date_value).strip(), fmt).date()
                    except ValueError:
                        continue
        return None
    
    @staticmethod
    def create_ui_frame(parent, title, subtitle):
        """创建统一的UI标题框架"""
        title_frame = tk.Frame(parent, bg='#f5f8ff')
        title_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = tk.Label(title_frame, text=title, 
                             font=("微软雅黑", 18, "bold"), bg='#f5f8ff', fg='#2c7be5')
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame, text=subtitle, 
                                font=("微软雅黑", 12), bg='#f5f8ff', fg='#6c757d')
        subtitle_label.pack()
        
        return title_frame
    
    @staticmethod
    def create_file_selection_frame(parent, label_text="Excel文件:", var=None):
        """创建文件选择框架"""
        file_frame = ttk.LabelFrame(parent, text="📁 选择Excel文件", padding=15)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(file_frame, text=label_text, font=("微软雅黑", 10)).pack(side=tk.LEFT)
        
        if var is None:
            var = tk.StringVar()
        
        entry_file = ttk.Entry(file_frame, textvariable=var, width=50)
        entry_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 10))
        
        return file_frame, var
    
    @staticmethod
    def add_status_message(status_text, msg, is_error=False):
        """添加状态消息到文本框"""
        status_text.insert(tk.END, f"{msg}\n")
        if is_error:
            status_text.tag_add("error", "end-2l", "end-1l")
            status_text.tag_config("error", foreground="red")
        status_text.see(tk.END)


# 日期格式选项常量
DATE_FORMATS = ["YYYY-MM-DD", "YYYY/MM/DD", "YYYY年MM月DD日", "MM-DD-YYYY", "DD/MM/YYYY"]
