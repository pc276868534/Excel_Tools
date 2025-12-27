#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 工具集 - 主程序
提供VLOOKUP工具和日期分类工具的统一入口
"""
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys


class ExcelToolsMain:
    """Excel工具集主窗口"""
    
    def __init__(self, root):
        self.root = root
        
        # 工具管理变量 - 必须在create_widgets之前初始化
        self.current_tool = None
        self.vlookup_tool = None
        self.datefilter_tool = None
        self.current_tool_frame = None
        
        self.setup_window()
        self.create_widgets()
        
        # 确保窗口关闭时释放资源
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def setup_window(self):
        """设置窗口基本属性"""
        self.root.title("Excel 工具集 - 主页")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        self.root.minsize(800, 600)
    
    def create_widgets(self):
        """创建界面组件"""
        # 主容器
        main_container = tk.Frame(self.root, bg='#f5f8ff')
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        
        # 创建工具显示区域
        self.tool_display_frame = tk.Frame(main_container, bg='#f5f8ff')
        self.tool_display_frame.pack(expand=True, fill=tk.BOTH)
        

        
        # 显示主页内容（默认）
        self.show_home_page()
        
        # 创建底部信息区域
        bottom_frame = tk.Frame(main_container, bg='#f5f8ff')
        bottom_frame.pack(fill=tk.X, pady=(30, 0))
        
        # 版本信息
        version_label = tk.Label(bottom_frame, text="版本: 2.0 © 2025", 
                               font=("微软雅黑", 10), bg='#f5f8ff', fg='#6c757d')
        version_label.pack(side=tk.LEFT)
        
        # 创建菜单栏
        self.create_menu_bar()
    

    
    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 工具菜单
        tool_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="🔧 工具", menu=tool_menu)
        tool_menu.add_command(label="🏠 主页", command=self.show_home_page)
        tool_menu.add_separator()
        tool_menu.add_command(label="🔍 VLOOKUP工具", command=self.show_vlookup_tool)
        tool_menu.add_command(label="📊 日期分类工具", command=self.show_datefilter_tool)
        tool_menu.add_separator()
        tool_menu.add_command(label="❌ 退出", command=self.on_close)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="❓ 帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
    
    def show_home_page(self):
        """显示主页"""
        # 检查是否有工具正在处理
        if self.check_processing_state():
            return
            
        self.hide_current_tool()
        
        home_frame = tk.Frame(self.tool_display_frame, bg='#f5f8ff')
        home_frame.pack(fill=tk.BOTH, expand=True)
        
        # 欢迎信息
        welcome_frame = tk.Frame(home_frame, bg='#f5f8ff')
        welcome_frame.pack(expand=True)
        
        welcome_label = tk.Label(welcome_frame, text="🏠 Excel工具集 ", 
                               font=("微软雅黑", 20, "bold"), bg='#f5f8ff', fg='#2c7be5')
        welcome_label.pack(pady=(0, 20))
        
        desc_label = tk.Label(welcome_frame, text="请点击下方工具卡片选择要使用的工具", 
                            font=("微软雅黑", 14), bg='#f5f8ff', fg='#6c757d')
        desc_label.pack()
        
        # 工具介绍
        tools_info = tk.Frame(welcome_frame, bg='#f5f8ff')
        tools_info.pack(pady=(40, 0))
        
        # VLOOKUP工具卡片（可点击）
        vlookup_info = tk.Frame(tools_info, bg='#e3f2fd', relief=tk.RAISED, bd=2, cursor='hand2')
        vlookup_info.pack(fill=tk.X, pady=(0, 20))
        vlookup_info.bind("<Button-1>", lambda e: self.show_vlookup_tool())
        vlookup_info.bind("<Enter>", lambda e: vlookup_info.config(relief=tk.SOLID, bd=3))
        vlookup_info.bind("<Leave>", lambda e: vlookup_info.config(relief=tk.RAISED, bd=2))
        
        # 让所有子组件也响应点击事件
        def bind_click_to_children(widget):
            widget.bind("<Button-1>", lambda e: self.show_vlookup_tool())
            widget.bind("<Enter>", lambda e: vlookup_info.config(relief=tk.SOLID, bd=3))
            widget.bind("<Leave>", lambda e: vlookup_info.config(relief=tk.RAISED, bd=2))
            for child in widget.winfo_children():
                bind_click_to_children(child)
        
        tk.Label(vlookup_info, text="🔍 VLOOKUP 工具", 
                font=("微软雅黑", 14, "bold"), bg='#e3f2fd').pack(pady=(10, 5))
        tk.Label(vlookup_info, text="强大的Excel数据查找和匹配工具", 
                font=("微软雅黑", 11), bg='#e3f2fd').pack()
        tk.Label(vlookup_info, text="• 支持多值查找（换行符分隔）", 
                font=("微软雅黑", 10), bg='#e3f2fd').pack()
        tk.Label(vlookup_info, text="• 快速处理和标准处理模式", 
                font=("微软雅黑", 10), bg='#e3f2fd').pack()
        tk.Label(vlookup_info, text="• 完美保留原文件格式", 
                font=("微软雅黑", 10), bg='#e3f2fd').pack(pady=(0, 10))
        
        bind_click_to_children(vlookup_info)
        
        # 日期分类工具卡片（可点击）
        datefilter_info = tk.Frame(tools_info, bg='#f3e5f5', relief=tk.RAISED, bd=2, cursor='hand2')
        datefilter_info.pack(fill=tk.X)
        datefilter_info.bind("<Button-1>", lambda e: self.show_datefilter_tool())
        datefilter_info.bind("<Enter>", lambda e: datefilter_info.config(relief=tk.SOLID, bd=3))
        datefilter_info.bind("<Leave>", lambda e: datefilter_info.config(relief=tk.RAISED, bd=2))
        
        def bind_click_to_datefilter(widget):
            widget.bind("<Button-1>", lambda e: self.show_datefilter_tool())
            widget.bind("<Enter>", lambda e: datefilter_info.config(relief=tk.SOLID, bd=3))
            widget.bind("<Leave>", lambda e: datefilter_info.config(relief=tk.RAISED, bd=2))
            for child in widget.winfo_children():
                bind_click_to_datefilter(child)
        
        tk.Label(datefilter_info, text="📊 日期分类工具", 
                font=("微软雅黑", 14, "bold"), bg='#f3e5f5').pack(pady=(10, 5))
        tk.Label(datefilter_info, text="按日期自动分类Excel数据", 
                font=("微软雅黑", 11), bg='#f3e5f5').pack()
        tk.Label(datefilter_info, text="• 支持多种日期格式", 
                font=("微软雅黑", 10), bg='#f3e5f5').pack()
        tk.Label(datefilter_info, text="• 可选择保留原数据", 
                font=("微软雅黑", 10), bg='#f3e5f5').pack()
        tk.Label(datefilter_info, text="• 统一设置行高和格式", 
                font=("微软雅黑", 10), bg='#f3e5f5').pack(pady=(0, 10))
        
        bind_click_to_datefilter(datefilter_info)
        
        self.current_tool_frame = home_frame
        self.current_tool = "home"
        
        # 调整窗口大小以适应主页
        self.root.update_idletasks()  # 强制更新界面
        self.root.geometry("900x750")  # 恢复主页窗口大小
    
    def show_vlookup_tool(self):
        """显示VLOOKUP工具"""
        # 检查当前是否有工具正在处理
        if self.check_processing_state():
            return
            
        self.hide_current_tool()
        
        try:
            # 导入VLOOKUP工具模块
            from vlookup import VlookupTool
            
            # 创建内嵌的VLOOKUP工具
            vlookup_frame = tk.Frame(self.tool_display_frame, bg='#f5f8ff')
            vlookup_frame.pack(fill=tk.BOTH, expand=True)
            
            self.vlookup_tool = VlookupTool(vlookup_frame)
            self.vlookup_tool.window = vlookup_frame  # 更新窗口引用
            
            self.current_tool_frame = vlookup_frame
            self.current_tool = "vlookup"
            
            # 调整窗口大小以适应VLOOKUP工具
            self.root.update_idletasks()  # 强制更新界面
            self.root.geometry("1100x850")  # 设置适合VLOOKUP工具的窗口大小
            
        except ImportError as e:
            messagebox.showerror("错误", f"无法导入VLOOKUP工具模块: {str(e)}\n请确保vlookup.py文件存在")
            self.show_home_page()
        except Exception as e:
            messagebox.showerror("错误", f"启动VLOOKUP工具失败: {str(e)}")
            self.show_home_page()
    
    def show_datefilter_tool(self):
        """显示日期分类工具"""
        # 检查当前是否有工具正在处理
        if self.check_processing_state():
            return
            
        self.hide_current_tool()
        
        try:
            # 导入日期分类工具模块
            from datefilter import DateFilterTool
            
            # 创建内嵌的日期分类工具
            datefilter_frame = tk.Frame(self.tool_display_frame, bg='#f5f8ff')
            datefilter_frame.pack(fill=tk.BOTH, expand=True)
            
            self.datefilter_tool = DateFilterTool(datefilter_frame)
            self.datefilter_tool.window = datefilter_frame  # 更新窗口引用
            
            self.current_tool_frame = datefilter_frame
            self.current_tool = "datefilter"
            
            # 调整窗口大小以适应日期分类工具
            self.root.update_idletasks()  # 强制更新界面
            self.root.geometry("1000x800")  # 设置适合日期分类工具的窗口大小
            
        except ImportError as e:
            messagebox.showerror("错误", f"无法导入日期分类工具模块: {str(e)}\n请确保datefilter.py文件存在")
            self.show_home_page()
        except Exception as e:
            messagebox.showerror("错误", f"启动日期分类工具失败: {str(e)}")
            self.show_home_page()
    
    def check_processing_state(self):
        """检查处理状态，如果有工具正在处理则弹出确认对话框"""
        # 检查VLOOKUP工具是否正在处理
        if self.vlookup_tool and hasattr(self.vlookup_tool, 'processing') and self.vlookup_tool.processing:
            if messagebox.askokcancel("停止处理", "VLOOKUP工具正在处理中，确定要停止并切换吗？"):
                # 停止VLOOKUP工具的处理
                if hasattr(self.vlookup_tool, 'xl_app') and self.vlookup_tool.xl_app:
                    try:
                        self.vlookup_tool.xl_app.quit()
                    except:
                        pass
                self.vlookup_tool.processing = False
                return False  # 允许切换
            else:
                return True  # 阻止切换
        
        # 检查日期分类工具是否正在处理
        if self.datefilter_tool and hasattr(self.datefilter_tool, 'processing') and self.datefilter_tool.processing:
            if messagebox.askokcancel("停止处理", "日期分类工具正在处理中，确定要停止并切换吗？"):
                # 停止日期分类工具的处理
                if hasattr(self.datefilter_tool, 'xl_app') and self.datefilter_tool.xl_app:
                    try:
                        self.datefilter_tool.xl_app.quit()
                    except:
                        pass
                self.datefilter_tool.processing = False
                return False  # 允许切换
            else:
                return True  # 阻止切换
        
        return False  # 没有处理中的工具，允许切换
    
    def hide_current_tool(self):
        """隐藏当前工具"""
        if self.current_tool_frame:
            self.current_tool_frame.destroy()
            self.current_tool_frame = None
        
        # 清理工具实例
        if self.vlookup_tool:
            self.vlookup_tool = None
        if self.datefilter_tool:
            self.datefilter_tool = None
    
    def show_help(self):
        """显示帮助信息"""
        help_text = """
🔧 Excel 工具集使用说明

📋 功能介绍：
1. VLOOKUP工具 - 用于Excel数据的快速查找和匹配
   • 支持多值查找（换行符分隔）
   • 快速处理和标准处理模式
   • 完美保留原文件格式

2. 日期分类工具 - 按日期自动分类Excel数据
   • 支持多种日期格式
   • 可选择保留原数据
   • 统一设置行高和格式

💡 使用提示：
• 点击工具卡片或菜单项启动对应工具
• 每个工具都有独立的操作界面
• 处理完成后会自动保存为新文件

📞 技术支持：
• 开发人员：Jason
• 联系电话：18816703105
        """
        messagebox.showinfo("使用说明", help_text.strip())
    
    def show_about(self):
        """显示关于信息"""
        about_text = """
🔧 Excel 工具集 - 美少女专用版

版本: 2.0
开发时间: 2025年


🌟 功能特点：
• 专业的Excel数据处理
• 友好的用户界面
• 高效的批量处理
• 完整的格式保留

© 2025 Jason. All rights reserved.
        """
        messagebox.showinfo("关于", about_text.strip())
    
    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def on_close(self):
        """窗口关闭事件处理"""
        # 检查是否有工具正在处理
        if self.check_processing_state():
            return
        
        if messagebox.askokcancel("退出", "确定要退出Excel工具集吗？"):
            # 清理所有可能的Excel应用实例
            try:
                if self.vlookup_tool and hasattr(self.vlookup_tool, 'xl_app') and self.vlookup_tool.xl_app:
                    self.vlookup_tool.xl_app.quit()
                if self.datefilter_tool and hasattr(self.datefilter_tool, 'xl_app') and self.datefilter_tool.xl_app:
                    self.datefilter_tool.xl_app.quit()
            except:
                pass
            self.root.destroy()


def check_dependencies():
    """检查必要的依赖库"""
    required_libraries = ['tkinter', 'pandas', 'xlwings', 'openpyxl']
    missing_libraries = []
    
    for lib in required_libraries:
        try:
            if lib == 'tkinter':
                import tkinter
            else:
                __import__(lib)
        except ImportError:
            missing_libraries.append(lib)
    
    if missing_libraries:
        print("❌ 缺少必要的库:")
        for lib in missing_libraries:
            print(f"  • {lib}")
        print("\n请运行以下命令安装:")
        print(f"pip install {' '.join(missing_libraries)}")
        return False
    
    print("✅ 所有必要库已安装")
    return True


def main():
    """主程序入口"""
    # 检查依赖库
    if not check_dependencies():
        input("按Enter键退出...")
        return
    
    # 创建主窗口
    root = tk.Tk()
    
    # 设置窗口图标（如果存在）
    try:
        if os.path.exists("icon.ico"):
            root.iconbitmap("icon.ico")
    except:
        pass
    
    # 创建应用实例
    app = ExcelToolsMain(root)
    
    # 启动主循环
    root.mainloop()


if __name__ == "__main__":
    main()

