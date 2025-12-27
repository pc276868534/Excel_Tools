# Excel工具集

一个强大的Excel数据处理工具集，包含VLOOKUP工具和日期分类工具。

## 功能特点

### 🔍 VLOOKUP工具
- 支持多值查找（换行符分隔）
- 快速处理和标准处理模式
- 完美保留原文件格式

### 📊 日期分类工具
- 支持多种日期格式
- 可选择保留原数据
- 统一设置行高和格式

## 安装使用

### 方法1：直接运行Python脚本
```bash
pip install pandas xlwings openpyxl
python main.py
```

### 方法2：打包成可执行文件
```bash
# 使用PyInstaller打包
pip install pyinstaller
pyinstaller --onefile --windowed --name ExcelTools main.py
```

## 项目结构

```
Excel_Tool/
├── main.py              # 主程序入口
├── vlookup.py           # VLOOKUP工具模块
├── datefilter.py        # 日期分类工具模块
├── excel_utils.py       # Excel工具函数库
├── run.bat              # Windows运行脚本
├── build.bat            # 打包脚本
├── setup.py             # 打包配置
└── README.md            # 项目说明
```

## 技术支持


- 版本：2.0
- 开发时间：2025年

## 许可证

© 2025 Jason. All rights reserved.
