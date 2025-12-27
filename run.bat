@echo off
chcp 65001 >nul
echo 🔧 Excel工具集启动程序
echo ========================================

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到Python，请先安装Python
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ 检测到Python

REM 检查依赖库
echo 正在检查依赖库...
python -c "import pandas, xlwings, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ⚠️ 缺少必要的依赖库，正在安装...
    pip install pandas xlwings openpyxl --user
    if errorlevel 1 (
        echo ❌ 依赖安装失败，请检查网络连接
        pause
        exit /b 1
    )
    echo ✅ 依赖安装完成
)

echo ✅ 所有依赖库已就绪

REM 启动主程序
echo 正在启动Excel工具集...
python main.py

if errorlevel 1 (
    echo ❌ 程序启动失败
    pause
    exit /b 1
)

echo ✅ 程序已退出