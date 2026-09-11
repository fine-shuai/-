@echo off
chcp 65001 >nul
title 生成 XK060 文件清单

:: 检查是否拖拽了文件夹
if "%~1"=="" (
    echo 请将文件夹拖拽到此批处理文件上。
    pause
    exit /b
)

set target=%~1
:: 去掉可能的尾部反斜杠
if "%target:~-1%"=="\" set target=%target:~0,-1%

:: 检查文件夹是否存在
if not exist "%target%" (
    echo 文件夹不存在: %target%
    pause
    exit /b
)

:: 检查 Python 是否可用
python --version >nul 2>&1
if errorlevel 1 (
    echo 未找到 Python，请先安装 Python 并添加到 PATH。
    pause
    exit /b
)

:: 检查 openpyxl 库是否已安装
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 缺少 openpyxl 库，正在尝试安装...
    pip install openpyxl
    if errorlevel 1 (
        echo 安装失败，请手动执行：pip install openpyxl
        pause
        exit /b
    )
    echo 安装成功。
)

:: 获取批处理所在目录，并执行 Python 脚本
set script_dir=%~dp0
set python_script=%script_dir%generate_xlsx.py

if not exist "%python_script%" (
    echo 未找到 Python 脚本: %python_script%
    pause
    exit /b
)

echo 正在处理文件夹: %target%
echo 请稍候...

python "%python_script%" "%target%"

echo.
echo 处理完成。按任意键退出...
pause >nul