@echo off
chcp 65001 >nul
title 表格数据汇总工具
echo 正在启动数据汇总脚本...

:: 获取拖拽的文件或文件夹路径
set "target=%~1"
if "%target%"=="" (
    echo 未检测到文件或文件夹，请将 Excel 文件或文件夹拖拽到此批处理文件上。
    pause
    exit /b
)

:: 切换到脚本所在目录
cd /d "%~dp0"

:: 执行Python脚本
python summary_script.py "%target%"

:: 如果Python未安装或路径错误，给出提示
if errorlevel 1 (
    echo 执行失败，请确保已安装 Python 并配置了环境变量。
    pause
)