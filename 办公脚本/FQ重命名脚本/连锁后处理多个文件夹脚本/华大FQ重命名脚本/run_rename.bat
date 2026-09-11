@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ========================================
echo   华大MGI T7测序下机数据文件批量重命名工具
echo          （批量处理版）
echo ========================================

REM 获取批处理文件所在目录
set "SCRIPT_DIR=%~dp0"
set "MAPPING_FILE=%SCRIPT_DIR%mapping.xlsx"

REM 检查映射文件是否存在
if not exist "%MAPPING_FILE%" (
    echo 错误: 未找到映射文件 mapping.xlsx
    echo 请确保 mapping.xlsx 文件与本批处理文件在同一目录下
    pause
    exit /b 1
)

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python
    echo 请先安装Python并确保已添加到系统PATH环境变量中
    pause
    exit /b 1
)

REM 检查Python脚本是否存在
set "PYTHON_SCRIPT=%SCRIPT_DIR%rename_mgi_fq_files_batch.py"
if not exist "%PYTHON_SCRIPT%" (
    echo 错误: 未找到Python脚本 rename_mgi_fq_files_batch.py
    echo 请确保 rename_mgi_fq_files_batch.py 文件与本批处理文件在同一目录下
    pause
    exit /b 1
)

REM 初始化参数
set "DATA_DIR="
set "DRY_RUN="
set "BATCH_MODE="

REM 解析命令行参数
:parse_args
if "%~1"=="" goto end_parse

if /i "%~1"=="--dry-run" (
    set "DRY_RUN=--dry-run"
    shift
    goto parse_args
)

if /i "%~1"=="--batch" (
    set "BATCH_MODE=--batch"
    shift
    goto parse_args
)

REM 如果不是参数，则认为是数据目录
if not defined DATA_DIR (
    set "DATA_DIR=%~1"
) else (
    echo 警告: 忽略额外的参数: %~1
)
shift
goto parse_args

:end_parse

REM 如果没有指定数据目录，提示用户输入
if not defined DATA_DIR (
    echo 使用方法:
    echo   1. 将包含子文件夹的根文件夹拖放到此批处理文件上
    echo   2. 或双击此批处理文件，然后手动输入根文件夹路径
    echo   3. 或使用命令行: run_rename.bat [根目录] [选项]
    echo.
    echo 可用选项:
    echo   --dry-run     试运行，不实际修改文件
    echo   --batch       强制批量模式（默认自动识别，但可显式指定）
    echo.
    set /p "DATA_DIR=请输入根目录路径: "
)

REM 移除路径两端的引号（如果有）
set "DATA_DIR=%DATA_DIR:"=%"

REM 检查数据目录是否存在
if not exist "%DATA_DIR%" (
    echo 错误: 数据目录不存在 - %DATA_DIR%
    pause
    exit /b 1
)

echo.
echo 正在处理根目录: %DATA_DIR%
echo 使用映射文件: %MAPPING_FILE%
if defined DRY_RUN echo 模式: 试运行 (不实际修改文件)
if defined BATCH_MODE echo 模式: 批量处理子文件夹
echo.

REM 执行Python脚本
python "%PYTHON_SCRIPT%" -m "%MAPPING_FILE%" -d "%DATA_DIR%" %DRY_RUN% %BATCH_MODE%

REM 检查执行结果
if errorlevel 1 (
    echo.
    echo 脚本执行失败，请检查错误信息
) else (
    echo.
    echo 脚本执行完成
)

echo.
pause