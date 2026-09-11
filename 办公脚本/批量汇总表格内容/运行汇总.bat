@echo off
chcp 65001 > nul
setlocal
:: 获取当前批处理脚本所在目录
set SCRIPT_DIR=%~dp0
echo 正在启动表格汇总工具...

:: 如果拖拽了文件夹，则传递给 Python 脚本
if "%~1"=="" (
    echo 未提供文件夹路径，将手动输入...
    python "%SCRIPT_DIR%batch_summary.py"
) else (
    echo 接收到拖拽文件夹: %~1
    python "%SCRIPT_DIR%batch_summary.py" "%~1"
)
echo.
echo 程序执行完毕，按任意键退出...
pause > nul