@echo off
chcp 65001 >nul
title 批量复制并重命名工具

:: 获取脚本所在目录
set "scriptDir=%~dp0"

:: 调用 Python 脚本，传递所有拖拽的参数
python "%scriptDir%batch_copy_rename.py" %*

:: 等待用户按键（便于查看结果）
pause