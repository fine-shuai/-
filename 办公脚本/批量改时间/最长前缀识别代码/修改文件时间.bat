@echo off
title 文件时间批量修改工具
echo 正在启动Python脚本...
set script_dir=%~dp0
python "%script_dir%file_time_modifier.py" %*
pause