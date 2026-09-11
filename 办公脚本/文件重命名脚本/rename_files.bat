@echo off
chcp 65001 >nul
python "%~dp0rename_files.py" %*
pause