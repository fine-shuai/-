@echo off
chcp 936 >nul
setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "PY_SCRIPT=%SCRIPT_DIR%replace_excel_data.py"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Please install Python first.
    pause
    exit /b 1
)

if "%~1"=="" (
    echo Drag a file or folder onto this script.
    pause
    exit /b
)

echo.
echo Select mode:
echo   [1] Replace directly (modify files)
echo   [2] Dry run (only count, no changes)
set /p MODE="Enter 1 or 2: "

if "%MODE%"=="1" (
    set "DRY_RUN="
) else if "%MODE%"=="2" (
    set "DRY_RUN=--dry-run"
) else (
    echo Invalid input.
    pause
    exit /b
)

echo.
echo Running...
python "%PY_SCRIPT%" "%~1" %DRY_RUN%

echo.
pause