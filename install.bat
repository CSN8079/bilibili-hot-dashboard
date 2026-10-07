@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

rem 可选：若本机已有可用的 Python，可在运行前设置环境变量 CODEX_PYTHON 指向其完整路径（此处不写死任何本机路径）。
if defined CODEX_PYTHON if exist "%CODEX_PYTHON%" goto already_ready

where python >nul 2>nul
if not errorlevel 1 goto use_python
where py >nul 2>nul
if not errorlevel 1 goto use_py

echo 未找到 Python。请先安装 Python 3.10 或更高版本，并勾选 Add Python to PATH。
pause
exit /b 1

:use_python
set "PYTHON=python"
goto create_venv

:use_py
set "PYTHON=py -3"

:create_venv
%PYTHON% -m venv .venv
if errorlevel 1 goto failed

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto failed

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed

echo.
echo 安装完成。请双击 run.bat 启动。
pause
exit /b 0

:already_ready
echo 已检测到可用的 Python 和 openpyxl，无需安装。
echo 请直接双击 run.bat 启动程序。
pause
exit /b 0

:failed
echo.
echo 安装失败，请检查网络连接和上方错误信息。
pause
exit /b 1