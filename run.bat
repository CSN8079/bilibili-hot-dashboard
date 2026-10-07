@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

rem 可选：若本机已有可用的 Python，可在运行前设置环境变量 CODEX_PYTHON 指向其完整路径（此处不写死任何本机路径）。
if defined CODEX_PYTHON if exist "%CODEX_PYTHON%" goto bundled_python

if not exist ".venv\Scripts\python.exe" goto check_system_python
".venv\Scripts\python.exe" -c "import openpyxl" >nul 2>nul
if not errorlevel 1 goto venv_python

:check_system_python
where python >nul 2>nul
if not errorlevel 1 goto system_python
where py >nul 2>nul
if not errorlevel 1 goto py_launcher

echo 未找到带 openpyxl 的 Python，请先双击 install.bat。
pause
exit /b 1

:venv_python
".venv\Scripts\python.exe" bilibili_hot_excel.py %*
set "EXIT_CODE=%ERRORLEVEL%"
goto done

:system_python
python bilibili_hot_excel.py %*
set "EXIT_CODE=%ERRORLEVEL%"
goto done

:py_launcher
py -3 bilibili_hot_excel.py %*
set "EXIT_CODE=%ERRORLEVEL%"
goto done

:bundled_python
"%CODEX_PYTHON%" bilibili_hot_excel.py %*
set "EXIT_CODE=%ERRORLEVEL%"

:done
if "%EXIT_CODE%"=="130" goto end
if not "%EXIT_CODE%"=="0" (
    echo.
    echo 程序异常退出，错误码：%EXIT_CODE%
    pause
)

:end
exit /b %EXIT_CODE%