@echo off
setlocal
cd /d "%~dp0"
if exist "bilibili_hot_dashboard.url" (
    start "" "bilibili_hot_dashboard.url"
) else (
    echo 看板地址文件尚未生成，请先运行 run.bat。
    pause
)