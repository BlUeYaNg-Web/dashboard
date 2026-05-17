@echo off
chcp 65001 >nul
cd /d "%~dp0.."
echo [%date% %time%] 同步最新腳本...
git pull origin main
echo [%date% %time%] 開始執行更新...
python scripts\update_realprice.py
pause
