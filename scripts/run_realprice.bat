@echo off
chcp 65001 >nul
echo [%date% %time%] 啟動實價登錄自動更新...

:: 切換到 repo 目錄（請修改為你的實際路徑）
cd /d "%~dp0.."

:: 執行更新腳本
python scripts\update_realprice.py

if %errorlevel% neq 0 (
    echo 更新失敗，錯誤碼：%errorlevel%
    pause
) else (
    echo 更新完成！
)
