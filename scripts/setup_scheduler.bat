@echo off
chcp 65001 >nul
echo 正在設定每週日 13:00 自動更新實價登錄...

:: 取得目前 bat 檔的所在目錄（即 scripts 資料夾）
set SCRIPT_DIR=%~dp0
set BAT_PATH=%SCRIPT_DIR%run_realprice.bat

:: 建立工作排程（每週日 13:00，登入後即可執行）
schtasks /create ^
  /tn "實價登錄週報自動更新" ^
  /tr "%BAT_PATH%" ^
  /sc weekly ^
  /d SUN ^
  /st 13:00 ^
  /f ^
  /rl highest

if %errorlevel% equ 0 (
    echo.
    echo ✅ 設定完成！
    echo    每週日 13:00 會自動執行更新。
    echo    工作名稱：實價登錄週報自動更新
    echo.
    echo 若要手動觸發，請執行：
    echo    schtasks /run /tn "實價登錄週報自動更新"
) else (
    echo.
    echo ❌ 設定失敗，請用系統管理員身份執行此檔案。
)
pause
