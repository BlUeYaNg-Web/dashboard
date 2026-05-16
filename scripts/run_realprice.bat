@echo off
cd /d "%~dp0.."
python scripts\update_realprice.py
pause
