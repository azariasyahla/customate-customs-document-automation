@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
echo Jalankan SETUP_SEKALI.bat terlebih dahulu.
pause
exit /b
)
.venv\Scripts\python app.py
pause
