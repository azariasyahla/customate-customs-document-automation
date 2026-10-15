@echo off
cd /d "%~dp0"
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onedir --name CUSTOMATE app.py
if errorlevel 1 goto failed
echo Share the entire dist\CUSTOMATE folder, not just the EXE.
pause
exit /b
:failed
echo Build failed. Check the error above.
pause
