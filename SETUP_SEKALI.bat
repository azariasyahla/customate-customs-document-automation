@echo off
cd /d "%~dp0"
echo Membuat lingkungan aplikasi...
python -m venv .venv
if errorlevel 1 goto gagal
echo Memasang kebutuhan aplikasi...
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 goto gagal
echo Setup selesai. Klik JALANKAN.bat.
pause
exit /b
:gagal
echo Setup gagal. Pastikan Python terpasang dan internet tersedia.
pause
