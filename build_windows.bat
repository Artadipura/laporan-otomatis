@echo off
echo ============================================================
echo   Build Laporan Harian Otomatis - Windows (.exe)
echo ============================================================
echo.
echo Script ini akan:
echo   1. Install semua library yang dibutuhkan
echo   2. Build aplikasi menjadi 1 file .exe
echo.
echo Proses ini hanya perlu dilakukan SEKALI. Setelah selesai,
echo file .exe di folder "dist" itu yang dikirim ke teman-teman
echo Windows (mereka tidak perlu install apapun).
echo.
pause

cd /d "%~dp0"

echo.
echo [1/2] Install dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/2] Build .exe (mohon tunggu, biasanya 2-5 menit)...
pyinstaller --onefile --windowed --name "LaporanHarianOtomatis" app.py

echo.
echo ============================================================
echo   SELESAI
echo   File .exe ada di: dist\LaporanHarianOtomatis.exe
echo   File inilah yang dikirim ke teman-teman Windows Anda.
echo ============================================================
pause
