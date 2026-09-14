#!/bin/bash
cd "$(dirname "$0")"

echo "============================================================"
echo "  Build Laporan Harian Otomatis - macOS (.app)"
echo "============================================================"
echo ""
echo "Script ini akan:"
echo "  1. Membuat lingkungan Python terpisah (venv)"
echo "  2. Install semua library yang dibutuhkan"
echo "  3. Build aplikasi menjadi 1 file .app"
echo ""
echo "Proses ini hanya perlu dilakukan SEKALI. Setelah selesai,"
echo "file .app di folder 'dist' itu yang dikirim ke teman-teman"
echo "pengguna Mac (mereka tidak perlu install apapun)."
echo ""
read -p "Tekan ENTER untuk mulai..."

python3 -m venv venv
source venv/bin/activate

echo ""
echo "[1/2] Install dependencies..."
pip install -r requirements.txt
pip install pyinstaller

echo ""
echo "[2/2] Build .app (mohon tunggu, biasanya 2-5 menit)..."
pyinstaller --windowed --name "LaporanHarianOtomatis" app.py

echo ""
echo "============================================================"
echo "  SELESAI"
echo "  File .app ada di: dist/LaporanHarianOtomatis.app"
echo "  File inilah yang dikirim ke teman-teman pengguna Mac Anda."
echo "============================================================"
read -p "Tekan ENTER untuk menutup..."
