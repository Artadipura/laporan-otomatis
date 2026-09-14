# Laporan Harian Otomatis - Desktop App

## Struktur

```
laporan_otomatis_app/
├── app.py                      ← jalankan file ini untuk membuka aplikasi
├── requirements.txt
├── backend/
│   └── report_processor.py     ← logika inti (mapping, dedup, dst) - SAMA PERSIS dgn main.py asli
└── gui/
    ├── main_window.py          ← tampilan window utama
    └── worker.py                ← menjalankan proses di background thread
```

## Packaging jadi .exe / .app (supaya teman-teman tinggal pakai)

Ini cukup dilakukan **SEKALI per platform**, bukan oleh setiap orang yang mau pakai.

**Windows** (di komputer Anda sendiri): double-click `build_windows.bat`, tunggu sampai selesai. Hasilnya di `dist\LaporanHarianOtomatis.exe` — file inilah yang dikirim ke teman-teman Windows.

**macOS** (harus dijalankan di komputer Mac asli, titip ke satu teman yang punya Mac): double-click `build_macos.command`. Kalau macOS menolak jalankan karena "unidentified developer", klik kanan file itu → **Open** → konfirmasi Open sekali. Hasilnya di `dist/LaporanHarianOtomatis.app` — ini yang disebar ke teman-teman pengguna Mac.

Catatan: karena aplikasinya belum "ditandatangani" (code signing, butuh akun developer berbayar), saat pertama dibuka:
- **Windows**: mungkin muncul "Windows protected your PC" → klik "More info" → "Run anyway".
- **Mac**: klik kanan file `.app` → Open → konfirmasi sekali saja.
Setelah itu bisa dibuka normal dengan double-click seperti aplikasi biasa.

## Cara Menjalankan (mode developer, tanpa di-package)

1. Install Python 3.10+ (kalau belum ada).
2. Buka terminal / Command Prompt di folder ini, lalu jalankan:

   ```
   pip install -r requirements.txt
   ```

3. Jalankan aplikasi:

   ```
   python app.py
   ```

4. Di jendela yang terbuka:
   - **Browse Folder ZIP** → pilih folder yang berisi file-file `.zip` (sama seperti folder `input/` di versi lama)
   - **Browse Template Excel** → pilih file Template Excel Master
   - **Browse Folder Output** → pilih folder tujuan hasil digabung
   - Klik **Process** → progress bar & log berjalan realtime
   - Setelah selesai, klik **Open Output Folder** untuk langsung membuka folder hasil

## Yang TIDAK berubah dari versi asli

Seluruh isi `backend/report_processor.py` — cara mapping header, cara deteksi baris data, cara cek duplikat, cara menggabungkan ke master — **identik** dengan `main.py` versi lama. Sudah diuji dengan membandingkan hasil proses file yang sama pada kedua versi, hasilnya sama persis.

Perubahan yang dilakukan **hanya struktural**:
- Path (`MASTER_FILE`, `OUTPUT_FILE`) yang tadinya hardcoded, sekarang diisi lewat GUI.
- `print(...)` diganti jadi callback log, supaya bisa tampil realtime di GUI.
- Import `win32com.client` dibungkus try/except — kalau tidak tersedia (misalnya nanti di-build untuk macOS), fitur konversi `.xls` lama akan mencatat error yang jelas di log dan **melewati file itu saja**, bukan meng-crash seluruh aplikasi.

## Belum dikerjakan (langkah selanjutnya)

1. **Packaging ke `.exe` (Windows) / `.app` (macOS)** menggunakan PyInstaller — belum dibuatkan `.spec`-nya, ini Tahap 5 di rencana kita. Kalau siap lanjut, saya akan buatkan konfigurasi build-nya.
2. **Investigasi bug ZIP 1-bulan** yang datanya kosong — masih tertunda, akan dilanjutkan setelah Anda kirim daftar isi ZIP-nya.
3. Ikon aplikasi (`.ico` untuk Windows, `.icns` untuk macOS) belum dibuat — bisa ditambahkan saat packaging.
