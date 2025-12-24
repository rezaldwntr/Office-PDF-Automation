@echo off
color 0A
title Generator Aplikasi PDF (.EXE)

echo ========================================================
echo        OTOMATISASI PEMBUATAN APLIKASI (.EXE)
echo ========================================================
echo.

:: 1. Cek keberadaan script Python
if not exist "PDF_Renamer_GUI.py" (
    color 0C
    echo [ERROR] File 'PDF_Renamer_GUI.py' TIDAK DITEMUKAN!
    echo.
    echo Pastikan Anda sudah menyimpan script Python dengan nama:
    echo PDF_Renamer_GUI.py
    echo di dalam folder yang sama dengan file ini.
    echo.
    pause
    exit
)

:: 2. Instalasi Library (Mencegah error 'Module Not Found')
echo [1/3] Mengecek dan menginstal library yang dibutuhkan...
echo        (Pandas, Openpyxl, PyInstaller)
echo.
py -m pip install pandas openpyxl pyinstaller --user --quiet
:: Baris bawah untuk backup jika perintah 'py' tidak jalan di komputer tertentu
if %errorlevel% neq 0 python -m pip install pandas openpyxl pyinstaller --user --quiet

echo.
echo [2/3] Sedang merakit file .EXE... 
echo        (Proses ini memakan waktu 1-3 menit, mohon bersabar)
echo.

:: 3. Perintah Membuat EXE
:: --onefile : Jadikan satu file tunggal
:: --noconsole : Hilangkan layar hitam CMD saat aplikasi dijalankan nanti
:: --clean : Bersihkan cache sebelum build
py -m PyInstaller --noconsole --onefile --clean --name "Aplikasi_Rename_PDF" PDF_Renamer_GUI.py
if %errorlevel% neq 0 python -m PyInstaller --noconsole --onefile --clean --name "Aplikasi_Rename_PDF" PDF_Renamer_GUI.py

:: 4. Pembersihan dan Pemindahan File
echo.
echo [3/3] Membersihkan file sampah dan merapikan hasil...

:: Pindahkan file EXE dari folder 'dist' ke folder utama
if exist "dist\Aplikasi_Rename_PDF.exe" (
    move /Y "dist\Aplikasi_Rename_PDF.exe" "."
)

:: Hapus folder sisa build yang tidak perlu
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Aplikasi_Rename_PDF.spec" del /f /q "Aplikasi_Rename_PDF.spec"

echo.
echo ========================================================
echo                    SELESAI!
echo ========================================================
echo File aplikasi bernama 'Aplikasi_Rename_PDF.exe' 
echo sudah siap digunakan di folder ini.
echo.
pause
