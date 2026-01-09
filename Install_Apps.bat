@echo off
color 0A
title Generator Aplikasi PDF (.EXE)

echo ========================================================
echo        OTOMATISASI PEMBUATAN SUPER APP (.EXE)
echo ========================================================
echo.

:: 1. Cek keberadaan script Python Utama
if not exist "Main_App.py" (
    color 0C
    echo [ERROR] File 'Main_App.py' TIDAK DITEMUKAN!
    echo Pastikan nama file script gabungan Anda adalah: Main_App.py
    pause
    exit
)

:: 2. Instalasi Library (Update baru ada pypdf)
echo [1/3] Menginstal library: pandas, openpyxl, pyinstaller, pypdf...
echo.
py -m pip install pandas openpyxl pyinstaller pypdf --user --quiet
if %errorlevel% neq 0 python -m pip install pandas openpyxl pyinstaller pypdf --user --quiet

echo.
echo [2/3] Sedang merakit file .EXE (Mohon tunggu 1-3 menit)... 
echo.

:: 3. Perintah Membuat EXE
py -m PyInstaller --noconsole --onefile --clean --name "PDF_Tools_Master" Main_App.py
if %errorlevel% neq 0 python -m PyInstaller --noconsole --onefile --clean --name "PDF_Tools_Master" Main_App.py

:: 4. Pembersihan
echo.
echo [3/3] Membersihkan file sampah...

if exist "dist\PDF_Tools_Master.exe" (
    move /Y "dist\PDF_Tools_Master.exe" "."
)

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "PDF_Tools_Master.spec" del /f /q "PDF_Tools_Master.spec"

echo.
echo ========================================================
echo                    SELESAI!
echo ========================================================
echo File aplikasi baru: 'PDF_Tools_Master.exe' siap digunakan!
echo.
pause
