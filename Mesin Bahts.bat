@echo off
rem Membuka antarmuka Mesin Bahts (bahts_gui.py). Kalau Python belum ada, dipasang otomatis lewat winget.
setlocal
cd /d "%~dp0"
set "APP=%~dp0bahts_gui.py"
set "PYW="

call :cari
if defined PYW goto :jalan

echo Python 3 belum ditemukan di komputer ini.
where winget >nul 2>nul
if errorlevel 1 goto :tanpa_winget

echo.
echo Memasang Python 3.12 lewat winget. Mungkin muncul jendela izin Windows: klik Yes.
echo Tunggu sampai selesai, jangan tutup jendela ini.
echo.
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
call :cari
if defined PYW goto :jalan

echo.
echo Python sudah dipasang, tetapi belum terbaca oleh jendela ini.
echo Tutup jendela ini lalu klik dua kali "Mesin Bahts.bat" sekali lagi.
pause
exit /b 1

:tanpa_winget
echo winget tidak tersedia, jadi Python tidak bisa dipasang otomatis.
echo Unduh Python 3.12 dari https://www.python.org/downloads/windows/
echo Saat memasang, centang "Add python.exe to PATH".
echo Setelah selesai, klik dua kali "Mesin Bahts.bat" lagi.
start "" "https://www.python.org/downloads/windows/"
pause
exit /b 1

:jalan
start "" %PYW% "%APP%"
exit /b 0

rem ---- mencari Python yang punya tkinter; hasilnya di variabel PYW ----
:cari
py -3 -c "import tkinter" >nul 2>nul
if not errorlevel 1 (
    set "PYW=pyw -3"
    exit /b 0
)
python -c "import tkinter" >nul 2>nul
if not errorlevel 1 (
    set "PYW=pythonw"
    exit /b 0
)
for %%V in (313 312 311 310) do (
    if not defined PYW if exist "%LocalAppData%\Programs\Python\Python%%V\pythonw.exe" set PYW="%LocalAppData%\Programs\Python\Python%%V\pythonw.exe"
    if not defined PYW if exist "%ProgramFiles%\Python%%V\pythonw.exe" set PYW="%ProgramFiles%\Python%%V\pythonw.exe"
)
exit /b 0
