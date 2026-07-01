@echo off
cd /d "%~dp0"
title Excel Analyst - Cai dat lan dau

echo.
echo  ============================================
echo   Excel Analyst Tool - Cai dat lan dau
echo  ============================================
echo.
echo  Chi can chay file nay MOT LAN truoc khi dung app.
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo  [LOI] Chua co Python tren may.
    echo.
    echo  Cai Python 3.10+ tai: https://www.python.org/downloads/
    echo  Khi cai, danh dau "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

echo  [1/3] Tao moi truong ao (venv)...
if not exist "venv\Scripts\python.exe" (
    python -m venv venv
    if errorlevel 1 (
        echo  [LOI] Khong tao duoc venv.
        pause
        exit /b 1
    )
) else (
    echo        venv da ton tai - bo qua.
)

echo  [2/3] Cai thu vien...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip -q
pip install -r requirements.txt
if errorlevel 1 (
    echo  [LOI] Cai dat thu vien that bai.
    pause
    exit /b 1
)

echo  [3/3] Tao shortcut tren Desktop (neu co the)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\create_shortcut.ps1" 2>nul

echo.
echo  ============================================
echo   Cai dat XONG!
echo.
echo   Tu gio chi can double-click:
echo   "Mo Excel Analyst.bat"
echo  ============================================
echo.
pause
