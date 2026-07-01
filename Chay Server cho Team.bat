@echo off
cd /d "%~dp0"
title Excel Analyst - Server for Team

if not exist "venv\Scripts\streamlit.exe" (
    echo  Chua cai dat. Chay "Cai dat lan dau.bat" truoc.
    pause
    exit /b 1
)

echo.
echo  ============================================
echo   Excel Analyst - SERVER
echo  ============================================
echo.
echo   May nay chay app cho ca team.
echo   User chi can mo trinh duyet - khong can file source.
echo.
echo   Tren may nay:
echo     http://localhost:8501
echo.
echo   May khac trong mang LAN:
powershell -NoProfile -Command "Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike '127.*' -and $_.PrefixOrigin -ne 'WellKnown' } | ForEach-Object { Write-Host ('     http://' + $_.IPAddress + ':8501') }"
echo.
echo   Gui link tren cho dong nghiep.
echo   KHONG dong cua so nay khi team dang dung.
echo  ============================================
echo.

call venv\Scripts\activate.bat
venv\Scripts\streamlit.exe run app.py --server.address=0.0.0.0 --server.port=8501 --server.headless=true

echo.
pause
