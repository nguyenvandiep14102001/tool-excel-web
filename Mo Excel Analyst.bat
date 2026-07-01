@echo off
cd /d "%~dp0"
title Excel Analyst Tool

if not exist "venv\Scripts\streamlit.exe" (
    echo.
    echo  Chua cai dat. Dang mo trinh cai dat lan dau...
    echo.
    call "%~dp0Cai dat lan dau.bat"
    if not exist "venv\Scripts\streamlit.exe" exit /b 1
)

echo.
echo  ============================================
echo   Excel Analyst Tool dang chay...
echo  ============================================
echo.
echo   Trinh duyet se mo tu dong.
echo   KHONG dong cua so den nay khi dang dung app.
echo.
echo   De thoat: dong cua so nay hoac nhan Ctrl+C
echo  ============================================
echo.

call venv\Scripts\activate.bat
venv\Scripts\streamlit.exe run app.py

echo.
echo  App da dung.
pause
