@echo off
cd /d "%~dp0\.."
set OUT=%CD%\release\Excel Analyst Tool

echo.
echo  Dong goi phat hanh -^> release\Excel Analyst Tool
echo.

if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"

copy /Y app.py "%OUT%\" >nul
copy /Y requirements.txt "%OUT%\" >nul
copy /Y "Mo Excel Analyst.bat" "%OUT%\" >nul
copy /Y "Cai dat lan dau.bat" "%OUT%\" >nul
copy /Y "Chay Server cho Team.bat" "%OUT%\" >nul
copy /Y "HUONG DAN NGUOI DUNG.md" "%OUT%\" >nul

xcopy /E /I /Q assets "%OUT%\assets" >nul
xcopy /E /I /Q modules "%OUT%\modules" >nul
xcopy /E /I /Q tools "%OUT%\tools" >nul
xcopy /E /I /Q .streamlit "%OUT%\.streamlit" >nul

set /p COPY_VENV=Copy venv (user khong can Python)? [y/N]: 
if /i "%COPY_VENV%"=="y" (
    if exist "venv" (
        echo  Dang copy venv...
        xcopy /E /I /Q venv "%OUT%\venv" >nul
        echo  Da copy venv.
    ) else (
        echo  Khong co thu muc venv.
    )
)

echo.
echo  XONG: %OUT%
echo.
pause
