@echo off
REM FAIRe Metadata Filler – Windows launcher
REM Double-click this file to start the tool.

title FAIRe Metadata Filler

cd /d "%~dp0"

REM Check for Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    where python3 >nul 2>nul
    if %errorlevel% neq 0 (
        echo.
        echo ============================================================
        echo   Python not found.
        echo.
        echo   Please install Python from https://www.python.org/downloads/
        echo   Make sure to tick "Add Python to PATH" during installation,
        echo   then double-click this file again.
        echo ============================================================
        echo.
        pause
        exit /b 1
    )
    set PYTHON=python3
) else (
    set PYTHON=python
)

echo.
echo   FAIRe Metadata Filler
echo   Installing / checking dependencies...
echo.
%PYTHON% -m pip install --quiet --upgrade pdfplumber anthropic 2>nul

echo   Starting server at http://127.0.0.1:8787
echo   Opening browser...
echo.
echo   To stop the tool close this window.
echo.

REM Open browser after short delay
start "" /b cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:8787"

%PYTHON% server.py

pause
