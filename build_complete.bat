@echo off
setlocal
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python build.py
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        py build.py
    ) else (
        echo ERROR: Python not found. Run build.py manually.
        exit /b 1
    )
)
if %errorlevel% neq 0 exit /b %errorlevel%
echo.
echo BUILD COMPLETE.
