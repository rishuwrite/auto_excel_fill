@echo off
setlocal
cd /d "%~dp0"

set "PYTHON=%~dp0runtime\python.exe"
set "APP=%~dp0fbt_gui.py"

if not exist "%PYTHON%" (
    echo.
    echo ERROR: Portable Python runtime was not found.
    echo.
    pause
    exit /b 1
)

if not exist "%APP%" (
    echo.
    echo ERROR: fbt_gui.py was not found.
    echo.
    pause
    exit /b 1
)

"%PYTHON%" "%APP%"

if errorlevel 1 (
    echo.
    echo FBT Generator stopped with an error.
    echo.
    pause
)

endlocal