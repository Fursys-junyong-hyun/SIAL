@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "APP_MAIN=%SCRIPT_DIR%src\outsourced_inventory_confirmation\main.py"
cd /d "%SCRIPT_DIR%"

if not exist "%APP_MAIN%" (
    echo App entry file not found.
    echo %APP_MAIN%
    exit /b 1
)

set "PYTHON_EXE=%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
if exist "%PYTHON_EXE%" goto run_app

where python >nul 2>nul
if errorlevel 1 (
    echo Python executable not found.
    exit /b 1
)

for /f "delims=" %%I in ('where python') do (
    set "PYTHON_EXE=%%I"
    goto run_app
)

echo Python executable not found.
exit /b 1

:run_app
"%PYTHON_EXE%" "%APP_MAIN%"
