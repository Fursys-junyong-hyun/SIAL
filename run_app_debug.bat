@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "APP_MAIN=%SCRIPT_DIR%src\outsourced_inventory_confirmation\main.py"
cd /d "%SCRIPT_DIR%"

echo SCRIPT_DIR=%SCRIPT_DIR%
echo APP_MAIN=%APP_MAIN%

if not exist "%APP_MAIN%" (
    echo App entry file not found.
    pause
    exit /b 1
)

set "PYTHON_EXE=%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
if exist "%PYTHON_EXE%" goto run_app

for /f "delims=" %%I in ('where python 2^>nul') do (
    set "PYTHON_EXE=%%I"
    goto run_app
)

echo Python executable not found.
pause
exit /b 1

:run_app
echo PYTHON_EXE=%PYTHON_EXE%
"%PYTHON_EXE%" "%APP_MAIN%"
echo.
echo ExitCode=%ERRORLEVEL%
pause
