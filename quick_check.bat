@echo off
echo ===============================================
echo Quick Service Configuration Check
echo ===============================================
echo.

setlocal enabledelayedexpansion

:: Set variables
set SERVICE_NAME=WebSchedulingSystem
set CURRENT_DIR=%~dp0
set PYTHON_EXE=%CURRENT_DIR%scheduling_env\Scripts\python.exe
set APP_FILE=%CURRENT_DIR%web_scheduling_system.py
set NSSM_EXE=%CURRENT_DIR%nssm\win64\nssm.exe

echo [CHECK 1] Virtual environment and dependencies...
if exist "%PYTHON_EXE%" (
    echo  Virtual environment Python: %PYTHON_EXE%
    echo  Python version:
    "%PYTHON_EXE%" --version
    
    echo.
    echo  Checking key packages:
    "%PYTHON_EXE%" -c "import flask; print(' Flask: OK')" 2>nul || echo " Flask: MISSING"
    "%PYTHON_EXE%" -c "import pyodbc; print(' pyodbc: OK')" 2>nul || echo " pyodbc: MISSING"
    "%PYTHON_EXE%" -c "import pandas; print(' pandas: OK')" 2>nul || echo " pandas: MISSING"
    "%PYTHON_EXE%" -c "import openpyxl; print(' openpyxl: OK')" 2>nul || echo " openpyxl: MISSING"
    "%PYTHON_EXE%" -c "import numpy; print(' numpy: OK')" 2>nul || echo " numpy: MISSING"
) else (
    echo  Virtual environment Python not found: %PYTHON_EXE%
    echo  Run: rebuild_venv.bat
)

echo.
echo [CHECK 2] Application file...
if exist "%APP_FILE%" (
    echo  Application file exists: %APP_FILE%
) else (
    echo  Application file not found: %APP_FILE%
)

echo.
echo [CHECK 3] NSSM tool...
if exist "%NSSM_EXE%" (
    echo  NSSM tool found: %NSSM_EXE%
) else (
    echo  NSSM tool not found: %NSSM_EXE%
)

echo.
echo [CHECK 4] Service status (basic check)...
sc query "%SERVICE_NAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo  Service is registered in Windows
    for /f "tokens=3" %%i in ('sc query "%SERVICE_NAME%" ^| find "STATE"') do (
        echo  Service state: %%i
    )
) else (
    echo  Service not found or not installed
    echo  Run: install_as_service.bat (as administrator)
)

echo.
echo [CHECK 5] Database connection test...
echo  Testing database connection...
"%PYTHON_EXE%" -c "from web_scheduling_system import WebSchedulingSystem; ws = WebSchedulingSystem(); print(' Database connection: OK')" 2>nul || echo " Database connection: FAILED"

echo.
echo ===============================================
echo  Quick Check Summary
echo ===============================================
echo.
echo Next steps if issues found:
echo 1. Missing packages: run rebuild_venv.bat
echo 2. Service not installed: run install_as_service.bat (as admin)
echo 3. Database issues: check database server and credentials
echo 4. Full service check: run check_service_config.bat (as admin)
echo.
pause 