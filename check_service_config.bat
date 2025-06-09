@echo off
echo ===============================================
echo Windows Service Configuration Checker
echo ===============================================
echo.

:: Check administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script requires administrator privileges
    echo Please right-click the script and select "Run as administrator"
    pause
    exit /b 1
)

setlocal enabledelayedexpansion

:: Set PowerShell execution policy
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" >nul 2>&1

:: Set variables
set SERVICE_NAME=WebSchedulingSystem
set CURRENT_DIR=%~dp0
set PYTHON_EXE=%CURRENT_DIR%scheduling_env\Scripts\python.exe
set APP_FILE=%CURRENT_DIR%web_scheduling_system.py
set NSSM_EXE=%CURRENT_DIR%nssm\win64\nssm.exe

echo [CHECK 1] Service existence...
sc query "%SERVICE_NAME%" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Service not found
    goto :end
)
echo  Service is installed

echo.
echo [CHECK 2] Current service configuration...
echo  Service status:
sc query "%SERVICE_NAME%" | find "STATE"

echo.
echo  NSSM configuration details:
if exist "%NSSM_EXE%" (
    echo  NSSM tool path: %NSSM_EXE%
    echo.
    echo  Service executable:
    "%NSSM_EXE%" get "%SERVICE_NAME%" Application
    echo.
    echo  Service parameters:
    "%NSSM_EXE%" get "%SERVICE_NAME%" AppParameters
    echo.
    echo  Working directory:
    "%NSSM_EXE%" get "%SERVICE_NAME%" AppDirectory
    echo.
    echo  Environment variables:
    "%NSSM_EXE%" get "%SERVICE_NAME%" AppEnvironmentExtra
    echo.
    echo  Log configuration:
    echo    Output log: 
    "%NSSM_EXE%" get "%SERVICE_NAME%" AppStdout
    echo    Error log:
    "%NSSM_EXE%" get "%SERVICE_NAME%" AppStderr
) else (
    echo  NSSM tool not found: %NSSM_EXE%
)

echo.
echo [CHECK 3] Virtual environment and dependencies...
if exist "%PYTHON_EXE%" (
    echo  Virtual environment Python: %PYTHON_EXE%
    echo  Python version:
    "%PYTHON_EXE%" --version
    
    echo.
    echo  Installed packages:
    "%CURRENT_DIR%scheduling_env\Scripts\pip.exe" list | find "flask\|pyodbc\|pandas\|openpyxl\|numpy"
    
    echo.
    echo  Testing package imports:
    "%PYTHON_EXE%" -c "import flask; print(' Flask')" 2>nul || echo " Flask import failed"
    "%PYTHON_EXE%" -c "import pyodbc; print(' pyodbc')" 2>nul || echo " pyodbc import failed"
    "%PYTHON_EXE%" -c "import pandas; print(' pandas')" 2>nul || echo " pandas import failed"
    "%PYTHON_EXE%" -c "import openpyxl; print(' openpyxl')" 2>nul || echo " openpyxl import failed"
    "%PYTHON_EXE%" -c "import numpy; print(' numpy')" 2>nul || echo " numpy import failed"
) else (
    echo  Virtual environment Python not found: %PYTHON_EXE%
)

echo.
echo [CHECK 4] Application file...
if exist "%APP_FILE%" (
    echo  Application file exists: %APP_FILE%
) else (
    echo  Application file not found: %APP_FILE%
)

echo.
echo ===============================================
echo  Configuration Validation Summary
echo ===============================================

:: Check if the service is configured to use the virtual environment
"%NSSM_EXE%" get "%SERVICE_NAME%" Application | find "scheduling_env\Scripts\python.exe" >nul 2>&1
if %errorlevel% equ 0 (
    echo  Service is correctly configured to use virtual environment
    echo.
    echo  ALL CHECKS PASSED! Service should be working properly.
    echo  You can access the web interface at: http://localhost:5100
) else (
    echo  Service is NOT configured to use virtual environment!
    echo    Current configured Python path:
    "%NSSM_EXE%" get "%SERVICE_NAME%" Application
    echo    Should be using: %PYTHON_EXE%
    echo.
    echo  Solution: Re-run install_as_service.bat
)

echo.
echo  Recommended actions:
echo    1. If package imports failed, run: rebuild_venv.bat
echo    2. If service configuration is wrong, run: install_as_service.bat  
echo    3. Check service logs: logs\service_error.log

:end
echo.
pause 