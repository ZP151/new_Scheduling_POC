@echo off
echo ===============================================
echo Web Scheduling System - Service Uninstaller
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
echo [SUCCESS] Administrator privileges confirmed
echo.

setlocal enabledelayedexpansion

:: Set PowerShell execution policy
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" >nul 2>&1

:: Set variables
set SERVICE_NAME=WebSchedulingSystem
set CURRENT_DIR=%~dp0
set NSSM_EXE=%CURRENT_DIR%nssm\win64\nssm.exe

echo [STEP 1/4] Checking if service exists...
sc query "%SERVICE_NAME%" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Service '%SERVICE_NAME%' is not installed
    echo Nothing to uninstall
    goto :end
)
echo [SUCCESS] Service found: %SERVICE_NAME%
echo.

echo [STEP 2/4] Checking current service status...
for /f "tokens=3" %%i in ('sc query "%SERVICE_NAME%" ^| find "STATE"') do set SERVICE_STATE=%%i
echo [INFO] Current service state: !SERVICE_STATE!
echo.

echo [STEP 3/4] Stopping service...
if "!SERVICE_STATE!"=="RUNNING" (
    echo [INFO] Stopping running service...
    net stop "%SERVICE_NAME%"
    if %errorlevel% equ 0 (
        echo [SUCCESS] Service stopped successfully
    ) else (
        echo [WARNING] Failed to stop service normally, forcing stop...
        sc stop "%SERVICE_NAME%" >nul 2>&1
        timeout /t 5 >nul
    )
) else (
    echo [INFO] Service is not running
)

:: Kill any remaining Python processes related to our service
echo [INFO] Checking for related processes...
tasklist /FI "IMAGENAME eq python.exe" | find /I "scheduling_env" >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Found related Python processes, terminating...
    wmic process where "CommandLine like '%%scheduling_env%%'" delete >nul 2>&1
)

:: Wait a moment for processes to terminate
timeout /t 3 >nul
echo.

echo [STEP 4/4] Removing service...
if exist "%NSSM_EXE%" (
    echo [INFO] Using NSSM to remove service...
    "%NSSM_EXE%" remove "%SERVICE_NAME%" confirm
    if %errorlevel% equ 0 (
        echo [SUCCESS] Service removed successfully using NSSM
    ) else (
        echo [WARNING] NSSM removal failed, trying alternative method...
        sc delete "%SERVICE_NAME%"
        if %errorlevel% equ 0 (
            echo [SUCCESS] Service removed using sc delete
        ) else (
            echo [ERROR] Failed to remove service
        )
    )
) else (
    echo [INFO] NSSM not found, using Windows sc command...
    sc delete "%SERVICE_NAME%"
    if %errorlevel% equ 0 (
        echo [SUCCESS] Service removed successfully
    ) else (
        echo [ERROR] Failed to remove service
    )
)
echo.

echo [OPTIONAL] Additional cleanup...
set /p "cleanup=Do you want to clean up logs and temporary files? (Y/N): "
if /i "%cleanup%"=="Y" (
    echo [INFO] Cleaning up log files...
    if exist "logs" (
        del /Q "logs\*.log" 2>nul
        echo [SUCCESS] Log files cleaned
    )
    
    echo [INFO] Killing any remaining processes on port 5100...
    for /f "tokens=5" %%p in ('netstat -ano ^| find ":5100"') do (
        taskkill /PID %%p /F >nul 2>&1
    )
    echo [SUCCESS] Port cleanup completed
)
echo.

echo ===============================================
echo Service Uninstallation Complete!
echo ===============================================
echo.
echo Summary:
echo - Service '%SERVICE_NAME%' has been removed
echo - Processes terminated
echo - Port 5100 cleaned up
echo.
echo To reinstall the service, run: install_as_service.bat
echo To check system status, run: quick_check.bat
echo.

:end
pause 