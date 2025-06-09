@echo off
echo ===============================================
echo Web Scheduling System - Windows Service Installer
echo Using NSSM (Non-Sucking Service Manager)
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

:: Set PowerShell execution policy
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" >nul 2>&1

:: Set variables
set SERVICE_NAME=WebSchedulingSystem
set CURRENT_DIR=%~dp0
set PYTHON_EXE=%CURRENT_DIR%scheduling_env\Scripts\python.exe
set APP_FILE=%CURRENT_DIR%web_scheduling_system.py
set NSSM_EXE=%CURRENT_DIR%nssm\win64\nssm.exe

:: Check if NSSM exists
echo [STEP 1/5] Checking NSSM tool...
if not exist "%NSSM_EXE%" (
    echo [INFO] NSSM not found, downloading...
    
    :: Create nssm directory
    if not exist "nssm" mkdir nssm
    cd nssm
    
    :: Download NSSM (using PowerShell)
    echo [INFO] Downloading NSSM 2.24...
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://nssm.cc/release/nssm-2.24.zip' -OutFile 'nssm.zip'}"
    
    if exist "nssm.zip" (
        echo [INFO] Extracting NSSM...
        powershell -Command "Expand-Archive -Path 'nssm.zip' -DestinationPath '.' -Force"
        if exist "nssm-2.24" (
            xcopy /E /Y "nssm-2.24\*" "." >nul
            rmdir /s /q "nssm-2.24" 2>nul
        )
        del "nssm.zip" 2>nul
    ) else (
        echo ERROR: Failed to download NSSM
        pause
        exit /b 1
    )
    
    cd "%CURRENT_DIR%"
)

if not exist "%NSSM_EXE%" (
    echo ERROR: NSSM tool not found at: %NSSM_EXE%
    pause
    exit /b 1
)
echo [SUCCESS] NSSM tool ready
echo.

:: Check Python virtual environment
echo [STEP 2/5] Checking Python virtual environment...
if not exist "%PYTHON_EXE%" (
    echo [INFO] Python virtual environment not found, creating...
    
    :: Check Python
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: Python not found in PATH
        echo Please install Python 3.10+ and add it to PATH
        pause
        exit /b 1
    )
    
    :: Create virtual environment
    echo [INFO] Creating virtual environment...
    python -m venv scheduling_env
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    
    :: Install dependencies
    echo [INFO] Installing dependencies...
    "%CURRENT_DIR%scheduling_env\Scripts\python.exe" -m pip install --upgrade pip
    
    if exist "requirements.txt" (
        "%CURRENT_DIR%scheduling_env\Scripts\pip.exe" install -r requirements.txt
    ) else (
        "%CURRENT_DIR%scheduling_env\Scripts\pip.exe" install flask pyodbc pandas openpyxl numpy
    )
    
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
    
    echo [SUCCESS] Virtual environment created and configured
) else (
    echo [SUCCESS] Python virtual environment found
)
echo.

:: Check application file
echo [STEP 3/5] Checking application file...
if not exist "%APP_FILE%" (
    echo ERROR: web_scheduling_system.py not found
    pause
    exit /b 1
)
echo [SUCCESS] Application file found
echo.

:: Stop and remove existing service (if exists)
echo [STEP 4/5] Managing existing service...
sc query "%SERVICE_NAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Found existing service, removing...
    "%NSSM_EXE%" stop "%SERVICE_NAME%" >nul 2>&1
    timeout /t 3 >nul
    "%NSSM_EXE%" remove "%SERVICE_NAME%" confirm >nul 2>&1
    echo [INFO] Existing service removed
)
echo.

:: Install new service
echo [STEP 5/5] Installing Windows service...
"%NSSM_EXE%" install "%SERVICE_NAME%" "%PYTHON_EXE%" "%APP_FILE%"
if %errorlevel% neq 0 (
    echo ERROR: Service installation failed
    pause
    exit /b 1
)

:: Configure service parameters
echo [INFO] Configuring service parameters...
"%NSSM_EXE%" set "%SERVICE_NAME%" DisplayName "Web Scheduling System"
"%NSSM_EXE%" set "%SERVICE_NAME%" Description "Flask-based Web Scheduling Management System"
"%NSSM_EXE%" set "%SERVICE_NAME%" Start SERVICE_AUTO_START
"%NSSM_EXE%" set "%SERVICE_NAME%" AppDirectory "%CURRENT_DIR:~0,-1%"

:: Configure environment variables for production
"%NSSM_EXE%" set "%SERVICE_NAME%" AppEnvironmentExtra "FLASK_ENV=production" "FLASK_DEBUG=0"

:: Configure logging
"%NSSM_EXE%" set "%SERVICE_NAME%" AppStdout "%CURRENT_DIR:~0,-1%\logs\service_output.log"
"%NSSM_EXE%" set "%SERVICE_NAME%" AppStderr "%CURRENT_DIR:~0,-1%\logs\service_error.log"
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRotateFiles 1
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRotateOnline 0
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRotateSeconds 86400
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRotateBytes 1048576

:: Create logs directory
if not exist "logs" mkdir logs

:: Configure service restart policy
"%NSSM_EXE%" set "%SERVICE_NAME%" AppExit Default Restart
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRestartDelay 60000

echo [SUCCESS] Service installed successfully
echo.

:: Start service
echo Starting service...
set /p "start_service=Do you want to start the service now? (Y/N): "
if /i "%start_service%" neq "Y" (
    echo Service installed but not started
    echo You can start it with: net start %SERVICE_NAME%
    pause
    exit /b 0
)

net start "%SERVICE_NAME%"
if %errorlevel% neq 0 (
    echo ERROR: Service failed to start
    echo Please check log file: %CURRENT_DIR%logs\service_error.log
    pause
    exit /b 1
)

echo [SUCCESS] Service started successfully!
echo.
echo ===============================================
echo Windows Service Installation Complete!
echo ===============================================
echo.
echo Service Name: %SERVICE_NAME%
echo Web Interface: http://localhost:5100
echo Log Directory: %CURRENT_DIR%logs\
echo.
echo Service Management Commands:
echo   Start: net start %SERVICE_NAME%
echo   Stop: net stop %SERVICE_NAME%
echo   Status: sc query %SERVICE_NAME%
echo.
pause 