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

:: Set variables
set SERVICE_NAME=WebSchedulingSystem
set CURRENT_DIR=%~dp0
set PYTHON_EXE=%CURRENT_DIR%scheduling_env\Scripts\python.exe
set APP_FILE=%CURRENT_DIR%web_scheduling_system.py
set NSSM_EXE=%CURRENT_DIR%nssm\win64\nssm.exe

:: Check if NSSM exists
echo [STEP 1/6] Checking NSSM tool...
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
        echo ERROR: Failed to download NSSM, please download manually and extract to nssm directory
        echo Download URL: https://nssm.cc/download
        pause
        exit /b 1
    )
    
    cd "%CURRENT_DIR%"
)

if not exist "%NSSM_EXE%" (
    echo ERROR: NSSM tool not found at: %NSSM_EXE%
    echo Please ensure NSSM is properly installed
    pause
    exit /b 1
)
echo [SUCCESS] NSSM tool ready
echo.

:: Check and create Python virtual environment
echo [STEP 2/6] Checking and creating Python virtual environment...
if not exist "%PYTHON_EXE%" (
    echo [INFO] Python virtual environment not found, creating...
    
    :: Set PowerShell execution policy
    echo [INFO] Setting PowerShell execution policy...
    powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" >nul 2>&1
    
    :: Check for Python 3.10 specifically
    echo [INFO] Checking for Python 3.10...
    python3.10 --version >nul 2>&1
    if %errorlevel% equ 0 (
        set PYTHON_CMD=python3.10
        echo [SUCCESS] Found Python 3.10
    ) else (
        python --version 2>&1 | find "3.10" >nul
        if %errorlevel% equ 0 (
            set PYTHON_CMD=python
            echo [SUCCESS] Found Python 3.10 (as default python)
        ) else (
            echo [WARNING] Python 3.10 not found, checking available Python versions...
            python --version >nul 2>&1
            if %errorlevel% neq 0 (
                echo ERROR: No Python found in PATH
                echo Please install Python 3.10 and add it to PATH
                echo Download from: https://www.python.org/downloads/release/python-3108/
                pause
                exit /b 1
            ) else (
                echo [WARNING] Using available Python version (recommended: Python 3.10)
                set PYTHON_CMD=python
            )
        )
    )
    
    :: Display Python version
    echo [INFO] Using Python version:
    %PYTHON_CMD% --version
    
    :: Create virtual environment with specified Python version
    echo [INFO] Creating virtual environment 'scheduling_env' with %PYTHON_CMD%...
    %PYTHON_CMD% -m venv scheduling_env
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment
        echo Please ensure Python 3.10 is properly installed
        pause
        exit /b 1
    )
    
    :: Install dependencies
    echo [INFO] Installing dependencies...
    
    :: Upgrade pip first
    echo [INFO] Upgrading pip...
    "%CURRENT_DIR%scheduling_env\Scripts\python.exe" -m pip install --upgrade pip
    
    :: Install basic requirements
    echo [INFO] Installing Flask and required packages...
    "%CURRENT_DIR%scheduling_env\Scripts\pip.exe" install flask==2.3.3 pyodbc==4.0.39
    
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
    
    :: Verify installation
    echo [INFO] Verifying installation...
    "%CURRENT_DIR%scheduling_env\Scripts\python.exe" --version
    "%CURRENT_DIR%scheduling_env\Scripts\python.exe" -c "import flask, pyodbc; print('✅ Dependencies installed successfully')"
    
    echo [SUCCESS] Virtual environment created and configured with Python 3.10
) else (
    echo [SUCCESS] Python virtual environment found
    echo [INFO] Current Python version in virtual environment:
    "%PYTHON_EXE%" --version
)

:: Verify Python executable
if not exist "%PYTHON_EXE%" (
    echo ERROR: Python executable still not found after setup
    pause
    exit /b 1
)
echo.

:: Check application file
echo [STEP 3/6] Checking application file...
if not exist "%APP_FILE%" (
    echo ERROR: web_scheduling_system.py not found
    echo Please ensure the application file is in the current directory
    pause
    exit /b 1
)
echo [SUCCESS] Application file found
echo.

:: Stop and remove existing service (if exists)
echo [STEP 4/6] Checking existing service...
sc query "%SERVICE_NAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Found existing service, stopping and removing...
    "%NSSM_EXE%" stop "%SERVICE_NAME%" >nul 2>&1
    timeout /t 3 >nul
    "%NSSM_EXE%" remove "%SERVICE_NAME%" confirm >nul 2>&1
    echo [INFO] Existing service removed
)
echo.

:: Install new service
echo [STEP 5/6] Installing Windows service...
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
echo [INFO] Setting production environment variables...
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
echo [STEP 6/6] Starting service...
set /p "start_service=Do you want to start the service now? (Y/N): "
if /i "%start_service%" neq "Y" (
    echo.
    echo Service installed but not started
    echo You can manage the service using:
    echo   - Start service: net start %SERVICE_NAME%
    echo   - Stop service: net stop %SERVICE_NAME%
    echo   - Service manager: services.msc
    echo   - Web interface: http://localhost:5100
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
echo Display Name: Web Scheduling System
echo Web Interface: http://localhost:5100
echo Log Directory: %CURRENT_DIR%logs\
echo.
echo Service Management Commands:
echo   Start service: net start %SERVICE_NAME%
echo   Stop service: net stop %SERVICE_NAME%
echo   Restart service: net stop %SERVICE_NAME% && net start %SERVICE_NAME%
echo   Remove service: %NSSM_EXE% remove %SERVICE_NAME% confirm
echo.
echo System Service Manager: services.msc
echo.
echo Waiting for service to start...
timeout /t 5 >nul

:: Check service status
sc query "%SERVICE_NAME%" | find "RUNNING" >nul
if %errorlevel% equ 0 (
    echo [SUCCESS] Service is running normally
    echo [INFO] You can now access the system at http://localhost:5100
) else (
    echo [WARNING] Service may not have started properly, please check logs
)

echo.
pause 