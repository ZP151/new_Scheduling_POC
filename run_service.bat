@echo off
echo ===============================================
echo Web Scheduling System - One Click Setup
echo ===============================================
echo.

:: Check Python
echo [STEP 1/4] Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.1+ from https://python.org
    pause
    exit /b 1
)
echo [SUCCESS] Python detected
echo.

:: Setup virtual environment
echo [STEP 2/4] Setting up virtual environment...
set VENV_NAME=scheduling_env

if exist "%VENV_NAME%" (
    echo [INFO] Removing old virtual environment...
    rmdir /s /q "%VENV_NAME%" 2>nul
)

echo [INFO] Creating new virtual environment...
python -m venv %VENV_NAME%
if %errorlevel% neq 0 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

call %VENV_NAME%\Scripts\activate.bat

:: Create/update requirements.txt
echo [STEP 3/4] Installing dependencies...
if not exist "requirements.txt" (
    echo numpy==1.24.3 > requirements.txt
    echo pandas==2.0.3 >> requirements.txt
    echo pyodbc==4.0.39 >> requirements.txt
    echo sqlalchemy==2.0.19 >> requirements.txt
    echo flask==2.3.2 >> requirements.txt
    echo openpyxl==3.1.2 >> requirements.txt
    echo werkzeug==2.3.6 >> requirements.txt
    echo jinja2==3.1.2 >> requirements.txt
    echo markupsafe==2.1.3 >> requirements.txt
    echo itsdangerous==2.1.2 >> requirements.txt
)

python -m pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    echo Please check your internet connection
    pause
    exit /b 1
)
echo [SUCCESS] Dependencies installed
echo.

:: Test the application
echo [STEP 4/4] Testing application...
if not exist "web_scheduling_system.py" (
    echo ERROR: web_scheduling_system.py not found
    echo Please ensure the main application file is in this directory
    pause
    exit /b 1
)

echo [INFO] Testing database connection...
timeout /t 2 >nul

echo.
echo ===============================================
echo Installation Complete! Starting Application...
echo ===============================================
echo.

set /p "start_now=Do you want to start the application now? (Y/N): "
if /i "%start_now%" neq "Y" (
    echo.
    echo To start later, run: run_manual.bat
    echo Web interface: http://localhost:5100
    pause
    exit /b 0
)

echo.
echo [INFO] Starting Web Scheduling System...
echo [INFO] The application will run in this window
echo [INFO] Press Ctrl+C to stop the application
echo [INFO] Web interface will be available at: http://localhost:5100
echo.
echo ===============================================
echo Application is starting...
echo ===============================================
echo.

:: Run the application
python web_scheduling_system.py

:: If we reach here, the application has stopped
echo.
echo ===============================================
echo Application has stopped
echo ===============================================
pause 