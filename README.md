# Web Scheduling System

A Flask-based web application for automated course scheduling with Excel export functionality.

## Features

- Hierarchical resource selection (Term → Session → Campus → Academic Groups → Subjects → Courses)
- Multi-academic group support with cross-departmental subjects
- Smart room allocation based on course requirements
- Interactive time slot management (7/49 slots enabled by default)
- Automated scheduling with conflict detection
- Excel export with 36 standard columns
- SQL Server database integration

## Quick Start

### Option 1: One-Click Setup (Recommended)
```bash
run_service.bat
```
This script will:
- Check Python installation
- Create virtual environment
- Install dependencies
- Start the application

### Option 2: Windows Service Installation
```bash
# Install as Windows service (requires admin privileges)
install_as_service.bat

# Uninstall service
uninstall_service.bat
```

## Requirements

- Python 3.10+
- SQL Server with ODBC Driver 17
- Windows OS (for batch scripts)

## Database Configuration

Edit the database settings in `web_scheduling_system.py`:
```python
DB_SERVER = "localhost"
DB_NAME = "TestSchedulingDB" 
DB_USER = "SchedulingAppUser"
DB_PASSWORD = "SchedulingApp2025!"
```

## Access

- Web Interface: http://localhost:5100
- Excel files are generated in the application directory

## File Operations

- **Export**: Generates Excel files with 36 standard columns
- **Download**: Direct file download from web interface
- **Import**: File validation and processing simulation

## Service Management

```bash
net start WebSchedulingSystem    # Start service
net stop WebSchedulingSystem     # Stop service
sc query WebSchedulingSystem     # Check status
```

## Architecture

- **Frontend**: HTML/CSS/JavaScript with jQuery
- **Backend**: Flask REST API
- **Database**: SQL Server with SQLAlchemy + PyODBC
- **Export**: pandas + openpyxl for Excel generation 