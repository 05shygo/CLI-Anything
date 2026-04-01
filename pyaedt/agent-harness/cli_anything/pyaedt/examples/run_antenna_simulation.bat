@echo off
REM Printed Dipole Antenna HFSS Simulation Runner
REM Run in Windows CMD or PowerShell (NOT WSL/Git Bash)

echo ==============================================================
echo Printed Dipole Antenna HFSS Simulation
echo ==============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Find Python in Windows
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found in PATH
    echo Please install Python or add it to your PATH
    pause
    exit /b 1
)

echo Using Python:
python --version
echo.

REM Run the simulation
echo Starting simulation...
echo.

python run_antenna_simulation.py

if %errorlevel% neq 0 (
    echo.
    echo ==============================================================
    echo SIMULATION FAILED
    echo ==============================================================
    echo.
    echo Troubleshooting:
    echo 1. Make sure AEDT is installed and licensed
    echo 2. Verify PyAEDT is installed: pip install ansys-aedt-core
    echo 3. Try running directly in CMD first to see full error messages
    pause
) else (
    echo.
    echo ==============================================================
    echo SIMULATION COMPLETED SUCCESSFULLY
    echo ==============================================================
)

pause
