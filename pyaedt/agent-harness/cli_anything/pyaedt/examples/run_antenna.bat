@echo off
REM Run Printed Dipole Antenna HFSS Simulation
REM Set environment and run the Python script

echo Setting AEDT environment variables...
set "AWP_ROOT193=C:\Program Files\ANSYS Inc\v193"
set "ANSYSEM_ROOT193=C:\Program Files\ANSYS Inc\v193"
set "PATH=%AWP_ROOT193%\Win64;%PATH%"

echo Running simulation...
cd /d "D:\work_tool\CLI-Anything\pyaedt\agent-harness\cli_anything\pyaedt\examples"
python run_antenna_simulation.py

pause
