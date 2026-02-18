@echo off
REM ============================================================
REM PARKED TRUCKS REPORT AUTOMATION
REM ============================================================
REM
REM This script:
REM - Activates the project virtual environment (if present)
REM - Generates the Parked Trucks Excel + Summary Image
REM - Sends the report via email
REM
REM Can be run:
REM - Manually (double click)
REM - Via Windows Task Scheduler (e.g. 10:00 AM)
REM ============================================================

echo ============================================================
echo EVENTS TRUCKS REPORT AUTOMATION
echo ============================================================
echo Start Time: %date% %time%
echo.

REM ------------------------------------------------------------
REM Project directory
REM ------------------------------------------------------------
set PROJECT_DIR=C:\Users\SAMA\3D Objects\violations

cd /d "%PROJECT_DIR%" || (
    echo ERROR: Failed to change directory to %PROJECT_DIR%
    goto :END
)

REM ------------------------------------------------------------
REM Activate virtual environment if it exists
REM ------------------------------------------------------------
if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call ".venv\Scripts\activate.bat"
) else (
    echo WARNING: Virtual environment not found, using system Python
)

echo.
echo Running Event Trucks report...
echo.

REM ------------------------------------------------------------
REM Run the report runner (this sends email)
REM ------------------------------------------------------------
REM Run with venv Python if it exists, otherwise system Python
REM ------------------------------------------------------------
if exist ".venv\Scripts\python.exe" (
    echo Using virtual environment Python...
    ".venv\Scripts\python.exe" pull_galco_all_events.py --day today
) else (
    echo Using system Python...
    python pull_galco_all_events.py --day today
)

REM ------------------------------------------------------------
REM Check result
REM ------------------------------------------------------------
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================================
    echo PARKED TRUCKS REPORT COMPLETED SUCCESSFULLY
    echo ============================================================
) else (
    echo.
    echo ============================================================
    echo ERROR: Parked Trucks report failed with error code %ERRORLEVEL%
    echo ============================================================
)

:END
echo.
echo End Time: %date% %time%
echo ============================================================
echo.

REM ------------------------------------------------------------
REM Pause only if run manually
REM ------------------------------------------------------------
if "%1"=="" (
    pause
)
