@echo off
setlocal EnableExtensions

REM ============================================================
REM VIOLATION AUTOMATION - Windows Runner
REM ============================================================

echo ============================================================
echo ViolationAUTOMATION
echo ============================================================
echo Start Time: %date% %time%
echo.

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

REM ------------------------------------------------------------
REM Ensure VIOLATIONS_DIR environment variable exists
REM ------------------------------------------------------------
if not defined VIOLATIONS_DIR (
    echo ERROR: VIOLATIONS_DIR is not set.
    echo.
    echo Run this once in Command Prompt:
    echo setx VIOLATIONS_DIR "C:\Users\arksecurity\Documents\Violations"
    goto :END
)

set "PROJECT_DIR=%VIOLATIONS_DIR%"
set "TARGET_SCRIPT=run_pull_violation.py"

echo Using Project Directory: %PROJECT_DIR%
echo.

REM ------------------------------------------------------------
REM Change to project directory
REM ------------------------------------------------------------
cd /d "%PROJECT_DIR%" || (
    echo ERROR: Failed to change directory to %PROJECT_DIR%
    goto :END
)

REM ------------------------------------------------------------
REM Check if script exists
REM ------------------------------------------------------------
if not exist "%TARGET_SCRIPT%" (
    echo ERROR: %TARGET_SCRIPT% not found in %PROJECT_DIR%
    goto :END
)

REM ------------------------------------------------------------
REM Use virtual environment if it exists
REM ------------------------------------------------------------
if exist ".venv\Scripts\python.exe" (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
    set "PYTHON_EXE=python"
)

echo Running %TARGET_SCRIPT%...
echo.

"%PYTHON_EXE%" "%TARGET_SCRIPT%" %*
set "PY_EXIT=%ERRORLEVEL%"

echo.
if %PY_EXIT% EQU 0 (
    echo ============================================================
    echo VIOLATION AUTOMATION COMPLETED SUCCESSFULLY
    echo ============================================================
) else (
    echo ============================================================
    echo ERROR: VIOLATION automation failed with error code %PY_EXIT%
    echo ============================================================
)

:END
echo.
echo End Time: %date% %time%
echo.
pause
endlocal