@echo off
REM ==========================================
REM Afiq Autoexcel Launcher
REM ==========================================

REM Move to the folder where this BAT file lives
cd /d "%~dp0"

REM Optional: show where we are (for debugging)
echo Working directory: %cd%

REM Run the Python script
python autotanggal.py

REM Keep window open so errors are visible
echo.
echo Script finished or stopped.
pause
