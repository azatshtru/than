@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "backup_diff_gui.py"
    goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python "backup_diff_gui.py"
    goto :end
)

echo Python was not found. Please install Python 3 and try again.
pause
exit /b 1

:end
if errorlevel 1 pause
endlocal
