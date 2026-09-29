@echo off
setlocal

cd /d "%~dp0"

echo Smartsheet Export
echo ======================================
echo.

where py >nul 2>nul

if %errorlevel%==0 (
    .\SmartsheetExport.exe
)

if errorlevel 1 (
    echo.
    echo Export failed. Review the error above.
) else (
    echo.
    echo Export completed successfully.
)

echo.
pause

endlocal