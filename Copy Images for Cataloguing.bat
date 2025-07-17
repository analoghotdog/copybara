@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0CopyImages.ps1"
echo.
echo Script execution completed. Press any key to exit.
pause >nul