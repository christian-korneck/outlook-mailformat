@echo off
rem Uninstalls the mailformat Outlook add-in (registry entries and %LOCALAPPDATA%\mailformat).
rem Kept on one line with exit /b: this file may be deleted while it runs.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" & (if errorlevel 1 pause) & exit /b
