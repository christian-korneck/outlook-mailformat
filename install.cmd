@echo off
rem Installs the mailformat Outlook add-in into %LOCALAPPDATA%\mailformat for the current user.
rem From the source tree it builds first if needed; from the installed folder it re-registers in place.
setlocal
cd /d "%~dp0"
if not exist mailformat.dll if not exist bin\mailformat.dll if exist build.cmd (call build.cmd || exit /b 1)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 pause
