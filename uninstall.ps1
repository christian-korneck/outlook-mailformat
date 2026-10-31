# Uninstalls the mailformat COM add-in for the current user: removes the registry
# entries and the install folder %LOCALAPPDATA%\mailformat.
# The width setting in HKCU\Software\mailformat is kept.
$ErrorActionPreference = 'Stop'

$clsid  = '{6F3B5C2A-8E1D-4C7B-9A54-2D7E1F0B3A91}'
$progId = 'mailformat.Connect'
$target = Join-Path $env:LOCALAPPDATA 'mailformat'

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { Write-Host 'Outlook is running. Close Outlook first, then uninstall again.' -ForegroundColor Red; exit 1 }

foreach ($p in "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId",
               "HKCU:\Software\Classes\CLSID\$clsid",
               "HKCU:\Software\Classes\$progId",
               'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\mailformat') {
    if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force }
}

# This script may itself be running from the install folder, so delete the folder
# from a detached process once we (and the calling uninstall.cmd) have exited.
if (Test-Path $target) {
    Start-Process -FilePath "$env:ComSpec" -WorkingDirectory $env:TEMP -WindowStyle Hidden `
        -ArgumentList "/c timeout /t 2 /nobreak >nul & rmdir /s /q `"$target`""
}

Write-Host 'Uninstalled mailformat.'
