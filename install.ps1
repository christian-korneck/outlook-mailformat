# Installs the mailformat COM add-in for the current user (no admin rights needed)
# into %LOCALAPPDATA%\mailformat and registers it from there.
# Works from the source tree (after build.cmd) and from the installed folder.
# Close Outlook first, then run install.cmd.
$ErrorActionPreference = 'Stop'

$clsid   = '{6F3B5C2A-8E1D-4C7B-9A54-2D7E1F0B3A91}'
$progId  = 'mailformat.Connect'
$class   = 'MailformatAddin.Connect'
$target  = Join-Path $env:LOCALAPPDATA 'mailformat'
$scripts = 'install.ps1', 'install.cmd', 'uninstall.ps1', 'uninstall.cmd'

$srcDll = @((Join-Path $PSScriptRoot 'mailformat.dll'), (Join-Path $PSScriptRoot 'bin\mailformat.dll')) |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $srcDll) { throw "mailformat.dll not found next to install.ps1 or in bin\ - run build.cmd first." }
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { Write-Host 'Outlook is running. Close Outlook first, then run install again.' -ForegroundColor Red; exit 1 }

# Copy the add-in and the scripts into the install folder (skip files that are already there).
New-Item -ItemType Directory -Path $target -Force | Out-Null
foreach ($src in @($srcDll) + ($scripts | ForEach-Object { Join-Path $PSScriptRoot $_ })) {
    $dst = Join-Path $target (Split-Path $src -Leaf)
    if ([IO.Path]::GetFullPath($src) -ne [IO.Path]::GetFullPath($dst)) { Copy-Item -Path $src -Destination $dst -Force }
}

$dll      = Join-Path $target 'mailformat.dll'
$asmName  = [Reflection.AssemblyName]::GetAssemblyName($dll)
$codeBase = ([Uri]$dll).AbsoluteUri

function Set-Key($path, $values) {
    New-Item -Path $path -Force | Out-Null
    foreach ($name in $values.Keys) {
        $v = $values[$name]
        if ($name -eq '') { Set-Item -Path $path -Value $v; continue }
        $type = if ($v -is [int] -or $v -is [long]) { 'DWord' } else { 'String' }
        New-ItemProperty -Path $path -Name $name -Value $v -PropertyType $type -Force | Out-Null
    }
}

# COM registration (per user).
$classes = 'HKCU:\Software\Classes'
Set-Key "$classes\$progId"                     @{ '' = $class }
Set-Key "$classes\$progId\CLSID"               @{ '' = $clsid }
Set-Key "$classes\CLSID\$clsid"                @{ '' = $class }
Set-Key "$classes\CLSID\$clsid\ProgId"         @{ '' = $progId }
Set-Key "$classes\CLSID\$clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" @{}
Set-Key "$classes\CLSID\$clsid\InprocServer32" @{
    '' = 'mscoree.dll'; ThreadingModel = 'Both'; Class = $class
    Assembly = $asmName.FullName; RuntimeVersion = 'v4.0.30319'; CodeBase = $codeBase
}

# Outlook add-in registration.
Set-Key "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId" @{
    FriendlyName = 'mailformat'
    Description  = 'Wrap and quote selected text for mailing list replies.'
    LoadBehavior = 3
}

# Windows "Installed apps" entry.
$sizeKb  = [int][Math]::Ceiling(((Get-ChildItem $target -File | Measure-Object Length -Sum).Sum) / 1KB)
$uninst  = @{
    DisplayName     = 'mailformat for Outlook'
    DisplayVersion  = $asmName.Version.ToString(3)
    Publisher       = 'Christian Korneck <christian@korneck.de>'
    InstallLocation = $target
    InstallDate     = (Get-Date -Format 'yyyyMMdd')
    EstimatedSize   = $sizeKb
    NoModify        = 1
    NoRepair        = 1
    UninstallString = '"' + (Join-Path $target 'uninstall.cmd') + '"'
}
$outlook = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue).'(default)'
if ($outlook) { $uninst.DisplayIcon = "$outlook,0" }
Set-Key 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\mailformat' $uninst

Write-Host "Installed mailformat $($asmName.Version.ToString(3)) to $target. Start Outlook."
