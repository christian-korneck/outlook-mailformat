@echo off
rem Builds mailformat.dll with the C# compiler that ships with .NET Framework 4.x, then runs the tests.
rem Usage: build.cmd [version]   e.g. build.cmd 1.2.3 (default 0.1.0)
setlocal
cd /d "%~dp0"
set VERSION=%~1
if "%VERSION%"=="" set VERSION=0.1.0

set CSC=%WINDIR%\Microsoft.NET\FrameworkArm64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist bin mkdir bin

> bin\Version.cs echo [assembly: System.Reflection.AssemblyVersion("%VERSION%.0")]

"%CSC%" /nologo /target:library /platform:anycpu /optimize /out:bin\mailformat.dll /reference:Microsoft.CSharp.dll /reference:Microsoft.VisualBasic.dll Connect.cs Mailformat.cs bin\Version.cs || exit /b 1

"%CSC%" /nologo /target:exe /out:bin\tests.exe Tests.cs Mailformat.cs || exit /b 1

rem Integration test; needs the add-in installed (install.cmd), so it is built but not run here.
"%CSC%" /nologo /target:exe /platform:anycpu /out:bin\comtest.exe /reference:System.Xml.dll ComTest.cs || exit /b 1

bin\tests.exe
