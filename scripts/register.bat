@echo off
:: ====================================================================
:: Register SolidWorks Release Pack Add-In
:: Run this script as Administrator
:: ====================================================================

setlocal

set "DLL_PATH=%~dp0..\src\ReleasePack.AddIn\bin\x64\Release\net48\ReleasePack.AddIn.dll"

if not exist "%DLL_PATH%" (
    set "DLL_PATH=%~dp0..\src\ReleasePack.AddIn\bin\x64\Debug\net48\ReleasePack.AddIn.dll"
)

if not exist "%DLL_PATH%" (
    echo ERROR: Cannot find ReleasePack.AddIn.dll
    echo Build the solution first: dotnet build SolidWorksReleasePack.sln
    pause
    exit /b 1
)

echo Registering: %DLL_PATH%

:: Use the 64-bit .NET Framework RegAsm
set "REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if not exist "%REGASM%" (
    echo ERROR: RegAsm.exe not found at %REGASM%
    pause
    exit /b 1
)

"%REGASM%" /codebase "%DLL_PATH%"
if errorlevel 1 (
    echo.
    echo Registration FAILED. Make sure you are running as Administrator.
    pause
    exit /b 1
)

echo.
echo ✓ Registration successful. Restart SolidWorks to load the add-in.
pause
