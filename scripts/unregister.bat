@echo off
:: ====================================================================
:: Unregister SolidWorks Release Pack Add-In
:: Run this script as Administrator
:: ====================================================================

setlocal

set "DLL_PATH=%~dp0..\src\ReleasePack.AddIn\bin\Release\net48\ReleasePack.AddIn.dll"

if not exist "%DLL_PATH%" (
    set "DLL_PATH=%~dp0..\src\ReleasePack.AddIn\bin\Debug\net48\ReleasePack.AddIn.dll"
)

if not exist "%DLL_PATH%" (
    echo ERROR: Cannot find ReleasePack.AddIn.dll
    pause
    exit /b 1
)

echo Unregistering: %DLL_PATH%

set "REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if not exist "%REGASM%" (
    echo ERROR: RegAsm.exe not found at %REGASM%
    pause
    exit /b 1
)

"%REGASM%" /unregister "%DLL_PATH%"
if errorlevel 1 (
    echo.
    echo Unregistration FAILED. Make sure you are running as Administrator.
    pause
    exit /b 1
)

echo.
echo Unregistration successful.
pause
