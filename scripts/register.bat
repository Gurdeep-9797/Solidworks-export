@echo off
:: ====================================================================
:: Register SolidWorks Release Pack Add-In
:: Run this script as Administrator
::
:: Per codestack.net: SolidWorks add-ins MUST be registered with
:: regasm /codebase — the VS "Register for COM Interop" option
:: does NOT use /codebase and won't work correctly.
:: ====================================================================

setlocal

set "DLL_PATH=%~dp0..\src\ReleasePack.AddIn\bin\Release\net48\ReleasePack.AddIn.dll"

if not exist "%DLL_PATH%" (
    set "DLL_PATH=%~dp0..\src\ReleasePack.AddIn\bin\Debug\net48\ReleasePack.AddIn.dll"
)

if not exist "%DLL_PATH%" (
    echo ERROR: Cannot find ReleasePack.AddIn.dll
    echo Build the solution first: dotnet build SolidWorksReleasePack.sln
    pause
    exit /b 1
)

echo Registering: %DLL_PATH%

:: Use the 64-bit .NET Framework 4 RegAsm (SolidWorks is 64-bit)
set "REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if not exist "%REGASM%" (
    echo ERROR: RegAsm.exe not found at %REGASM%
    echo Make sure .NET Framework 4.x is installed.
    pause
    exit /b 1
)

:: /codebase is REQUIRED — tells CLR where to find the assembly
:: without needing it in the GAC
"%REGASM%" /codebase "%DLL_PATH%"
if errorlevel 1 (
    echo.
    echo Registration FAILED. Make sure you are running as Administrator.
    pause
    exit /b 1
)

echo.
echo Registration successful. Restart SolidWorks to load the add-in.
echo The add-in will appear in Tools ^> Add-Ins.
pause
