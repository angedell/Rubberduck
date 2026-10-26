@echo off
setlocal
rem Double-click to restore and build Rubberduck.sln (Release, Any CPU).
rem Pass "Debug" as the first argument for a debug build.

set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Preview\MSBuild\Current\Bin\MSBuild.exe"
set "DROPBOX_DIR=C:\Users\Angelo\SARO Dropbox\Angelo Dell Acqua\MacroCompactSources"
set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Release"

if not exist "%MSBUILD%" (
    echo MSBuild not found at: %MSBUILD%
    echo Edit the MSBUILD path at the top of this file.
    pause
    exit /b 1
)

cd /d "%~dp0"

rem Auto-bump the version: BASEVER.N, where N is a counter kept in .build-version.
rem The counter is only saved after a successful build, so failed builds don't burn numbers.
rem Every installer has a higher version than the previous one, so it upgrades the installed copy.
set "BASEVER=2.5.92"
set "BUILDNUM=0"
if exist "%~dp0.build-version" set /p BUILDNUM=<"%~dp0.build-version"
set /a BUILDNUM+=1
set "VERSION=%BASEVER%.%BUILDNUM%"
echo Building version %VERSION% [%CONFIG%]
echo.

"%MSBUILD%" Rubberduck.sln -t:Restore,Build -p:Configuration=%CONFIG% -p:Platform="Any CPU" -p:Version=%VERSION% -v:minimal -nologo -m
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
    echo BUILD FAILED, exit code %RC%
    pause
    exit /b %RC%
)

>"%~dp0.build-version" echo %BUILDNUM%
echo BUILD SUCCEEDED [%CONFIG%] version %VERSION%
echo.
echo Building installer...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-installer.ps1"
set "INSTALLER_RC=%ERRORLEVEL%"
echo.
echo Built add-in and all dependencies (what the installer packages):
echo   %~dp0Rubberduck.Deployment\bin\Rubberduck.dll
echo Add-in project output:
echo   %~dp0Rubberduck.Main\bin\%CONFIG%\net462\Rubberduck.dll
echo Test assemblies:
echo   %~dp0RubberduckTests\bin\RubberduckTests.dll
echo.
if "%INSTALLER_RC%"=="0" (
    echo Self-contained installer ^(Inno Setup^):
    echo   %~dp0Rubberduck.Deployment\InnoSetup\Installers\Rubberduck.Setup.exe
    echo.
    copy /y "%~dp0Rubberduck.Deployment\InnoSetup\Installers\Rubberduck.Setup.exe" "%DROPBOX_DIR%\Rubberduck.Setup.exe" >nul
    if errorlevel 1 (
        echo COPY TO DROPBOX FAILED: %DROPBOX_DIR%
    ) else (
        echo Copied to Dropbox:
        echo   %DROPBOX_DIR%\Rubberduck.Setup.exe
    )
    echo.
    explorer /select,"%~dp0Rubberduck.Deployment\InnoSetup\Installers\Rubberduck.Setup.exe"
) else (
    echo INSTALLER BUILD FAILED - see messages above.
    explorer "%~dp0Rubberduck.Deployment\bin"
)
pause
exit /b %RC%
