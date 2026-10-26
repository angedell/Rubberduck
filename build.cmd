@echo off
setlocal
rem Double-click to restore and build Rubberduck.sln (Release, Any CPU), build the installer
rem and copy it to your deployment folder.
rem Pass "Debug" as the first argument for a debug build.

set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Preview\MSBuild\Current\Bin\MSBuild.exe"
set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=Release"

if not exist "%MSBUILD%" (
    echo MSBuild not found at: %MSBUILD%
    echo Edit the MSBUILD path at the top of this file.
    pause
    exit /b 1
)

cd /d "%~dp0"

rem Deployment folder: asked the first time, then remembered in .build-deploy-dir.
rem "-" in that file means "don't deploy". Delete the file to be asked again.
set "DEPLOY_DIR="
if exist "%~dp0.build-deploy-dir" set /p DEPLOY_DIR=<"%~dp0.build-deploy-dir"
if defined DEPLOY_DIR goto deploy_known

:ask_deploy
echo First run: choose the folder Rubberduck.Setup.exe is copied to after each build.
set "DEPLOY_DIR="
set /p "DEPLOY_DIR=Folder path (leave empty to skip deployment): "
if not defined DEPLOY_DIR set "DEPLOY_DIR=-"
set "DEPLOY_DIR=%DEPLOY_DIR:"=%"
if "%DEPLOY_DIR%"=="-" goto save_deploy
if not exist "%DEPLOY_DIR%\" echo Folder not found: %DEPLOY_DIR%& goto ask_deploy

:save_deploy
>"%~dp0.build-deploy-dir" echo %DEPLOY_DIR%
echo Saved to .build-deploy-dir
echo.

:deploy_known

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
if not "%INSTALLER_RC%"=="0" goto installer_failed

echo Self-contained installer ^(Inno Setup^):
echo   %~dp0Rubberduck.Deployment\InnoSetup\Installers\Rubberduck.Setup.exe
echo.
call :deploy
echo.
explorer /select,"%~dp0Rubberduck.Deployment\InnoSetup\Installers\Rubberduck.Setup.exe"
goto done

:installer_failed
echo INSTALLER BUILD FAILED - see messages above.
explorer "%~dp0Rubberduck.Deployment\bin"

:done
pause
exit /b %RC%

:deploy
if "%DEPLOY_DIR%"=="-" echo Deployment skipped ^(delete .build-deploy-dir to be asked for a folder again^).& exit /b 0
copy /y "%~dp0Rubberduck.Deployment\InnoSetup\Installers\Rubberduck.Setup.exe" "%DEPLOY_DIR%\Rubberduck.Setup.exe" >nul
if errorlevel 1 echo COPY FAILED: %DEPLOY_DIR%& exit /b 1
echo Copied to:
echo   %DEPLOY_DIR%\Rubberduck.Setup.exe
exit /b 0
