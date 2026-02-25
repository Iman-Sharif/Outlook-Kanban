@echo off
setlocal EnableExtensions EnableDelayedExpansion

title Kanban for Outlook - Setup

set "APPNAME=Kanban for Outlook"
set "APPDIR=%USERPROFILE%\kanban-for-outlook"
set "SOURCEDIR=%~dp0"
set "HAS_CHOICE=0"

where choice >nul 2>&1
if not errorlevel 1 set "HAS_CHOICE=1"

:MENU
cls
echo ============================================================
echo %APPNAME% - Setup
echo Maintained by Iman Sharif
echo ============================================================
echo.
echo Disclaimer: provided "AS IS" (no warranty).
echo.
echo Install folder:
echo   %APPDIR%
echo.
echo This tool can install/upgrade, repair, or uninstall.
echo It writes registry keys under HKCU (current user).
echo.
echo  1^) Install / Upgrade (copy files + register home page)
echo  2^) Repair (register home page only)
echo  3^) Uninstall (unregister + remove installed files)
echo  4^) Open Start Here / Docs
echo  5^) Exit
echo.

call :PromptMenuChoice
if "%KFO_MENU%"=="1" (call :InstallUpgrade & goto :MENU)
if "%KFO_MENU%"=="2" (call :Repair & goto :MENU)
if "%KFO_MENU%"=="3" (call :Uninstall & goto :MENU)
if "%KFO_MENU%"=="4" (call :OpenDocs & goto :MENU)
goto :EOF

:PromptMenuChoice
set "KFO_MENU="
:PromptMenuChoiceLoop
if "%HAS_CHOICE%"=="1" (
  choice /c 12345 /m "Select an option"
  set "KFO_MENU=%errorlevel%"
) else (
  set /p "KFO_MENU=Select an option (1-5): "
  set "KFO_MENU=%KFO_MENU:~0,1%"
)

if "%KFO_MENU%"=="1" exit /b 0
if "%KFO_MENU%"=="2" exit /b 0
if "%KFO_MENU%"=="3" exit /b 0
if "%KFO_MENU%"=="4" exit /b 0
if "%KFO_MENU%"=="5" exit /b 0

echo.
echo Invalid selection. Please choose 1-5.
echo.
goto :PromptMenuChoiceLoop

:ConfirmYN
set "KFO_YN=0"
if "%HAS_CHOICE%"=="1" (
  choice /c YN /m "%~1"
  if errorlevel 2 (set "KFO_YN=0") else (set "KFO_YN=1")
) else (
  set /p "ans=%~1 (Y/N): "
  if /I "%ans%"=="Y" set "KFO_YN=1"
)
exit /b 0

:DetectOfficeVersion
set "KFO_OFFVER="
for /l %%x in (12,1,16) do (
  reg query "HKCU\Software\Microsoft\Office\%%x.0\Outlook\Today" >nul 2>&1 && set "KFO_OFFVER=%%x"
)
exit /b 0

:RegisterHomePage
call :DetectOfficeVersion
if "%KFO_OFFVER%"=="" (
  echo.
  echo Could not detect your Office version.
  echo You can still set the Folder Home Page manually to:
  echo   %~1
  echo.
  pause
  exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v Stamp /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v UserDefinedUrl /t REG_SZ /d "%~1" /f >nul
exit /b 0

:UnregisterHomePage
call :DetectOfficeVersion
if "%KFO_OFFVER%"=="" (
  exit /b 0
)
reg add "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v Stamp /t REG_DWORD /d 0 /f >nul
reg delete "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v UserDefinedUrl /f >nul 2>&1
exit /b 0

:CheckSourceFolder
if not exist "%SOURCEDIR%kanban.html" (
  echo.
  echo This script is running in the wrong folder.
  echo It must be located next to kanban.html.
  echo.
  pause
  exit /b 1
)
exit /b 0

:InstallUpgrade
cls
echo ============================================================
echo Install / Upgrade
echo ============================================================
echo.
echo Tip: close Outlook before installing/upgrading.
echo.

call :ConfirmYN "Continue with Install / Upgrade?"
if "%KFO_YN%"=="0" exit /b 0

call :CheckSourceFolder
if errorlevel 1 exit /b 0

if not exist "%APPDIR%" mkdir "%APPDIR%" >nul 2>&1

echo.
echo Copying files to:
echo   %APPDIR%

rem Copy files to a stable local folder (exclude repo/maintainer files if present)
robocopy /mir "%SOURCEDIR%" "%APPDIR%" /XD .git .github dist node_modules tools tests /XF "*.zip" "package.json" "package-lock.json" ".gitignore" >nul
set "RC=%errorlevel%"
if %RC% GEQ 8 (
  echo.
  echo Copy failed (robocopy errorlevel %RC%).
  pause
  exit /b 1
)

call :RegisterHomePage "%APPDIR%\kanban.html"

echo.
echo Installed to:
echo   %APPDIR%
echo.
echo Restart Outlook to load the board.
pause
exit /b 0

:Repair
cls
echo ============================================================
echo Repair (register home page)
echo ============================================================
echo.

call :ConfirmYN "Register the home page for this user profile?"
if "%KFO_YN%"=="0" exit /b 0

set "TARGET="
if exist "%APPDIR%\kanban.html" set "TARGET=%APPDIR%\kanban.html"
if "%TARGET%"=="" if exist "%SOURCEDIR%kanban.html" set "TARGET=%SOURCEDIR%kanban.html"

if "%TARGET%"=="" (
  echo.
  echo Could not find kanban.html.
  echo Run Install / Upgrade first.
  pause
  exit /b 1
)

call :RegisterHomePage "%TARGET%"
echo.
echo Registered:
echo   %TARGET%
echo.
echo Restart Outlook to load the board.
pause
exit /b 0

:Uninstall
cls
echo ============================================================
echo Uninstall
echo ============================================================
echo.

call :ConfirmYN "Uninstall Kanban for Outlook from this user profile?"
if "%KFO_YN%"=="0" exit /b 0

call :UnregisterHomePage

rem If running from inside the installed folder, move out first.
cd /d "%TEMP%" >nul 2>&1

if exist "%APPDIR%" (
  rmdir "%APPDIR%" /s /q
)

echo.
echo Uninstalled.
pause
exit /b 0

:OpenDocs
cls
echo ============================================================
echo Open Start Here / Docs
echo ============================================================
echo.

set "BASEDIR=%SOURCEDIR%"
if exist "%APPDIR%\START_HERE.html" set "BASEDIR=%APPDIR%\"

echo Opening:
echo   %BASEDIR%START_HERE.html
echo   %BASEDIR%docs\index.html
echo.

start "" "%BASEDIR%START_HERE.html" >nul 2>&1
start "" "%BASEDIR%docs\index.html" >nul 2>&1

echo If nothing opened, you can browse to the files above.
pause
exit /b 0
