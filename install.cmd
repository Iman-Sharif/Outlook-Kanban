@echo off
setlocal EnableExtensions EnableDelayedExpansion

title Kanban for Outlook - Setup

set "APPNAME=Kanban for Outlook"
set "APPDIR=%USERPROFILE%\kanban-for-outlook"
set "SOURCEDIR=%~dp0"
set "HAS_CHOICE=0"
set "KFO_LOG=%TEMP%\kanban-for-outlook-install.log"

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
echo   "%APPDIR%"
echo.
echo Installer log:
echo   "%KFO_LOG%"
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
if "%KFO_MENU%"=="1" call :InstallUpgrade
if "%KFO_MENU%"=="2" call :Repair
if "%KFO_MENU%"=="3" call :Uninstall
if "%KFO_MENU%"=="4" call :OpenDocs
if "%KFO_MENU%"=="5" goto :EOF
goto :MENU

:PromptMenuChoice
set "KFO_MENU="
:PromptMenuChoiceLoop
if "%HAS_CHOICE%"=="1" (
  choice /c 12345 /m "Select an option"
  set "KFO_MENU=!errorlevel!"
) else (
  set /p "KFO_MENU=Select an option (1-5): "
  set "KFO_MENU=!KFO_MENU:~0,1!"
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
  set "ans="
  set /p "ans=%~1 (Y/N): "
  if /I "!ans!"=="Y" set "KFO_YN=1"
)
exit /b 0

:DetectOfficeVersion
set "KFO_OFFVER="
set "KFO_CURVER="
for /f "tokens=3" %%A in ('reg query "HKCR\Outlook.Application\CurVer" /ve 2^>nul ^| findstr /i "REG_SZ"') do (
  set "KFO_CURVER=%%A"
)

if not "%KFO_CURVER%"=="" (
  for /f "tokens=3 delims=." %%V in ("%KFO_CURVER%") do set "KFO_OFFVER=%%V"
)

rem Fallback: probe common Office versions for an existing Outlook hive.
if "%KFO_OFFVER%"=="" (
  for /l %%x in (12,1,16) do (
    reg query "HKCU\Software\Microsoft\Office\%%x.0\Outlook" >nul 2>&1 && set "KFO_OFFVER=%%x"
  )
)
exit /b 0

:RegisterHomePage
call :DetectOfficeVersion
if "%KFO_OFFVER%"=="" goto :RegisterHomePageNoOffice
if not "%KFO_LOG%"=="" echo RegisterHomePage: Office=%KFO_OFFVER% Path="%~1">> "%KFO_LOG%"

reg add "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v Stamp /t REG_DWORD /d 1 /f >nul 2>&1
set "R1=!errorlevel!"
if not "%KFO_LOG%"=="" echo reg add Stamp exit code: !R1!>> "%KFO_LOG%"
if not "!R1!"=="0" exit /b 1

reg add "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v UserDefinedUrl /t REG_SZ /d "%~1" /f >nul 2>&1
set "R2=!errorlevel!"
if not "%KFO_LOG%"=="" echo reg add UserDefinedUrl exit code: !R2!>> "%KFO_LOG%"
if not "!R2!"=="0" exit /b 1
exit /b 0

:RegisterHomePageNoOffice
if not "%KFO_LOG%"=="" echo RegisterHomePage: could not detect Office version>> "%KFO_LOG%"
echo.
echo Could not detect your Office version.
echo You can still set the Folder Home Page manually to:
echo   "%~1"
echo.
exit /b 1

:UnregisterHomePage
call :DetectOfficeVersion
if "%KFO_OFFVER%"=="" (
  exit /b 0
)
reg add "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v Stamp /t REG_DWORD /d 0 /f >nul
reg delete "HKCU\Software\Microsoft\Office\%KFO_OFFVER%.0\Outlook\Today" /v UserDefinedUrl /f >nul 2>&1
exit /b 0

:CheckSourceFolder
if not exist "%SOURCEDIR%kanban.html" goto :CheckSourceFolderWrong
rem Extra sanity checks (avoid installing from a partial/extracted script only)
if not exist "%SOURCEDIR%js\app\controller.js" goto :CheckSourceFolderMissing
exit /b 0

:CheckSourceFolderWrong
echo.
echo This script is running in the wrong folder.
echo It must be located next to kanban.html.
echo.
echo If you ran this from inside a zip preview, extract the zip first.
echo.
pause
exit /b 1

:CheckSourceFolderMissing
echo.
echo This folder is missing required app files (js\app\controller.js).
echo Please extract the full release zip and run install.cmd from inside it.
echo.
pause
exit /b 1

:InitLog
> "%KFO_LOG%" echo ============================================================
>> "%KFO_LOG%" echo %APPNAME% - install log
>> "%KFO_LOG%" echo ============================================================
>> "%KFO_LOG%" echo Time: %DATE% %TIME%
>> "%KFO_LOG%" echo Source: "%SOURCEDIR%."
>> "%KFO_LOG%" echo Target: "%APPDIR%"
>> "%KFO_LOG%" echo.
exit /b 0

:VerifyInstallFiles
if exist "%APPDIR%\kanban.html" if exist "%APPDIR%\js\app\controller.js" exit /b 0
exit /b 1

:CopyFiles
call :InitLog

echo Running file copy...>> "%KFO_LOG%"
echo.>> "%KFO_LOG%"

where robocopy >> "%KFO_LOG%" 2>&1
if errorlevel 1 goto :CopyFilesXcopy

rem Copy/overwrite files without purging extras (safer than /MIR).
rem NOTE: %SOURCEDIR% ends with a trailing backslash. Quoting a path that ends in \
rem can break argument parsing for some programs (the closing quote is escaped).
rem Using "%SOURCEDIR%." avoids a trailing backslash inside the quoted argument.
robocopy "%SOURCEDIR%." "%APPDIR%" /E /R:2 /W:1 /XD .git .github dist node_modules tools tests /XF "*.zip" "package.json" "package-lock.json" ".gitignore" >> "%KFO_LOG%" 2>&1
set "RC=!errorlevel!"
echo.>> "%KFO_LOG%"
echo robocopy exit code: !RC!>> "%KFO_LOG%"
if !RC! GEQ 8 exit /b 1

call :VerifyInstallFiles
if errorlevel 1 exit /b 1
exit /b 0

:CopyFilesXcopy
echo robocopy not available; falling back to xcopy>> "%KFO_LOG%"
echo.>> "%KFO_LOG%"

rem Avoid quoting a destination that ends with a backslash (can confuse arg parsing).
xcopy "%SOURCEDIR%*" "%APPDIR%" /E /I /H /K /Y >> "%KFO_LOG%" 2>&1
set "RC=!errorlevel!"
echo.>> "%KFO_LOG%"
echo xcopy exit code: !RC!>> "%KFO_LOG%"

call :VerifyInstallFiles
if errorlevel 1 exit /b 1
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
if errorlevel 1 goto :InstallMkdirFailed

rem Ensure a log exists even if a later step fails.
call :InitLog

echo.
echo Copying files to:
echo   "%APPDIR%"

call :CopyFiles
if errorlevel 1 goto :InstallCopyFailed

call :RegisterHomePage "%APPDIR%\kanban.html"
if errorlevel 1 goto :InstallRegFailed

echo.
echo Installed to:
echo   "%APPDIR%"
echo.
echo Restart Outlook to load the board.
pause
exit /b 0

:InstallMkdirFailed
echo.
echo Could not create install folder:
echo   "%APPDIR%"
echo.
echo This may be blocked by security policy or permissions.
pause
exit /b 1

:InstallCopyFailed
echo.
echo Install failed: files were not copied into the install folder.
echo.
echo Install folder exists but required files are missing.
echo Common causes:
echo   - Antivirus / Controlled folder access blocked the copy
echo   - The zip was not fully extracted
echo.
echo Log saved to:
echo   "%KFO_LOG%"
echo.
echo Try: extract the zip to a normal folder (e.g. Desktop) and re-run install.cmd.
pause
exit /b 1

:InstallRegFailed
echo.
echo Files were installed to:
echo   "%APPDIR%"
echo.
echo But automatic home page registration failed.
echo You can still set the Folder Home Page manually to:
echo   "%APPDIR%\kanban.html"
echo.
echo Log saved to:
echo   "%KFO_LOG%"
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

if "%TARGET%"=="" goto :RepairNoTarget

call :RegisterHomePage "%TARGET%"
if errorlevel 1 goto :RepairRegFailed
echo.
echo Registered:
echo   "%TARGET%"
echo.
echo Restart Outlook to load the board.
pause
exit /b 0

:RepairNoTarget
echo.
echo Could not find kanban.html.
echo Run Install / Upgrade first.
pause
exit /b 1

:RepairRegFailed
echo.
echo Registration failed.
echo You can still set the Folder Home Page manually to:
echo   "%TARGET%"
pause
exit /b 1

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

if exist "%APPDIR%" rmdir "%APPDIR%" /s /q

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
echo   "%BASEDIR%START_HERE.html"
echo   "%BASEDIR%docs\index.html"
echo.

start "" "%BASEDIR%START_HERE.html" >nul 2>&1
start "" "%BASEDIR%docs\index.html" >nul 2>&1

echo If nothing opened, you can browse to the files above.
pause
exit /b 0
