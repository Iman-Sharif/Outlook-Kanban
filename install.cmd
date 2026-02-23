@echo off
setlocal

title Kanban for Outlook - Installer

cls
echo ============================================================
echo Kanban for Outlook - Installer
echo Maintained by Iman Sharif
echo ============================================================
echo.
echo Disclaimer: provided "AS IS" (no warranty).
echo.
echo This installer will:
echo - Copy files to a local folder under your user profile
echo - Set Outlook Folder Home Page registry keys under HKCU
echo.
echo Tip: If Windows SmartScreen shows "Windows protected your PC",
echo click "More info" then "Run anyway".
echo.

where choice >nul 2>&1
if errorlevel 1 (
  set /p KFO_CONTINUE="Continue with local installation? (Y/N): "
  if /I not "%KFO_CONTINUE%"=="Y" goto :EOF
) else (
  choice /c YN /m "Continue with local installation?"
  if errorlevel 2 goto :EOF
)

set "APPDIR=%USERPROFILE%\kanban-for-outlook"

if not exist "kanban.html" (
  echo The install script is running in the wrong folder.
  echo Please run it from the folder that contains kanban.html
  pause
  goto :EOF
)

if not exist "%APPDIR%" mkdir "%APPDIR%"

rem Copy files to a stable local folder (exclude repo/maintainer files if present)
robocopy /mir . "%APPDIR%" /XD .git .github dist node_modules tools tests /XF "*.zip" "package.json" "package-lock.json" ".gitignore" >nul

set "offver="
for /l %%x in (12,1,16) do (
  reg query "HKCU\Software\Microsoft\Office\%%x.0\Outlook\Today" >nul 2>&1 && set "offver=%%x"
)

if "%offver%"=="" (
  echo The install script could not detect your Office version.
  echo You can still set the Folder Home Page manually to:
  echo   %APPDIR%\kanban.html
  pause
  goto :EOF
)

reg add "HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today" /v Stamp /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today" /v UserDefinedUrl /t REG_SZ /d "%APPDIR%\kanban.html" /f >nul

cls
echo Kanban for Outlook successfully set up.
echo Restart Outlook to load the board.
pause
goto :EOF
