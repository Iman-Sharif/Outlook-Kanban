@echo off
setlocal

title Kanban for Outlook - Uninstall

cls
echo ============================================================
echo Kanban for Outlook - Uninstall
echo Maintained by Iman Sharif
echo ============================================================
echo.
echo Disclaimer: provided "AS IS" (no warranty).
echo.

where choice >nul 2>&1
if errorlevel 1 (
  set /p KFO_CONTINUE="Uninstall Kanban for Outlook from this user profile? (Y/N): "
  if /I not "%KFO_CONTINUE%"=="Y" goto :EOF
) else (
  choice /c YN /m "Uninstall Kanban for Outlook from this user profile?"
  if errorlevel 2 goto :EOF
)

set "APPDIR=%USERPROFILE%\kanban-for-outlook"
if exist "%APPDIR%" (rmdir "%APPDIR%" /s /q)

set "offver="
for /l %%x in (12,1,16) do (
  reg query "HKCU\Software\Microsoft\Office\%%x.0\Outlook\Today" >nul 2>&1 && set "offver=%%x"
)

if "%offver%"=="" (
  echo Could not detect Office version; nothing to unregister.
  goto :EOF
)

reg add "HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today" /v Stamp /t REG_DWORD /d 0 /f >nul
reg delete "HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today" /v UserDefinedUrl /f >nul 2>&1

echo Uninstalled.
goto :EOF
