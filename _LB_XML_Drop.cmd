@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "PS1=%SCRIPT_DIR%\_LB_XML_Cleaner.ps1"
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%POWERSHELL_EXE%" (
  echo powershell.exe not found: "%POWERSHELL_EXE%"
  echo.
  pause
  exit /b 1
)

if not exist "%PS1%" (
  echo Script not found: "%PS1%"
  echo.
  pause
  exit /b 1
)

if "%~1"=="" (
  "%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -RomsPath "%SCRIPT_DIR%"
) else (
  "%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%~1" -RomsPath "%SCRIPT_DIR%"
)
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo Finished successfully.
) else (
  echo Finished with error code %EXITCODE%.
)
echo.
pause
exit /b %EXITCODE%
