@echo off
setlocal
set "ROOT=%~dp0"
set "ENTRY=%ROOT%Schuman-Main.ps1"

if not exist "%ENTRY%" (
  echo Schuman-Main.ps1 not found: "%ENTRY%"
  exit /b 1
)

start "" powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%ENTRY%"
exit /b 0
