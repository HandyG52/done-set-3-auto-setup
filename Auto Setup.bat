@echo off
cd /d "%~dp0"

REM Check for Administrator privileges to ensure script features work correctly
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~0' -Verb RunAs"
    exit /b
)

REM Download the latest version of the script from github
Powershell.exe -Command "try { Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/QuackWalks/done-set-3-auto-setup/refs/heads/main/done-auto.ps1' -OutFile '.\done-auto.ps1' -Force -ErrorAction Stop } catch { exit 1 }"
if %errorlevel% neq 0 (
    echo Failed to download the setup script. Please check your internet connection.
    pause
    exit /b
)

REM Unblock the file to prevent 'Mark of the Web' security warnings
Powershell.exe -Command "Unblock-File -Path '.\done-auto.ps1' -ErrorAction SilentlyContinue"

REM Run the script
Powershell.exe -ExecutionPolicy Bypass -File "./done-auto.ps1"
