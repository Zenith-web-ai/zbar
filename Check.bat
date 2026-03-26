@echo off
:: Request admin so we can see SYSTEM tasks and processes
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process powershell.exe -Verb RunAs -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0payload\zbar-check.ps1\" -UsbDir \"%~dp0\"'"
    exit /b
)
powershell -ExecutionPolicy Bypass -File "%~dp0payload\zbar-check.ps1" -UsbDir "%~dp0"
