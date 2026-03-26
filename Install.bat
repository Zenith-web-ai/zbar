@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0payload\zbar-install.ps1" -UsbDir "%~dp0"
pause
