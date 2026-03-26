@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0payload\zbar-uninstall.ps1" -UsbDir "%~dp0"
pause
