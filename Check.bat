@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0payload\zbar-check.ps1" -UsbDir "%~dp0"
