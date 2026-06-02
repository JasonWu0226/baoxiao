@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_invoice_audit.ps1" %*
echo.
pause
