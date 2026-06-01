@echo off
chcp 65001 >nul
cd /d "%~dp0"
"%~dp0自动化报销匹配工作台\ReimbursementMatcher.exe"
