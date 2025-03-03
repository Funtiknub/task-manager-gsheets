@echo off
:: Check if the script is running with administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    :: If running as admin, change directory and run the Python script
    cd /d C:\Users\Fikusnat\Desktop\task-manager-gsheets
    py todo.py
) else (
    :: If not running as admin, relaunch the script with administrative privileges
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
)