@echo off
REM GURI GUI Startup Script
REM This script starts the GURI GUI application

cd /d "C:\GeoFooter"
python guri_gui.py

if errorlevel 1 (
    echo Error starting GURI GUI
    pause
)

