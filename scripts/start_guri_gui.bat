@echo off
REM GURI GUI Startup Script
REM This script starts the GURI GUI application (suite root = parent of scripts\)

cd /d "%~dp0.."
python guri\gui.py

if errorlevel 1 (
    echo Error starting GURI GUI
    pause
)

