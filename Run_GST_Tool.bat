@echo off
title GST Reconciliation Tool
color 1F
echo.
echo  =====================================================
echo   GST Reconciliation Tool - Free Edition
echo  =====================================================
echo.
echo  Starting tool, please wait...
echo.

:: Try python first, then py launcher
python "%~dp0gst_recon_tool.py" 2>nul
if errorlevel 1 (
    py "%~dp0gst_recon_tool.py" 2>nul
)
if errorlevel 1 (
    echo.
    echo  ERROR: Python not found.
    echo  Please install Python from https://www.python.org/downloads/
    echo  Make sure to check "Add Python to PATH" during install.
    echo.
    pause
)
