@echo off
title Installing GST Tool Dependencies
echo.
echo  Installing required Python packages...
echo  (Internet connection needed - one time only)
echo.
pip install openpyxl xlsxwriter pandas pdfplumber
echo.
echo  =====================================================
echo   Installation Complete!
echo   Now double-click Run_GST_Tool.bat to start.
echo  =====================================================
echo.
pause
