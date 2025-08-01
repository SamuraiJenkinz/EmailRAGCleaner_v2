@echo off
REM Email RAG Cleaner v2.0 - Direct GUI Launcher
REM This ensures PowerShell runs in STA mode for WPF

echo Starting Email RAG Cleaner v2.0 GUI...
echo.

REM Launch PowerShell in STA mode with the fixed GUI script
powershell.exe -sta -ExecutionPolicy Bypass -File "C:\EmailRAGCleaner\Scripts\EmailRAGCleaner_GUI_v2_Fixed.ps1"

if errorlevel 1 (
    echo.
    echo GUI failed to start. Check the error messages above.
    pause
)