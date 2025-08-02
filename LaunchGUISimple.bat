@echo off
echo Starting Email RAG Cleaner v2.0 GUI (Simple Mode)...
echo.

REM Launch without trying to load modules first
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Add-Type -AssemblyName PresentationFramework; Write-Host 'Launching GUI...' -ForegroundColor Green; & 'C:\EmailRAGCleaner\Scripts\EmailRAGCleaner_GUI_v2_Fixed.ps1'}"

pause