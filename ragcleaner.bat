@echo off
REM Email RAG Cleaner v2.0 - Quick Launch Batch File
REM Place this in a folder in your PATH for easy access

if "%1"=="gui" goto gui
if "%1"=="cli" goto cli
if "%1"=="test" goto test
if "%1"=="help" goto help

:gui
echo Starting Email RAG Cleaner GUI...
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\EmailRAGCleaner\Scripts\Start-EmailRAGCleanerGUI.ps1"
goto end

:cli
echo Starting Email RAG Cleaner CLI...
powershell -ExecutionPolicy Bypass -File "C:\EmailRAGCleaner\Start-EmailRAGCleaner.ps1" %2 %3 %4 %5
goto end

:test
echo Starting Email RAG Cleaner in Test Mode...
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\EmailRAGCleaner\Scripts\Start-EmailRAGCleanerGUI.ps1" -TestMode
goto end

:help
echo Email RAG Cleaner v2.0 - Quick Launch Commands
echo.
echo Usage: ragcleaner [option]
echo.
echo Options:
echo   gui    - Launch the modern GUI interface (default)
echo   cli    - Launch the command line interface
echo   test   - Launch GUI in test mode
echo   help   - Show this help message
echo.
echo Examples:
echo   ragcleaner          - Launches GUI
echo   ragcleaner gui      - Launches GUI
echo   ragcleaner cli      - Launches CLI
echo   ragcleaner test     - Launches GUI test mode
echo.
goto end

:end