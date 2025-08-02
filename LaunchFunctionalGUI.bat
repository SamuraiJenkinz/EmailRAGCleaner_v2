@echo off
echo ====================================================
echo Email RAG Cleaner v2.0 - FUNCTIONAL GUI Launcher
echo ====================================================
echo.
echo 🚀 Starting REAL email processing with PowerShell 7
echo 📧 MSG file processing with Azure AI Search integration
echo ⚡ Parallel processing and real-time progress tracking
echo.

REM Copy fixed modules to installation directory
echo Copying enhanced modules...
if not exist "C:\EmailRAGCleaner\Modules" mkdir "C:\EmailRAGCleaner\Modules"

copy /Y "Enhanced_v2_Modules\RAGConfigManager_v2_Fixed.psm1" "C:\EmailRAGCleaner\Modules\" >nul 2>&1
copy /Y "Enhanced_v2_Modules\EmailRAGProcessor_v2_Fixed.psm1" "C:\EmailRAGCleaner\Modules\" >nul 2>&1  
copy /Y "Enhanced_v2_Modules\RAGTestFramework_v2_Fixed.psm1" "C:\EmailRAGCleaner\Modules\" >nul 2>&1

REM Copy functional GUI script
if not exist "C:\EmailRAGCleaner\Scripts" mkdir "C:\EmailRAGCleaner\Scripts"
copy /Y "Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1" "C:\EmailRAGCleaner\Scripts\" >nul 2>&1

echo ✅ Enhanced modules and functional GUI copied
echo.

REM Launch with PowerShell 7 if available, otherwise PowerShell 5
where pwsh >nul 2>&1
if %ERRORLEVEL%==0 (
    echo 🔥 Launching with PowerShell 7 for maximum performance...
    pwsh.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1'"
) else (
    echo 📋 PowerShell 7 not found, using Windows PowerShell...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1'"
)

pause