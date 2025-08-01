@echo off
REM Email RAG Cleaner v2.0 - Module Copy Script
REM Copies modules from source to installation directory

echo Copying Email RAG Cleaner v2.0 modules...
echo.

set SOURCE_DIR=C:\users\taylo\downloads\EmailRAGCleaner_v2
set DEST_DIR=C:\EmailRAGCleaner\Modules

REM Create destination directory if it doesn't exist
if not exist "%DEST_DIR%" (
    echo Creating Modules directory...
    mkdir "%DEST_DIR%"
)

echo Copying Enhanced v2 Modules...

REM Copy v2 modules
copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\RAGConfigManager_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] RAGConfigManager_v2.psm1) else (echo [FAIL] RAGConfigManager_v2.psm1)

copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\EmailRAGProcessor_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] EmailRAGProcessor_v2.psm1) else (echo [FAIL] EmailRAGProcessor_v2.psm1)

copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\EmailSearchInterface_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] EmailSearchInterface_v2.psm1) else (echo [FAIL] EmailSearchInterface_v2.psm1)

copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\RAGTestFramework_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] RAGTestFramework_v2.psm1) else (echo [FAIL] RAGTestFramework_v2.psm1)

copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\EmailChunkingEngine_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] EmailChunkingEngine_v2.psm1) else (echo [FAIL] EmailChunkingEngine_v2.psm1)

copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\AzureAISearchIntegration_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] AzureAISearchIntegration_v2.psm1) else (echo [FAIL] AzureAISearchIntegration_v2.psm1)

copy /Y "%SOURCE_DIR%\Enhanced_v2_Modules\EmailEntityExtractor_v2.psm1" "%DEST_DIR%\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] EmailEntityExtractor_v2.psm1) else (echo [FAIL] EmailEntityExtractor_v2.psm1)

echo.
echo Copying Core v1 Modules...

REM Copy v1 modules (without version suffix for compatibility)
copy /Y "%SOURCE_DIR%\Core_v1_Modules\MsgProcessor_v1.psm1" "%DEST_DIR%\MsgProcessor.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] MsgProcessor.psm1) else (echo [FAIL] MsgProcessor.psm1)

copy /Y "%SOURCE_DIR%\Core_v1_Modules\ContentCleaner_v1.psm1" "%DEST_DIR%\ContentCleaner.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] ContentCleaner.psm1) else (echo [FAIL] ContentCleaner.psm1)

copy /Y "%SOURCE_DIR%\Core_v1_Modules\AzureFlattener_v1.psm1" "%DEST_DIR%\AzureFlattener.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] AzureFlattener.psm1) else (echo [FAIL] AzureFlattener.psm1)

copy /Y "%SOURCE_DIR%\Core_v1_Modules\ConfigManager_v1.psm1" "%DEST_DIR%\ConfigManager.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] ConfigManager.psm1) else (echo [FAIL] ConfigManager.psm1)

echo.
echo Copying GUI Scripts...

REM Copy GUI scripts
if not exist "C:\EmailRAGCleaner\Scripts" mkdir "C:\EmailRAGCleaner\Scripts"

copy /Y "%SOURCE_DIR%\Scripts\EmailRAGCleaner_GUI_v2_Fixed.ps1" "C:\EmailRAGCleaner\Scripts\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] EmailRAGCleaner_GUI_v2_Fixed.ps1) else (echo [FAIL] EmailRAGCleaner_GUI_v2_Fixed.ps1)

copy /Y "%SOURCE_DIR%\Scripts\Start-EmailRAGCleanerGUI.ps1" "C:\EmailRAGCleaner\Scripts\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] Start-EmailRAGCleanerGUI.ps1) else (echo [FAIL] Start-EmailRAGCleanerGUI.ps1)

echo.
echo Copying Configuration...

REM Copy configuration schema
if not exist "C:\EmailRAGCleaner\Config" mkdir "C:\EmailRAGCleaner\Config"

copy /Y "%SOURCE_DIR%\Configuration\AzureAISearchSchema_v2.json" "C:\EmailRAGCleaner\Config\" >nul 2>&1
if %errorlevel% equ 0 (echo [OK] AzureAISearchSchema_v2.json) else (echo [FAIL] AzureAISearchSchema_v2.json)

echo.
echo Module copy complete!
echo.
pause