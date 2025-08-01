@echo off
REM Email RAG Cleaner v2.0 - Quick Setup Script
REM Creates the installation structure without running the full installer

echo Email RAG Cleaner v2.0 - Quick Setup
echo ====================================
echo.
echo This will create the necessary directories and copy files for the GUI.
echo.

set SOURCE_DIR=%~dp0
set INSTALL_DIR=C:\EmailRAGCleaner

echo Creating installation directories...

REM Create main directories
mkdir "%INSTALL_DIR%" 2>nul
mkdir "%INSTALL_DIR%\Modules" 2>nul
mkdir "%INSTALL_DIR%\Scripts" 2>nul
mkdir "%INSTALL_DIR%\Config" 2>nul
mkdir "%INSTALL_DIR%\Logs" 2>nul
mkdir "%INSTALL_DIR%\Documentation" 2>nul

echo.
echo Copying modules...

REM Copy v1 modules (remove version suffix for compatibility)
echo Core v1 Modules:
copy /Y "%SOURCE_DIR%Core_v1_Modules\MsgProcessor_v1.psm1" "%INSTALL_DIR%\Modules\MsgProcessor.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] MsgProcessor.psm1) else (echo   [FAIL] MsgProcessor.psm1)

copy /Y "%SOURCE_DIR%Core_v1_Modules\ContentCleaner_v1.psm1" "%INSTALL_DIR%\Modules\ContentCleaner.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] ContentCleaner.psm1) else (echo   [FAIL] ContentCleaner.psm1)

copy /Y "%SOURCE_DIR%Core_v1_Modules\AzureFlattener_v1.psm1" "%INSTALL_DIR%\Modules\AzureFlattener.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] AzureFlattener.psm1) else (echo   [FAIL] AzureFlattener.psm1)

copy /Y "%SOURCE_DIR%Core_v1_Modules\ConfigManager_v1.psm1" "%INSTALL_DIR%\Modules\ConfigManager.psm1" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] ConfigManager.psm1) else (echo   [FAIL] ConfigManager.psm1)

echo.
echo Enhanced v2 Modules:
REM Copy v2 modules (keep version suffix)
copy /Y "%SOURCE_DIR%Enhanced_v2_Modules\*.psm1" "%INSTALL_DIR%\Modules\" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] All v2 modules copied) else (echo   [FAIL] Some v2 modules failed)

echo.
echo Copying scripts...

REM Copy scripts
copy /Y "%SOURCE_DIR%Scripts\EmailRAGCleaner_GUI_v2_Fixed.ps1" "%INSTALL_DIR%\Scripts\" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] GUI script) else (echo   [FAIL] GUI script)

copy /Y "%SOURCE_DIR%Scripts\Start-EmailRAGCleanerGUI.ps1" "%INSTALL_DIR%\Scripts\" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] GUI launcher) else (echo   [FAIL] GUI launcher)

copy /Y "%SOURCE_DIR%Scripts\EmailCleaner_Main_v1.ps1" "%INSTALL_DIR%\Scripts\" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] Main v1 script) else (echo   [FAIL] Main v1 script)

echo.
echo Copying configuration...

REM Copy configuration
copy /Y "%SOURCE_DIR%Configuration\AzureAISearchSchema_v2.json" "%INSTALL_DIR%\Config\" >nul 2>&1
if %errorlevel% equ 0 (echo   [OK] Azure AI Search schema) else (echo   [FAIL] Azure AI Search schema)

echo.
echo Creating launchers...

REM Create simple launcher in installation directory
echo @echo off > "%INSTALL_DIR%\LaunchGUI.bat"
echo powershell.exe -sta -ExecutionPolicy Bypass -File "%INSTALL_DIR%\Scripts\EmailRAGCleaner_GUI_v2_Fixed.ps1" >> "%INSTALL_DIR%\LaunchGUI.bat"

REM Create default config
echo { > "%INSTALL_DIR%\Config\default-config.json"
echo   "Metadata": { >> "%INSTALL_DIR%\Config\default-config.json"
echo     "Name": "Email RAG Cleaner v2.0", >> "%INSTALL_DIR%\Config\default-config.json"
echo     "Version": "2.0", >> "%INSTALL_DIR%\Config\default-config.json"
echo     "InstallPath": "%INSTALL_DIR%" >> "%INSTALL_DIR%\Config\default-config.json"
echo   } >> "%INSTALL_DIR%\Config\default-config.json"
echo } >> "%INSTALL_DIR%\Config\default-config.json"

echo.
echo ========================================
echo Quick Setup Complete!
echo ========================================
echo.
echo Installation created at: %INSTALL_DIR%
echo.
echo To launch the GUI, run:
echo   %INSTALL_DIR%\LaunchGUI.bat
echo.
echo Or from any command prompt:
echo   C:\EmailRAGCleaner\LaunchGUI.bat
echo.
pause