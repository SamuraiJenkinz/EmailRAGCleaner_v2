# Install Functional Email RAG Cleaner v2.0
Write-Host "Installing FUNCTIONAL Email RAG Cleaner v2.0..." -ForegroundColor Green

$installPath = "C:\EmailRAGCleaner"
$sourceDir = "."

# Create installation directories
@("$installPath", "$installPath\Modules", "$installPath\Scripts", "$installPath\Logs") | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        Write-Host "Created directory: $_" -ForegroundColor Green
    }
}

# Copy FIXED modules
Write-Host "Installing enhanced modules..." -ForegroundColor Cyan
$modules = @(
    "RAGConfigManager_v2_Fixed.psm1",
    "EmailRAGProcessor_v2_Fixed.psm1", 
    "RAGTestFramework_v2_Fixed.psm1"
)

foreach ($module in $modules) {
    $sourcePath = Join-Path "Enhanced_v2_Modules" $module
    $targetPath = Join-Path "$installPath\Modules" $module
    
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath $targetPath -Force
        Write-Host "  Installed: $module" -ForegroundColor Green
    } else {
        Write-Host "  Not found: $module" -ForegroundColor Red
    }
}

# Copy functional GUI
Write-Host "Installing functional GUI..." -ForegroundColor Cyan
$guiSource = "Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1"
$guiTarget = "$installPath\Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1"

if (Test-Path $guiSource) {
    Copy-Item $guiSource $guiTarget -Force
    Write-Host "  Installed: Functional GUI" -ForegroundColor Green
} else {
    Write-Host "  Not found: Functional GUI script" -ForegroundColor Red
}

# Copy legacy modules for compatibility
Write-Host "Installing compatibility modules..." -ForegroundColor Cyan
$legacyModules = @(
    "EmailSearchInterface_v2.psm1",
    "EmailChunkingEngine_v2.psm1",
    "AzureAISearchIntegration_v2.psm1",
    "EmailEntityExtractor_v2.psm1"
)

foreach ($module in $legacyModules) {
    $sourcePath = Join-Path "Enhanced_v2_Modules" $module
    $targetPath = Join-Path "$installPath\Modules" $module
    
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath $targetPath -Force
        Write-Host "  Installed: $module" -ForegroundColor Green
    }
}

# Create launcher script
Write-Host "Creating launcher..." -ForegroundColor Cyan
$launcherContent = @"
@echo off
echo ====================================================
echo Email RAG Cleaner v2.0 - FUNCTIONAL EDITION
echo ====================================================
echo.
echo REAL email processing with PowerShell 7 features
echo MSG files to Azure AI Search RAG pipeline  
echo Parallel processing with real-time progress
echo.

cd /d "$installPath"

REM Try PowerShell 7 first, fallback to Windows PowerShell
where pwsh >nul 2>&1
if %ERRORLEVEL%==0 (
    echo Launching with PowerShell 7...
    pwsh.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1'"
) else (
    echo Using Windows PowerShell...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'Scripts\EmailRAGCleaner_GUI_v2_Functional.ps1'"
)

pause
"@

$launcherPath = "$installPath\LaunchFunctionalGUI.bat"
$launcherContent | Out-File -FilePath $launcherPath -Encoding ASCII
Write-Host "  Created: $launcherPath" -ForegroundColor Green

Write-Host "`nFUNCTIONAL Email RAG Cleaner v2.0 installed successfully!" -ForegroundColor Green
Write-Host "Installation: $installPath" -ForegroundColor Cyan
Write-Host "Launcher: $launcherPath" -ForegroundColor Cyan
Write-Host "`nTo launch the functional GUI:" -ForegroundColor Yellow
Write-Host "  $launcherPath" -ForegroundColor White