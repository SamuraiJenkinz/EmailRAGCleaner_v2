# Copy clean working modules to installation directory
Write-Host "Copying clean modules from Enhanced_v2_Modules..." -ForegroundColor Cyan

$sourceDir = "C:\users\taylo\Downloads\EmailRAGCleaner_v2\Enhanced_v2_Modules"  
$targetDir = "C:\EmailRAGCleaner\Modules"

# Ensure target directory exists
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

# Copy all modules
$modules = @(
    "RAGConfigManager_v2.psm1",
    "EmailRAGProcessor_v2.psm1", 
    "EmailChunkingEngine_v2.psm1",
    "AzureAISearchIntegration_v2.psm1",
    "EmailSearchInterface_v2.psm1",
    "EmailEntityExtractor_v2.psm1",
    "RAGTestFramework_v2.psm1"
)

foreach ($module in $modules) {
    $sourcePath = Join-Path $sourceDir $module
    $targetPath = Join-Path $targetDir $module
    
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath $targetPath -Force
        Write-Host "  Copied $module" -ForegroundColor Green
    } else {
        Write-Host "  Missing: $module" -ForegroundColor Yellow
    }
}

Write-Host "Module copy completed!" -ForegroundColor Green