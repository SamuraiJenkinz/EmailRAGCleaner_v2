# Test script to load modules and show actual errors
$ErrorActionPreference = "Continue"

Write-Host "Testing module loads..." -ForegroundColor Cyan

$modulePath = "C:\EmailRAGCleaner\Modules\RAGConfigManager_v2.psm1"

if (Test-Path $modulePath) {
    Write-Host "Found module at: $modulePath" -ForegroundColor Green
    
    # Try to parse the module
    try {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $modulePath,
            [ref]$tokens,
            [ref]$errors
        )
        
        if ($errors.Count -gt 0) {
            Write-Host "`nParse Errors Found:" -ForegroundColor Red
            foreach ($err in $errors) {
                Write-Host "  Line $($err.Extent.StartLineNumber): $($err.Message)" -ForegroundColor Yellow
                Write-Host "  Near: $($err.Extent.Text)" -ForegroundColor Gray
            }
        } else {
            Write-Host "No parse errors found!" -ForegroundColor Green
        }
    } catch {
        Write-Host "Failed to parse: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Module not found at expected location!" -ForegroundColor Red
}