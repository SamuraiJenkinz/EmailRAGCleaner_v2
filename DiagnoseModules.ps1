# Diagnostic script to find exact syntax errors
$ErrorActionPreference = "Continue"

Write-Host "Diagnosing module syntax errors..." -ForegroundColor Cyan

$modules = @(
    "RAGConfigManager_v2.psm1",
    "EmailRAGProcessor_v2.psm1", 
    "RAGTestFramework_v2.psm1"
)

foreach ($module in $modules) {
    $modulePath = "C:\EmailRAGCleaner\Modules\$module"
    
    if (Test-Path $modulePath) {
        Write-Host "`nChecking $module..." -ForegroundColor Yellow
        
        try {
            $tokens = $null
            $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $modulePath,
                [ref]$tokens,
                [ref]$errors
            )
            
            if ($errors.Count -gt 0) {
                Write-Host "  Parse Errors Found:" -ForegroundColor Red
                foreach ($err in $errors) {
                    Write-Host "    Line $($err.Extent.StartLineNumber): $($err.Message)" -ForegroundColor Red
                    Write-Host "    Near: $($err.Extent.Text)" -ForegroundColor Gray
                    Write-Host ""
                }
            } else {
                Write-Host "  No parse errors found!" -ForegroundColor Green
            }
        } catch {
            Write-Host "  Failed to parse: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  Module not found: $modulePath" -ForegroundColor Red
    }
}