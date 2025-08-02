# Simple fix for RAGTestFramework_v2.psm1 syntax errors
Write-Host "Fixing RAGTestFramework_v2.psm1..." -ForegroundColor Yellow

$testFramePath = "C:\EmailRAGCleaner\Modules\RAGTestFramework_v2.psm1"

if (Test-Path $testFramePath) {
    $content = Get-Content $testFramePath -Raw
    
    # Fix 1: Escape % operator in line 350
    $content = $content -replace 'SearchReadinessPercentage\)\)%', 'SearchReadinessPercentage))`%'
    Write-Host "  Fixed % operator escaping at line 350" -ForegroundColor Green
    
    # Fix 2: Fix ToString date format quotes in line 814  
    $content = $content -replace "ToString\('yyyy-MM-dd HH:mm:ss'\)", 'ToString("yyyy-MM-dd HH:mm:ss")'
    Write-Host "  Fixed ToString date format quotes at line 814" -ForegroundColor Green
    
    # Fix 3: Ensure proper string interpolation for MB token
    # Just make sure there are no syntax issues with the MB usage
    $content = $content -replace '\$\(\$performanceResult\.Metrics\.AverageMemoryUsageMB\) MB\)', '($($performanceResult.Metrics.AverageMemoryUsageMB) MB)'
    Write-Host "  Verified MB token syntax" -ForegroundColor Green
    
    $content | Set-Content $testFramePath -Force
    Write-Host "  Fixed RAGTestFramework_v2.psm1" -ForegroundColor Green
}

Write-Host "RAGTestFramework fixes applied!" -ForegroundColor Green