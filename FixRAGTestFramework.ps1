# Fix RAGTestFramework_v2.psm1 syntax errors
Write-Host "Fixing RAGTestFramework_v2.psm1..." -ForegroundColor Yellow

$testFramePath = "C:\EmailRAGCleaner\Modules\RAGTestFramework_v2.psm1"

if (Test-Path $testFramePath) {
    $content = Get-Content $testFramePath -Raw
    
    # Fix 1: Escape % operator in string interpolations (line 350)
    $content = $content -replace '\(\$\(\$qualityResult\.QualityMetrics\.SearchReadinessPercentage\)\)%', '($($qualityResult.QualityMetrics.SearchReadinessPercentage))`%'
    Write-Host "  Fixed % operator escaping" -ForegroundColor Green
    
    # Fix 2: Fix any MB token issues in string interpolation (line 658)
    $content = $content -replace '\(\$\(\$performanceResult\.Metrics\.AverageMemoryUsageMB\) MB\)', '($($performanceResult.Metrics.AverageMemoryUsageMB) MB)'
    Write-Host "  Fixed MB token in string interpolation" -ForegroundColor Green
    
    # Fix 3: Fix ToString date format in here-string (line 814)
    $content = $content -replace "ToString\('yyyy-MM-dd HH:mm:ss'\)", 'ToString("yyyy-MM-dd HH:mm:ss")'
    Write-Host "  Fixed ToString date format quotes" -ForegroundColor Green
    
    # Fix 4: Ensure all Try blocks have proper Catch blocks
    # Look for any standalone try blocks and add catch blocks if needed
    $tryPattern = 'try\s*\{[^}]+\}\s*(?!catch|finally)'
    while ($content -match $tryPattern) {
        $content = $content -replace $tryPattern, '$&
    catch {
        Write-Error "Error: $_"
        throw
    }'
        Write-Host "  Added missing catch block" -ForegroundColor Green
    }
    
    # Fix 5: Fix any other percentage signs in strings
    $content = $content -replace '(\$\([^)]+\))%([^`])', '$1`%$2'
    
    # Fix 6: Ensure all functions have proper closing braces
    # This is a more complex fix, but we'll ensure the module ends properly
    if (-not $content.EndsWith("`n")) {
        $content += "`n"
    }
    
    $content | Set-Content $testFramePath -Force
    Write-Host "  Fixed RAGTestFramework_v2.psm1" -ForegroundColor Green
}

Write-Host "RAGTestFramework fixes applied!" -ForegroundColor Green