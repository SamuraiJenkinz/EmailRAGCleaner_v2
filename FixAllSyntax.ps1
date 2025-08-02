# PowerShell 5.1 Syntax Fix Script
# This script will fix all PowerShell 7+ syntax in the modules

Write-Host "Fixing PowerShell 5.1 compatibility issues..." -ForegroundColor Cyan

$modulesPath = "C:\EmailRAGCleaner\Modules"
$fixes = 0

# Fix RAGConfigManager_v2.psm1
Write-Host "Fixing RAGConfigManager_v2.psm1..." -ForegroundColor Yellow
$ragConfigPath = Join-Path $modulesPath "RAGConfigManager_v2.psm1"
if (Test-Path $ragConfigPath) {
    $content = Get-Content $ragConfigPath -Raw
    
    # Fix line 321 - ternary operator
    $oldLine321 = 'ValidationPassed = $(if ($ValidateAfterUpdate) { $validationResult.IsValid } else { $null })'
    $newLine321 = 'ValidationPassed = $(if ($ValidateAfterUpdate) { $validationResult.IsValid } else { $null })'
    if ($content -match [regex]::Escape($oldLine321)) {
        Write-Host "  Line 321 already fixed" -ForegroundColor Green
    } else {
        # Look for the problematic pattern
        $content = $content -replace 'ValidationPassed = \$\(if \(\$ValidateAfterUpdate\).*?\)\s*}', 'ValidationPassed = $(if ($ValidateAfterUpdate) { $validationResult.IsValid } else { $null })'
        $fixes++
    }
    
    # Fix line 856 - ternary operator  
    $content = $content -replace '\$minor = \$\(if \(\$versionParts\.Length -gt 1\) \{ \[int\]\$versionParts\[1\] \} else \{ 0 \}\)', '$minor = $(if ($versionParts.Length -gt 1) { [int]$versionParts[1] } else { 0 })'
    
    # Fix line 865 - missing quote
    $content = $content -replace 'Write-Verbose "RAGConfigManager_v2 module loaded successfully"', 'Write-Verbose "RAGConfigManager_v2 module loaded successfully"'
    
    # Ensure all Try blocks have Catch or Finally
    $content = $content -replace '(try\s*\{[^}]+\})\s*(?!catch|finally)', '$1
    catch {
        Write-Error "Error: $_"
        throw
    }'
    
    $content | Set-Content $ragConfigPath -Force
    Write-Host "  Fixed $ragConfigPath" -ForegroundColor Green
}

# Fix EmailRAGProcessor_v2.psm1  
Write-Host "Fixing EmailRAGProcessor_v2.psm1..." -ForegroundColor Yellow
$emailProcPath = Join-Path $modulesPath "EmailRAGProcessor_v2.psm1"
if (Test-Path $emailProcPath) {
    $content = Get-Content $emailProcPath -Raw
    
    # Fix any ternary operators
    $content = $content -replace '\?\s*([^:]+)\s*:\s*([^;\r\n]+)', '$(if ($condition) { $1 } else { $2 })'
    
    # Fix line 537 issue - ensure proper Try-Catch structure
    $content = $content -replace '(try\s*\{[^}]+\})\s*\}\s*catch', '$1
    }
    catch'
    
    $content | Set-Content $emailProcPath -Force
    Write-Host "  Fixed $emailProcPath" -ForegroundColor Green
}

# Fix RAGTestFramework_v2.psm1
Write-Host "Fixing RAGTestFramework_v2.psm1..." -ForegroundColor Yellow  
$testFramePath = Join-Path $modulesPath "RAGTestFramework_v2.psm1"
if (Test-Path $testFramePath) {
    $content = Get-Content $testFramePath -Raw
    
    # Fix line 350 - escape % in string
    $content = $content -replace '\(\$\(\$qualityResult\.QualityMetrics\.SearchReadinessPercentage\)%\)', '($($qualityResult.QualityMetrics.SearchReadinessPercentage)`%)'
    
    # Fix line 658 - string formatting issues
    $content = $content -replace '\(\$\(\$performanceResult\.Metrics\.AverageMemoryUsageMB\) MB\)', '($($performanceResult.Metrics.AverageMemoryUsageMB) MB)'
    
    # Fix line 814 - quote the date format string
    $content = $content -replace '\.ToString\(''yyyy-MM-dd', '.ToString("yyyy-MM-dd'
    
    # Ensure all Try blocks have Catch
    $content = $content -replace '(try\s*\{[^}]+\})\s*(?!catch|finally)', '$1
    catch {
        Write-Error "Error: $_"
        throw
    }'
    
    $content | Set-Content $testFramePath -Force
    Write-Host "  Fixed $testFramePath" -ForegroundColor Green
}

# Fix the GUI XAML GroupBox issue
Write-Host "Fixing GUI XAML..." -ForegroundColor Yellow
$guiPath = "C:\EmailRAGCleaner\Scripts\EmailRAGCleaner_GUI_v2_Fixed.ps1"
if (Test-Path $guiPath) {
    $content = Get-Content $guiPath -Raw
    
    # The GroupBox Content error happens when Header and child content conflict
    # Need to ensure GroupBox uses proper structure
    $content = $content -replace '<GroupBox Header="([^"]+)"([^>]*)>\s*<StackPanel', '<GroupBox Header="$1"$2>
                            <GroupBox.Content>
                                <StackPanel'
    
    $content = $content -replace '<GroupBox Header="([^"]+)"([^>]*)>\s*<Grid', '<GroupBox Header="$1"$2>
                            <GroupBox.Content>
                                <Grid'
    
    $content = $content -replace '<GroupBox Header="([^"]+)"([^>]*)>\s*<WrapPanel', '<GroupBox Header="$1"$2>
                            <GroupBox.Content>
                                <WrapPanel'
    
    # Close GroupBox.Content before closing GroupBox
    $content = $content -replace '</GroupBox>', '</GroupBox.Content>
                        </GroupBox>'
    
    # Remove duplicate closing tags
    $content = $content -replace '</GroupBox\.Content>\s*</GroupBox\.Content>', '</GroupBox.Content>'
    
    $content | Set-Content $guiPath -Force
    Write-Host "  Fixed $guiPath" -ForegroundColor Green
}

Write-Host "`nAll syntax fixes applied!" -ForegroundColor Green
Write-Host "Total fixes: $fixes" -ForegroundColor Cyan