# Fix EmailRAGProcessor_v2.psm1 syntax errors
Write-Host "Fixing EmailRAGProcessor_v2.psm1..." -ForegroundColor Yellow

$processorPath = "C:\EmailRAGCleaner\Modules\EmailRAGProcessor_v2.psm1"

if (Test-Path $processorPath) {
    $content = Get-Content $processorPath -Raw
    
    # The issue is with the Test-RAGPipeline function structure
    # There are nested try-catch blocks that are causing parsing issues
    
    # Fix the function by ensuring proper bracket alignment
    # The problem is that there's a return statement followed by a stray catch block
    
    # Replace the problematic section with a corrected version
    $pattern = 'Write-Host "RAG pipeline test completed successfully!" -ForegroundColor Green\s+return @\{[^}]+\}\s+\} catch \{'
    
    $replacement = @'
Write-Host "RAG pipeline test completed successfully!" -ForegroundColor Green
        
        return @{
            Status = "Success"
            ConnectionTest = $connectionTest
            IndexStats = $(if ($indexStats) { $indexStats } else { $null })
            SearchTest = $(if ($searchResult) { $searchResult.Results.Count } else { 0 })
            HybridSearchTest = $(if ($hybridResult) { $hybridResult.Results.Count } else { 0 })
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
'@
    
    if ($content -match $pattern) {
        $content = $content -replace $pattern, $replacement
        Write-Host "  Fixed try-catch structure" -ForegroundColor Green
    } else {
        # Alternative approach - fix the specific lines
        # Remove the extra closing brace that's causing the issue
        $content = $content -replace '\s+\}\s+\} catch \{', '
        
    } catch {'
        Write-Host "  Fixed bracket alignment" -ForegroundColor Green
    }
    
    $content | Set-Content $processorPath -Force
    Write-Host "  Fixed EmailRAGProcessor_v2.psm1" -ForegroundColor Green
}

Write-Host "EmailRAGProcessor fixes applied!" -ForegroundColor Green