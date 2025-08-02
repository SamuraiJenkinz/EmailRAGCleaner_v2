# Comprehensive PowerShell 5.1 Syntax Error Fix Script
Write-Host "Fixing all PowerShell syntax errors..." -ForegroundColor Cyan

$modulesPath = "C:\EmailRAGCleaner\Modules"

# Fix RAGConfigManager_v2.psm1
Write-Host "`nFixing RAGConfigManager_v2.psm1..." -ForegroundColor Yellow
$ragConfigPath = Join-Path $modulesPath "RAGConfigManager_v2.psm1"

if (Test-Path $ragConfigPath) {
    $content = Get-Content $ragConfigPath -Raw
    
    # The issue seems to be with missing closing braces. Let me reconstruct the problematic sections
    # Based on the parse errors, I need to ensure proper structure
    
    # Create a corrected version by replacing the problematic Test-RAGConfiguration function
    $correctedFunction = @'
function Test-RAGConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$false)]
        [switch]$TestConnections = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$TestPermissions = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$ValidateSettings = $true
    )
    
    try {
        Write-Host "Testing RAG configuration..." -ForegroundColor Yellow
        
        $testResults = @{
            OverallStatus = "Unknown"
            TestsRun = 0
            TestsPassed = 0
            TestsFailed = 0
            Details = @{}
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
        # Test 1: Configuration schema validation
        if ($ValidateSettings) {
            Write-Host "1. Validating configuration schema..." -ForegroundColor Cyan
            $testResults.TestsRun++
            
            try {
                $validationResult = Validate-ConfigurationSchema -Configuration $Configuration
                if ($validationResult.IsValid) {
                    Write-Host "   ✓ Configuration schema is valid" -ForegroundColor Green
                    $testResults.TestsPassed++
                    $testResults.Details.SchemaValidation = @{ Status = "Passed"; Message = "Schema validation successful" }
                } else {
                    Write-Host "   ✗ Configuration schema validation failed: $($validationResult.Errors -join ', ')" -ForegroundColor Red
                    $testResults.TestsFailed++
                    $testResults.Details.SchemaValidation = @{ Status = "Failed"; Errors = $validationResult.Errors }
                }
            } catch {
                Write-Host "   ✗ Schema validation error: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.TestsFailed++
                $testResults.Details.SchemaValidation = @{ Status = "Error"; Message = $_.Exception.Message }
            }
        }
        
        # Test 2: Azure Search connection
        if ($TestConnections -and $Configuration.AzureSearch) {
            Write-Host "2. Testing Azure Search connection..." -ForegroundColor Cyan
            $testResults.TestsRun++
            
            try {
                $searchConfig = @{
                    ServiceName = $Configuration.AzureSearch.ServiceName
                    ServiceUrl = $Configuration.AzureSearch.ServiceUrl
                    ApiKey = $Configuration.AzureSearch.ApiKey
                    ApiVersion = $Configuration.AzureSearch.ApiVersion
                    Headers = @{
                        'api-key' = $Configuration.AzureSearch.ApiKey
                        'Content-Type' = 'application/json'
                    }
                }
                
                $testUrl = "$($searchConfig.ServiceUrl)/servicestats?api-version=$($searchConfig.ApiVersion)"
                $response = Invoke-RestMethod -Uri $testUrl -Headers $searchConfig.Headers -Method GET -TimeoutSec 30
                
                Write-Host "   ✓ Azure Search connection successful" -ForegroundColor Green
                $testResults.TestsPassed++
                $testResults.Details.AzureSearchConnection = @{ 
                    Status = "Passed" 
                    ServiceName = $Configuration.AzureSearch.ServiceName
                    Response = $response
                }
            } catch {
                Write-Host "   ✗ Azure Search connection failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.TestsFailed++
                $testResults.Details.AzureSearchConnection = @{ Status = "Failed"; Error = $_.Exception.Message }
            }
        }
        
        # Test 3: OpenAI connection (if configured)
        if ($TestConnections -and $Configuration.OpenAI -and $Configuration.OpenAI.Enabled -and $Configuration.OpenAI.ApiKey) {
            Write-Host "3. Testing OpenAI connection..." -ForegroundColor Cyan
            $testResults.TestsRun++
            
            try {
                $openAIConfig = $Configuration.OpenAI
                $testUrl = "$($openAIConfig.Endpoint)/openai/deployments?api-version=$($openAIConfig.ApiVersion)"
                $headers = @{
                    'api-key' = $openAIConfig.ApiKey
                    'Content-Type' = 'application/json'
                }
                
                $response = Invoke-RestMethod -Uri $testUrl -Headers $headers -Method GET -TimeoutSec 30
                
                Write-Host "   ✓ OpenAI connection successful" -ForegroundColor Green
                $testResults.TestsPassed++
                $testResults.Details.OpenAIConnection = @{ 
                    Status = "Passed"
                    Endpoint = $openAIConfig.Endpoint
                    AvailableModels = $response.data.Count
                }
            } catch {
                Write-Host "   ✗ OpenAI connection failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.TestsFailed++
                $testResults.Details.OpenAIConnection = @{ Status = "Failed"; Error = $_.Exception.Message }
            }
        }
        
        # Test 4: Index existence check
        if ($TestPermissions -and $Configuration.AzureSearch) {
            Write-Host "4. Checking index permissions..." -ForegroundColor Cyan
            $testResults.TestsRun++
            
            try {
                $searchConfig = @{
                    ServiceUrl = $Configuration.AzureSearch.ServiceUrl
                    ApiKey = $Configuration.AzureSearch.ApiKey
                    ApiVersion = $Configuration.AzureSearch.ApiVersion
                    Headers = @{
                        'api-key' = $Configuration.AzureSearch.ApiKey
                        'Content-Type' = 'application/json'
                    }
                }
                
                $indexUrl = "$($searchConfig.ServiceUrl)/indexes?api-version=$($searchConfig.ApiVersion)"
                $response = Invoke-RestMethod -Uri $indexUrl -Headers $searchConfig.Headers -Method GET -TimeoutSec 30
                
                $targetIndex = $response.value | Where-Object { $_.name -eq $Configuration.AzureSearch.IndexName }
                
                if ($targetIndex) {
                    Write-Host "   ✓ Target index '$($Configuration.AzureSearch.IndexName)' exists" -ForegroundColor Green
                    $testResults.TestsPassed++
                    $testResults.Details.IndexPermissions = @{ 
                        Status = "Passed"
                        IndexExists = $true
                        IndexName = $Configuration.AzureSearch.IndexName
                    }
                } else {
                    Write-Host "   ⚠ Target index '$($Configuration.AzureSearch.IndexName)' does not exist (will be created)" -ForegroundColor Yellow
                    $testResults.TestsPassed++
                    $testResults.Details.IndexPermissions = @{ 
                        Status = "Warning"
                        IndexExists = $false
                        IndexName = $Configuration.AzureSearch.IndexName
                        Message = "Index will be created during processing"
                    }
                }
            } catch {
                Write-Host "   ✗ Index permissions check failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.TestsFailed++
                $testResults.Details.IndexPermissions = @{ Status = "Failed"; Error = $_.Exception.Message }
            }
        }
        
        # Test 5: Processing settings validation
        Write-Host "5. Validating processing settings..." -ForegroundColor Cyan
        $testResults.TestsRun++
        
        try {
            $processingValidation = Test-ProcessingSettings -Settings $Configuration.Processing
            if ($processingValidation.IsValid) {
                Write-Host "   ✓ Processing settings are valid" -ForegroundColor Green
                $testResults.TestsPassed++
                $testResults.Details.ProcessingSettings = @{ Status = "Passed"; Validation = $processingValidation }
            } else {
                Write-Host "   ⚠ Processing settings have warnings: $($processingValidation.Warnings -join ', ')" -ForegroundColor Yellow
                $testResults.TestsPassed++
                $testResults.Details.ProcessingSettings = @{ Status = "Warning"; Validation = $processingValidation }
            }
        } catch {
            Write-Host "   ✗ Processing settings validation failed: $($_.Exception.Message)" -ForegroundColor Red
            $testResults.TestsFailed++
            $testResults.Details.ProcessingSettings = @{ Status = "Failed"; Error = $_.Exception.Message }
        }
        
        # Calculate overall status
        $testResults.OverallStatus = if ($testResults.TestsFailed -eq 0) {
            $(if ($testResults.TestsPassed -eq $testResults.TestsRun) { "Passed" } else { "Warning" })
        } else {
            "Failed"
        }
        
        # Display summary
        Write-Host "`nConfiguration Test Summary:" -ForegroundColor Green
        Write-Host "Overall Status: $($testResults.OverallStatus)" -ForegroundColor $(if ($testResults.OverallStatus -eq "Passed") { "Green" } elseif ($testResults.OverallStatus -eq "Warning") { "Yellow" } else { "Red" })
        Write-Host "Tests Run: $($testResults.TestsRun)" -ForegroundColor Cyan
        Write-Host "Tests Passed: $($testResults.TestsPassed)" -ForegroundColor Green
        Write-Host "Tests Failed: $($testResults.TestsFailed)" -ForegroundColor Red
        
        return $testResults
        
    } catch {
        Write-Error "Configuration testing failed: $($_.Exception.Message)"
        
        return @{
            OverallStatus = "Error"
            TestsRun = 0
            TestsPassed = 0
            TestsFailed = 1
            Details = @{ Error = $_.Exception.Message }
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
    }
}
'@
    
    # Replace the problematic function
    $functionPattern = 'function Test-RAGConfiguration \{[\s\S]*?^Write-Verbose "RAGConfigManager_v2 module loaded successfully"'
    if ($content -match $functionPattern) {
        $content = $content -replace $functionPattern, ($correctedFunction + "`n`nWrite-Verbose `"RAGConfigManager_v2 module loaded successfully`"")
    }
    
    $content | Set-Content $ragConfigPath -Force
    Write-Host "  Fixed RAGConfigManager_v2.psm1" -ForegroundColor Green
}

Write-Host "`nAll fixes applied!" -ForegroundColor Green