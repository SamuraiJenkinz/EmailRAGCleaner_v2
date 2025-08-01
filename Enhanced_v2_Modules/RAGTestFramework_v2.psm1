# RAGTestFramework_v2.psm1 - Comprehensive Testing Framework for Email RAG Pipeline
# End-to-end testing suite for validating RAG pipeline components and performance

Export-ModuleMember -Function @(
    'Start-RAGPipelineTest',
    'Test-EmailProcessingPipeline',
    'Test-ChunkingQuality',
    'Test-SearchAccuracy',
    'Test-EntityExtraction',
    'Generate-TestReport',
    'Create-TestDataset',
    'Validate-RAGPerformance'
)

# Import required modules
Import-Module (Join-Path $PSScriptRoot "EmailRAGProcessor_v2.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "EmailChunkingEngine_v2.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "EmailEntityExtractor_v2.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "EmailSearchInterface_v2.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "RAGConfigManager_v2.psm1") -Force

function Start-RAGPipelineTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$RAGConfiguration,
        
        [Parameter(Mandatory=$false)]
        [string]$TestDataPath,
        
        [Parameter(Mandatory=$false)]
        [string]$ReportOutputPath = "RAG_Test_Report.html",
        
        [Parameter(Mandatory=$false)]
        [hashtable]$TestSettings = @{},
        
        [Parameter(Mandatory=$false)]
        [switch]$DetailedReport = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$CleanupAfterTest = $false
    )
    
    try {
        Write-Host "Starting comprehensive RAG pipeline test..." -ForegroundColor Green
        Write-Host "=================================" -ForegroundColor Green
        
        $testResults = @{
            StartTime = Get-Date
            TestConfiguration = $RAGConfiguration.Metadata.Name
            TestSettings = $TestSettings
            Results = @{}
            Summary = @{}
            Recommendations = @()
        }
        
        # Test 1: Configuration Validation
        Write-Host "`n1. Testing RAG Configuration..." -ForegroundColor Yellow
        $configResult = Test-RAGConfiguration -Configuration $RAGConfiguration -TestConnections:$true -TestPermissions:$true
        $testResults.Results.Configuration = $configResult
        
        if ($configResult.OverallStatus -ne "Passed") {
            Write-Host "   ⚠ Configuration issues detected. Continuing with limited tests..." -ForegroundColor Yellow
        }
        
        # Test 2: Email Processing Pipeline
        Write-Host "`n2. Testing Email Processing Pipeline..." -ForegroundColor Yellow
        if ($TestDataPath -and (Test-Path $TestDataPath)) {
            $processingResult = Test-EmailProcessingPipeline -RAGConfiguration $RAGConfiguration -TestDataPath $TestDataPath -MaxTestFiles 5
            $testResults.Results.Processing = $processingResult
        } else {
            Write-Host "   ⚠ No test data provided. Creating synthetic test data..." -ForegroundColor Yellow
            $syntheticData = Create-TestDataset -OutputPath "TestData" -Count 3
            $processingResult = Test-EmailProcessingPipeline -RAGConfiguration $RAGConfiguration -TestDataPath $syntheticData.OutputPath -MaxTestFiles 3
            $testResults.Results.Processing = $processingResult
        }
        
        # Test 3: Chunking Quality
        Write-Host "`n3. Testing Chunking Quality..." -ForegroundColor Yellow
        $chunkingResult = Test-ChunkingQuality -TestEmails $processingResult.ProcessedEmails -RAGConfiguration $RAGConfiguration
        $testResults.Results.Chunking = $chunkingResult
        
        # Test 4: Entity Extraction
        Write-Host "`n4. Testing Entity Extraction..." -ForegroundColor Yellow
        $entityResult = Test-EntityExtraction -TestEmails $processingResult.ProcessedEmails
        $testResults.Results.EntityExtraction = $entityResult
        
        # Test 5: Search Accuracy (if index has data)
        Write-Host "`n5. Testing Search Accuracy..." -ForegroundColor Yellow
        if ($processingResult.IndexedDocuments -gt 0) {
            $searchResult = Test-SearchAccuracy -RAGConfiguration $RAGConfiguration -TestQueries @("meeting", "project", "urgent", "attachment")
            $testResults.Results.Search = $searchResult
        } else {
            Write-Host "   ⚠ No indexed documents available for search testing" -ForegroundColor Yellow
            $testResults.Results.Search = @{ Status = "Skipped"; Reason = "No indexed documents" }
        }
        
        # Test 6: Performance Validation
        Write-Host "`n6. Testing Performance..." -ForegroundColor Yellow
        $performanceResult = Validate-RAGPerformance -RAGConfiguration $RAGConfiguration -ProcessingResults $processingResult
        $testResults.Results.Performance = $performanceResult
        
        # Calculate overall results
        $testResults.EndTime = Get-Date
        $testResults.TotalDuration = ($testResults.EndTime - $testResults.StartTime).TotalSeconds
        $testResults.Summary = Calculate-TestSummary -TestResults $testResults
        $testResults.Recommendations = Generate-TestRecommendations -TestResults $testResults
        
        # Generate report
        if ($DetailedReport) {
            $reportResult = Generate-TestReport -TestResults $testResults -OutputPath $ReportOutputPath
            Write-Host "`nDetailed test report generated: $($reportResult.ReportPath)" -ForegroundColor Green
        }
        
        # Cleanup if requested
        if ($CleanupAfterTest -and $syntheticData) {
            Remove-Item -Path $syntheticData.OutputPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Test data cleaned up" -ForegroundColor Gray
        }
        
        # Display summary
        Display-TestSummary -TestResults $testResults
        
        return $testResults
        
    } catch {
        Write-Error "RAG pipeline test failed: $($_.Exception.Message)"
        
        return @{
            StartTime = Get-Date
            EndTime = Get-Date
            Status = "Failed"
            Error = $_.Exception.Message
            Results = @{}
        }
    }
}

function Test-EmailProcessingPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$RAGConfiguration,
        
        [Parameter(Mandatory=$true)]
        [string]$TestDataPath,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxTestFiles = 10
    )
    
    try {
        Write-Verbose "Testing email processing pipeline with data from: $TestDataPath"
        
        $testResult = @{
            Status = "Unknown"
            ProcessedEmails = @()
            TotalFiles = 0
            ProcessedFiles = 0
            FailedFiles = 0
            IndexedDocuments = 0
            AverageProcessingTime = 0
            Errors = @()
            StartTime = Get-Date
        }
        
        # Find test MSG files
        $testFiles = Get-ChildItem -Path $TestDataPath -Filter "*.msg" -Recurse | Select-Object -First $MaxTestFiles
        $testResult.TotalFiles = $testFiles.Count
        
        if ($testFiles.Count -eq 0) {
            Write-Warning "No MSG files found in test data path: $TestDataPath"
            $testResult.Status = "NoTestData"
            return $testResult
        }
        
        Write-Host "   Processing $($testFiles.Count) test email files..." -ForegroundColor Cyan
        
        # Initialize RAG pipeline
        $pipelineConfig = Initialize-RAGPipeline -AzureSearchConfig @{
            ServiceName = $RAGConfiguration.AzureSearch.ServiceName
            ServiceUrl = $RAGConfiguration.AzureSearch.ServiceUrl
            ApiKey = $RAGConfiguration.AzureSearch.ApiKey
            ApiVersion = $RAGConfiguration.AzureSearch.ApiVersion
            Headers = @{
                'api-key' = $RAGConfiguration.AzureSearch.ApiKey
                'Content-Type' = 'application/json'
            }
            OpenAI = $RAGConfiguration.OpenAI
        } -IndexName $RAGConfiguration.AzureSearch.IndexName -CreateIndex:$false
        
        # Process each test file
        $processingTimes = @()
        foreach ($testFile in $testFiles) {
            try {
                $startTime = Get-Date
                $processResult = Process-EmailForRAG -PipelineConfig $pipelineConfig -EmailFilePath $testFile.FullName
                $endTime = Get-Date
                $processingTime = ($endTime - $startTime).TotalSeconds
                $processingTimes += $processingTime
                
                if ($processResult.Status -eq "Success") {
                    $testResult.ProcessedFiles++
                    $testResult.IndexedDocuments += $processResult.IndexedDocuments
                    $testResult.ProcessedEmails += @{
                        FileName = $testFile.Name
                        ProcessingTime = $processingTime
                        ChunkCount = $processResult.ChunkCount
                        EmailData = $processResult.EmailData
                        QualityReport = $processResult.QualityReport
                    }
                } else {
                    $testResult.FailedFiles++
                    $testResult.Errors += @{
                        FileName = $testFile.Name
                        Error = $processResult.Error
                    }
                }
            } catch {
                $testResult.FailedFiles++
                $testResult.Errors += @{
                    FileName = $testFile.Name
                    Error = $_.Exception.Message
                }
            }
        }
        
        # Calculate results
        $testResult.EndTime = Get-Date
        if ($processingTimes.Count -gt 0) {
            $testResult.AverageProcessingTime = [Math]::Round(($processingTimes | Measure-Object -Average).Average, 2)
        } else {
            $testResult.AverageProcessingTime = 0
        }
        
        if ($testResult.ProcessedFiles -eq $testResult.TotalFiles) {
            $testResult.Status = "Success"
        } elseif ($testResult.ProcessedFiles -gt 0) {
            $testResult.Status = "Partial"
        } else {
            $testResult.Status = "Failed"
        }
        
        Write-Host "   ✓ Processing complete: $($testResult.ProcessedFiles)/$($testResult.TotalFiles) files processed" -ForegroundColor Green
        Write-Host "   ✓ Average processing time: $($testResult.AverageProcessingTime) seconds" -ForegroundColor Green
        Write-Host "   ✓ Documents indexed: $($testResult.IndexedDocuments)" -ForegroundColor Green
        
        return $testResult
        
    } catch {
        Write-Error "Email processing pipeline test failed: $($_.Exception.Message)"
        
        return @{
            Status = "Error"
            Error = $_.Exception.Message
            ProcessedEmails = @()
            TotalFiles = 0
            ProcessedFiles = 0
            FailedFiles = 0
        }
    }
}

function Test-ChunkingQuality {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$TestEmails,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$RAGConfiguration
    )
    
    try {
        Write-Verbose "Testing chunking quality for $($TestEmails.Count) emails"
        
        $qualityResult = @{
            Status = "Unknown"
            TotalEmails = $TestEmails.Count
            TotalChunks = 0
            QualityMetrics = @{
                AverageTokenCount = 0
                OptimalSizePercentage = 0
                AverageQualityScore = 0
                SearchReadinessPercentage = 0
            }
            Issues = @()
            Recommendations = @()
        }
        
        if ($TestEmails.Count -eq 0) {
            $qualityResult.Status = "NoData"
            return $qualityResult
        }
        
        $allChunks = @()
        $qualityScores = @()
        $tokenCounts = @()
        $searchReadyCount = 0
        
        foreach ($email in $TestEmails) {
            if ($email.QualityReport -and $email.QualityReport.Chunks) {
                foreach ($chunk in $email.QualityReport.Chunks) {
                    $allChunks += $chunk
                    $qualityScores += $chunk.QualityScore
                    $tokenCounts += $chunk.TokenCount
                    
                    if ($chunk.SearchReadiness -and $chunk.SearchReadiness.IsReady) {
                        $searchReadyCount++
                    }
                    
                    # Check for quality issues
                    if ($chunk.TokenCount -lt 64) {
                        $qualityResult.Issues += "Chunk too small: $($chunk.TokenCount) tokens in $($email.FileName)"
                    }
                    
                    if ($chunk.TokenCount -gt 600) {
                        $qualityResult.Issues += "Chunk too large: $($chunk.TokenCount) tokens in $($email.FileName)"
                    }
                    
                    if ($chunk.QualityScore -lt 50) {
                        $qualityResult.Issues += "Low quality chunk: $($chunk.QualityScore) score in $($email.FileName)"
                    }
                }
            }
        }
        
        $qualityResult.TotalChunks = $allChunks.Count
        
        if ($allChunks.Count -gt 0) {
            # Calculate quality metrics
            $qualityResult.QualityMetrics.AverageTokenCount = [Math]::Round(($tokenCounts | Measure-Object -Average).Average, 0)
            $qualityResult.QualityMetrics.AverageQualityScore = [Math]::Round(($qualityScores | Measure-Object -Average).Average, 1)
            
            $optimalSizeChunks = $allChunks | Where-Object { $_.TokenCount -ge 256 -and $_.TokenCount -le 512 }
            $qualityResult.QualityMetrics.OptimalSizePercentage = [Math]::Round(($optimalSizeChunks.Count / $allChunks.Count) * 100, 1)
            
            $qualityResult.QualityMetrics.SearchReadinessPercentage = [Math]::Round(($searchReadyCount / $allChunks.Count) * 100, 1)
            
            # Generate recommendations
            if ($qualityResult.QualityMetrics.OptimalSizePercentage -lt 70) {
                $qualityResult.Recommendations += "Consider adjusting chunk size settings - only $($qualityResult.QualityMetrics.OptimalSizePercentage)% of chunks are optimally sized"
            }
            
            if ($qualityResult.QualityMetrics.AverageQualityScore -lt 70) {
                $qualityResult.Recommendations += "Average quality score ($($qualityResult.QualityMetrics.AverageQualityScore)) is below recommended threshold of 70"
            }
            
            if ($qualityResult.QualityMetrics.SearchReadinessPercentage -lt 90) {
                $qualityResult.Recommendations += "Search readiness ($($qualityResult.QualityMetrics.SearchReadinessPercentage)%) could be improved"
            }
            
            $qualityResult.Status = "Success"
        } else {
            $qualityResult.Status = "NoChunks"
        }
        
        Write-Host "   ✓ Chunking analysis: $($qualityResult.TotalChunks) chunks, avg $($qualityResult.QualityMetrics.AverageTokenCount) tokens" -ForegroundColor Green
        Write-Host "   ✓ Quality score: $($qualityResult.QualityMetrics.AverageQualityScore), optimal size: $($qualityResult.QualityMetrics.OptimalSizePercentage)%" -ForegroundColor Green
        
        return $qualityResult
        
    } catch {
        Write-Error "Chunking quality test failed: $($_.Exception.Message)"
        
        return @{
            Status = "Error"
            Error = $_.Exception.Message
            TotalEmails = $TestEmails.Count
            TotalChunks = 0
        }
    }
}

function Test-SearchAccuracy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$RAGConfiguration,
        
        [Parameter(Mandatory=$true)]
        [array]$TestQueries,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxResultsPerQuery = 10
    )
    
    try {
        Write-Verbose "Testing search accuracy with $($TestQueries.Count) test queries"
        
        $searchResult = @{
            Status = "Unknown"
            TestedQueries = $TestQueries.Count
            Results = @()
            Metrics = @{
                AverageResponseTime = 0
                AverageResultCount = 0
                QueriesWithResults = 0
                SuccessfulQueries = 0
            }
        }
        
        $searchConfig = @{
            ServiceName = $RAGConfiguration.AzureSearch.ServiceName
            ServiceUrl = $RAGConfiguration.AzureSearch.ServiceUrl
            ApiKey = $RAGConfiguration.AzureSearch.ApiKey
            ApiVersion = $RAGConfiguration.AzureSearch.ApiVersion
            Headers = @{
                'api-key' = $RAGConfiguration.AzureSearch.ApiKey
                'Content-Type' = 'application/json'
            }
            OpenAI = $RAGConfiguration.OpenAI
        }
        
        $responseTimes = @()
        $resultCounts = @()
        
        foreach ($query in $TestQueries) {
            try {
                $startTime = Get-Date
                $queryResult = Find-EmailContent -SearchConfig $searchConfig -Query $query -IndexName $RAGConfiguration.AzureSearch.IndexName -MaxResults $MaxResultsPerQuery -SearchType "Hybrid"
                $endTime = Get-Date
                $responseTime = ($endTime - $startTime).TotalMilliseconds
                
                $responseTimes += $responseTime
                $resultCounts += $queryResult.TotalResults
                
                $searchResult.Results += @{
                    Query = $query
                    ResultCount = $queryResult.TotalResults
                    ResponseTime = $responseTime
                    HasResults = $queryResult.TotalResults -gt 0
                    Status = "Success"
                    TopResults = $queryResult.Results | Select-Object -First 3 | ForEach-Object { 
                        @{ Subject = $_.Subject; Score = $_.SearchScore } 
                    }
                }
                
                if ($queryResult.TotalResults -gt 0) {
                    $searchResult.Metrics.QueriesWithResults++
                }
                
                $searchResult.Metrics.SuccessfulQueries++
                
            } catch {
                $searchResult.Results += @{
                    Query = $query
                    ResultCount = 0
                    ResponseTime = 0
                    HasResults = $false
                    Status = "Failed"
                    Error = $_.Exception.Message
                }
            }
        }
        
        # Calculate metrics
        if ($responseTimes.Count -gt 0) {
            $searchResult.Metrics.AverageResponseTime = [Math]::Round(($responseTimes | Measure-Object -Average).Average, 0)
            $searchResult.Metrics.AverageResultCount = [Math]::Round(($resultCounts | Measure-Object -Average).Average, 1)
        }
        
        if ($searchResult.Metrics.SuccessfulQueries -eq $TestQueries.Count) {
            $searchResult.Status = "Success"
        } elseif ($searchResult.Metrics.SuccessfulQueries -gt 0) {
            $searchResult.Status = "Partial"
        } else {
            $searchResult.Status = "Failed"
        }
        
        Write-Host "   ✓ Search testing: $($searchResult.Metrics.SuccessfulQueries)/$($TestQueries.Count) queries successful" -ForegroundColor Green
        Write-Host "   ✓ Average response time: $($searchResult.Metrics.AverageResponseTime)ms, avg results: $($searchResult.Metrics.AverageResultCount)" -ForegroundColor Green
        
        return $searchResult
        
    } catch {
        Write-Error "Search accuracy test failed: $($_.Exception.Message)"
        
        return @{
            Status = "Error"
            Error = $_.Exception.Message
            TestedQueries = $TestQueries.Count
            Results = @()
        }
    }
}

function Test-EntityExtraction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$TestEmails
    )
    
    try {
        Write-Verbose "Testing entity extraction for $($TestEmails.Count) emails"
        
        $entityResult = @{
            Status = "Unknown" 
            TestedEmails = $TestEmails.Count
            Metrics = @{
                AverageEntitiesPerEmail = 0
                MostCommonEntityTypes = @()
                ExtractionSuccessRate = 0
                AverageConfidenceScore = 0
            }
            EntityCounts = @{}
            Issues = @()
        }
        
        if ($TestEmails.Count -eq 0) {
            $entityResult.Status = "NoData"
            return $entityResult
        }
        
        $successfulExtractions = 0
        $totalEntityCounts = @()
        $allConfidenceScores = @()
        $entityTypeCounts = @{}
        
        foreach ($email in $TestEmails) {
            if ($email.EmailData -and $email.EmailData.Content -and $email.EmailData.Content.ExtractedEntities) {
                $entities = $email.EmailData.Content.ExtractedEntities
                $successfulExtractions++
                
                $emailEntityCount = $entities.EntityCount
                $totalEntityCounts += $emailEntityCount
                
                # Count entity types
                foreach ($entityType in @('Emails', 'URLs', 'PhoneNumbers', 'Dates', 'IPAddresses', 'Numbers')) {
                    if ($entities.$entityType -and $entities.$entityType.Count -gt 0) {
                        if (-not $entityTypeCounts.ContainsKey($entityType)) {
                            $entityTypeCounts[$entityType] = 0
                        }
                        $entityTypeCounts[$entityType] += $entities.$entityType.Count
                        
                        # Collect confidence scores if available
                        foreach ($entity in $entities.$entityType) {
                            if ($entity.ConfidenceScore) {
                                $allConfidenceScores += $entity.ConfidenceScore
                            }
                        }
                    }
                }
                
                # Check for issues
                if ($emailEntityCount -eq 0) {
                    $entityResult.Issues += "No entities extracted from $($email.FileName)"
                }
            } else {
                $entityResult.Issues += "Entity extraction data missing for $($email.FileName)"
            }
        }
        
        # Calculate metrics
        if ($totalEntityCounts.Count -gt 0) {
            $entityResult.Metrics.AverageEntitiesPerEmail = [Math]::Round(($totalEntityCounts | Measure-Object -Average).Average, 1)
        }
        
        if ($allConfidenceScores.Count -gt 0) {
            $entityResult.Metrics.AverageConfidenceScore = [Math]::Round(($allConfidenceScores | Measure-Object -Average).Average, 2)
        }
        
        $entityResult.Metrics.ExtractionSuccessRate = [Math]::Round(($successfulExtractions / $TestEmails.Count) * 100, 1)
        
        # Most common entity types
        $entityResult.EntityCounts = $entityTypeCounts
        $entityResult.Metrics.MostCommonEntityTypes = $entityTypeCounts.GetEnumerator() | 
            Sort-Object Value -Descending | 
            Select-Object -First 5 | 
            ForEach-Object { "$($_.Key): $($_.Value)" }
        
        if ($successfulExtractions -eq $TestEmails.Count) {
            $entityResult.Status = "Success"
        } elseif ($successfulExtractions -gt 0) {
            $entityResult.Status = "Partial"
        } else {
            $entityResult.Status = "Failed"
        }
        
        Write-Host "   ✓ Entity extraction: $($entityResult.Metrics.ExtractionSuccessRate)% success rate" -ForegroundColor Green
        Write-Host "   ✓ Average entities per email: $($entityResult.Metrics.AverageEntitiesPerEmail), confidence: $($entityResult.Metrics.AverageConfidenceScore)" -ForegroundColor Green
        
        return $entityResult
        
    } catch {
        Write-Error "Entity extraction test failed: $($_.Exception.Message)"
        
        return @{
            Status = "Error"
            Error = $_.Exception.Message
            TestedEmails = $TestEmails.Count
        }
    }
}

function Validate-RAGPerformance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$RAGConfiguration,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$ProcessingResults
    )
    
    try {
        Write-Verbose "Validating RAG pipeline performance"
        
        $performanceResult = @{
            Status = "Unknown"
            Metrics = @{
                ThroughputEmailsPerMinute = 0
                AverageMemoryUsageMB = 0
                IndexingEfficiency = 0
                OverallPerformanceGrade = "Unknown"
            }
            Benchmarks = @{
                ProcessingSpeed = "Unknown"
                ResourceUtilization = "Unknown"
                ScalabilityRating = "Unknown"
            }
            Recommendations = @()
        }
        
        # Calculate throughput
        if ($ProcessingResults.AverageProcessingTime -gt 0) {
            $performanceResult.Metrics.ThroughputEmailsPerMinute = [Math]::Round(60 / $ProcessingResults.AverageProcessingTime, 1)
            
            # Performance benchmarks
            if ($ProcessingResults.AverageProcessingTime -lt 5) {
                $performanceResult.Benchmarks.ProcessingSpeed = "Excellent"
            } elseif ($ProcessingResults.AverageProcessingTime -lt 15) {
                $performanceResult.Benchmarks.ProcessingSpeed = "Good"
            } elseif ($ProcessingResults.AverageProcessingTime -lt 30) {
                $performanceResult.Benchmarks.ProcessingSpeed = "Average"
            } else {
                $performanceResult.Benchmarks.ProcessingSpeed = "Poor"
                $performanceResult.Recommendations += "Consider optimizing processing pipeline - average processing time is $($ProcessingResults.AverageProcessingTime) seconds"
            }
        }
        
        # Calculate indexing efficiency
        if ($ProcessingResults.ProcessedFiles -gt 0 -and $ProcessingResults.IndexedDocuments -gt 0) {
            $performanceResult.Metrics.IndexingEfficiency = [Math]::Round(($ProcessingResults.IndexedDocuments / $ProcessingResults.ProcessedFiles), 1)
            
            if ($performanceResult.Metrics.IndexingEfficiency -lt 5) {
                $performanceResult.Recommendations += "Low indexing efficiency - only $($performanceResult.Metrics.IndexingEfficiency) documents per email indexed"
            }
        }
        
        # Memory usage estimation (basic)
        $estimatedMemoryMB = ($ProcessingResults.ProcessedFiles * 2) + ($ProcessingResults.IndexedDocuments * 0.1)
        $performanceResult.Metrics.AverageMemoryUsageMB = [Math]::Round($estimatedMemoryMB, 1)
        
        if ($performanceResult.Metrics.AverageMemoryUsageMB -gt 1024) {
            $performanceResult.Benchmarks.ResourceUtilization = "High"
            $performanceResult.Recommendations += "High memory usage detected ($($performanceResult.Metrics.AverageMemoryUsageMB) MB) - consider batch size optimization"
        } else {
            $performanceResult.Benchmarks.ResourceUtilization = "Normal"
        }
        
        # Overall performance grade
        $gradePoints = 0
        if ($performanceResult.Benchmarks.ProcessingSpeed -eq "Excellent") { $gradePoints += 3 }
        elseif ($performanceResult.Benchmarks.ProcessingSpeed -eq "Good") { $gradePoints += 2 }
        elseif ($performanceResult.Benchmarks.ProcessingSpeed -eq "Average") { $gradePoints += 1 }
        
        if ($performanceResult.Benchmarks.ResourceUtilization -eq "Normal") { $gradePoints += 2 }
        elseif ($performanceResult.Benchmarks.ResourceUtilization -eq "High") { $gradePoints += 1 }
        
        if ($performanceResult.Metrics.IndexingEfficiency -ge 5) { $gradePoints += 2 }
        elseif ($performanceResult.Metrics.IndexingEfficiency -ge 3) { $gradePoints += 1 }
        
        $performanceResult.Metrics.OverallPerformanceGrade = switch ($gradePoints) {
            { $_ -ge 6 } { "A - Excellent" }
            { $_ -ge 4 } { "B - Good" }
            { $_ -ge 2 } { "C - Average" }
            default { "D - Needs Improvement" }
        }
        
        $performanceResult.Status = "Success"
        
        Write-Host "   ✓ Performance grade: $($performanceResult.Metrics.OverallPerformanceGrade)" -ForegroundColor Green
        Write-Host "   ✓ Throughput: $($performanceResult.Metrics.ThroughputEmailsPerMinute) emails/min, memory: $($performanceResult.Metrics.AverageMemoryUsageMB) MB" -ForegroundColor Green
        
        return $performanceResult
        
    } catch {
        Write-Error "Performance validation failed: $($_.Exception.Message)"
        
        return @{
            Status = "Error"
            Error = $_.Exception.Message
            Metrics = @{}
        }
    }
}

function Create-TestDataset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [int]$Count = 5
    )
    
    try {
        Write-Host "Creating synthetic test dataset..." -ForegroundColor Yellow
        
        # Create test data directory
        if (Test-Path $OutputPath) {
            Remove-Item -Path $OutputPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        
        $testEmails = @(
            @{
                Subject = "Project Alpha Status Update"
                Body = "Hi team, I wanted to provide an update on Project Alpha. We've completed the initial phase and are moving into development. The deadline is next Friday. Please review the attached documents and let me know if you have any questions. Best regards, John"
                Sender = "john.doe@company.com"
                Recipients = @("team@company.com")
                HasAttachments = $true
            },
            @{
                Subject = "Urgent: Server Maintenance Window"
                Body = "URGENT: We need to schedule a maintenance window for our production servers this weekend. The maintenance will start at 2:00 AM EST on Saturday and should complete by 6:00 AM. Please confirm your availability. Contact: +1-555-0123"
                Sender = "admin@company.com"
                Recipients = @("devops@company.com", "operations@company.com")
                HasAttachments = $false
            },
            @{
                Subject = "Meeting Request: Q4 Planning"
                Body = "Hello everyone, I'm scheduling our Q4 planning meeting for next Wednesday at 10:00 AM in Conference Room B. We'll discuss our goals, budget allocation, and resource planning. Please bring your project status reports. Thanks, Sarah"
                Sender = "sarah.smith@company.com"
                Recipients = @("managers@company.com")
                HasAttachments = $false
            }
        )
        
        # Create MSG files (simplified - in real scenario, would use proper MSG format)
        for ($i = 0; $i -lt [Math]::Min($Count, $testEmails.Count); $i++) {
            $email = $testEmails[$i]
            $fileName = "TestEmail_$($i + 1).msg"
            $filePath = Join-Path $OutputPath $fileName
            
            # Create a mock MSG file (for testing purposes)
            $mockMsgContent = @"
Subject: $($email.Subject)
From: $($email.Sender)
To: $($email.Recipients -join '; ')
Date: $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))

$($email.Body)
"@
            
            $mockMsgContent | Out-File -FilePath $filePath -Encoding UTF8
        }
        
        Write-Host "   ✓ Created $($testEmails.Count) test email files in: $OutputPath" -ForegroundColor Green
        
        return @{
            OutputPath = $OutputPath
            FileCount = $testEmails.Count
            CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to create test dataset: $($_.Exception.Message)"
        throw
    }
}

function Generate-TestReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$TestResults,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    try {
        Write-Verbose "Generating comprehensive test report"
        
        $reportHtml = @"
<!DOCTYPE html>
<html>
<head>
    <title>RAG Pipeline Test Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
        .success { background-color: #d4edda; border-color: #c3e6cb; }
        .warning { background-color: #fff3cd; border-color: #ffeaa7; }
        .error { background-color: #f8d7da; border-color: #f5c6cb; }
        .metric { display: inline-block; margin: 10px; padding: 10px; background: #f8f9fa; border-radius: 5px; min-width: 150px; }
        .metric-value { font-size: 1.5em; font-weight: bold; color: #495057; }
        .metric-label { font-size: 0.9em; color: #6c757d; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f8f9fa; }
        .recommendation { background: #e3f2fd; padding: 10px; margin: 5px 0; border-left: 4px solid #2196f3; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Email RAG Pipeline Test Report</h1>
        <p>Configuration: $($TestResults.TestConfiguration)</p>
        <p>Generated: $($TestResults.EndTime.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        <p>Total Duration: $([Math]::Round($TestResults.TotalDuration, 1)) seconds</p>
    </div>

    <div class="section">
        <h2>Test Summary</h2>
        <div class="metric">
            <div class="metric-value">$($TestResults.Summary.OverallStatus)</div>
            <div class="metric-label">Overall Status</div>
        </div>
        <div class="metric">
            <div class="metric-value">$($TestResults.Summary.TestsPassed)/$($TestResults.Summary.TotalTests)</div>
            <div class="metric-label">Tests Passed</div>
        </div>
        <div class="metric">
            <div class="metric-value">$($TestResults.Summary.SuccessRate)%</div>
            <div class="metric-label">Success Rate</div>
        </div>
    </div>
"@

        # Add each test section
        foreach ($testName in $TestResults.Results.Keys) {
            $testResult = $TestResults.Results[$testName]
            $sectionClass = switch ($testResult.Status) {
                "Success" { "success" }
                "Passed" { "success" }
                { $_ -in @("Partial", "Warning") } { "warning" }
                default { "error" }
            }
            
            $reportHtml += @"
    <div class="section $sectionClass">
        <h2>$testName Test Results</h2>
        <p><strong>Status:</strong> $($testResult.Status)</p>
"@
            
            # Add test-specific details
            if ($testName -eq "Processing" -and $testResult.ProcessedFiles) {
                $reportHtml += @"
        <p><strong>Files Processed:</strong> $($testResult.ProcessedFiles)/$($testResult.TotalFiles)</p>
        <p><strong>Documents Indexed:</strong> $($testResult.IndexedDocuments)</p>
        <p><strong>Average Processing Time:</strong> $($testResult.AverageProcessingTime) seconds</p>
"@
            }
            
            if ($testName -eq "Chunking" -and $testResult.QualityMetrics) {
                $reportHtml += @"
        <p><strong>Total Chunks:</strong> $($testResult.TotalChunks)</p>
        <p><strong>Average Token Count:</strong> $($testResult.QualityMetrics.AverageTokenCount)</p>
        <p><strong>Quality Score:</strong> $($testResult.QualityMetrics.AverageQualityScore)</p>
        <p><strong>Optimal Size Percentage:</strong> $($testResult.QualityMetrics.OptimalSizePercentage)%</p>
"@
            }
            
            $reportHtml += @"
    </div>
"@
        }
        
        # Add recommendations
        if ($TestResults.Recommendations -and $TestResults.Recommendations.Count -gt 0) {
            $reportHtml += @"
    <div class="section">
        <h2>Recommendations</h2>
"@
            foreach ($recommendation in $TestResults.Recommendations) {
                $reportHtml += @"
        <div class="recommendation">$recommendation</div>
"@
            }
            $reportHtml += @"
    </div>
"@
        }
        
        $reportHtml += @"
</body>
</html>
"@
        
        # Write report to file
        $reportHtml | Out-File -FilePath $OutputPath -Encoding UTF8
        
        return @{
            ReportPath = $OutputPath
            GeneratedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            FileSize = (Get-Item $OutputPath).Length
        }
        
    } catch {
        Write-Error "Failed to generate test report: $($_.Exception.Message)"
        throw
    }
}

function Calculate-TestSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$TestResults
    )
    
    $totalTests = 0
    $passedTests = 0
    $failedTests = 0
    
    foreach ($testName in $TestResults.Results.Keys) {
        $testResult = $TestResults.Results[$testName]
        $totalTests++
        
        if ($testResult.Status -in @("Success", "Passed")) {
            $passedTests++
        } elseif ($testResult.Status -in @("Failed", "Error")) {
            $failedTests++
        }
    }
    
    if ($failedTests -eq 0) {
        if ($passedTests -eq $totalTests) {
            $overallStatus = "Passed"
        } else {
            $overallStatus = "Warning"
        }
    } else {
        $overallStatus = "Failed"
    }
    
    return @{
        TotalTests = $totalTests
        TestsPassed = $passedTests
        TestsFailed = $failedTests
        SuccessRate = $(if ($totalTests -gt 0) { [Math]::Round(($passedTests / $totalTests) * 100, 1) } else { 0 })
        OverallStatus = $overallStatus
    }
}

function Generate-TestRecommendations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$TestResults
    )
    
    $recommendations = @()
    
    # Configuration recommendations
    if ($TestResults.Results.Configuration -and $TestResults.Results.Configuration.OverallStatus -ne "Passed") {
        $recommendations += "Review and fix configuration issues before production deployment"
    }
    
    # Processing recommendations
    if ($TestResults.Results.Processing) {
        $processing = $TestResults.Results.Processing
        if ($processing.AverageProcessingTime -gt 30) {
            $recommendations += "Consider optimizing email processing pipeline - current average time is $($processing.AverageProcessingTime) seconds"
        }
        if ($processing.FailedFiles -gt 0) {
            $recommendations += "Investigate and resolve issues causing $($processing.FailedFiles) file processing failures"
        }
    }
    
    # Chunking recommendations
    if ($TestResults.Results.Chunking -and $TestResults.Results.Chunking.QualityMetrics) {
        $chunking = $TestResults.Results.Chunking.QualityMetrics
        if ($chunking.OptimalSizePercentage -lt 70) {
            $recommendations += "Adjust chunk size parameters to improve optimal size percentage (currently $($chunking.OptimalSizePercentage)%)"
        }
        if ($chunking.AverageQualityScore -lt 70) {
            $recommendations += "Improve content quality preprocessing to increase average quality score (currently $($chunking.AverageQualityScore))"
        }
    }
    
    # Search recommendations
    if ($TestResults.Results.Search -and $TestResults.Results.Search.Metrics) {
        $search = $TestResults.Results.Search.Metrics
        if ($search.AverageResponseTime -gt 1000) {
            $recommendations += "Optimize search performance - current average response time is $($search.AverageResponseTime)ms"
        }
        if ($search.QueriesWithResults -lt ($TestResults.Results.Search.TestedQueries * 0.8)) {
            $recommendations += "Consider improving index content or search relevance - only $($search.QueriesWithResults) out of $($TestResults.Results.Search.TestedQueries) queries returned results"
        }
    }
    
    return $recommendations
}

function Display-TestSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$TestResults
    )
    
    Write-Host "`n=================================" -ForegroundColor Green
    Write-Host "RAG PIPELINE TEST SUMMARY" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    
    Write-Host "Overall Status: $($TestResults.Summary.OverallStatus)" -ForegroundColor $(
        switch ($TestResults.Summary.OverallStatus) {
            "Passed" { "Green" }
            "Warning" { "Yellow" }
            "Failed" { "Red" }
            default { "Gray" }
        }
    )
    
    Write-Host "Tests Passed: $($TestResults.Summary.TestsPassed)/$($TestResults.Summary.TotalTests) ($($TestResults.Summary.SuccessRate)%)" -ForegroundColor Cyan
    Write-Host "Total Duration: $([Math]::Round($TestResults.TotalDuration, 1)) seconds" -ForegroundColor Cyan
    
    if ($TestResults.Recommendations -and $TestResults.Recommendations.Count -gt 0) {
        Write-Host "`nTop Recommendations:" -ForegroundColor Yellow
        for ($i = 0; $i -lt [Math]::Min(3, $TestResults.Recommendations.Count); $i++) {
            Write-Host "  $($i + 1). $($TestResults.Recommendations[$i])" -ForegroundColor White
        }
    }
    
    Write-Host "`n=================================" -ForegroundColor Green
}

Write-Verbose "RAGTestFramework_v2 module loaded successfully"