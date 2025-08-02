# RAGTestFramework_v2.psm1 - Comprehensive Testing Framework for Email RAG System
# Enhanced for PowerShell 7 with proper syntax and HTML generation

using namespace System.Collections.Generic
using namespace System.Text

Export-ModuleMember -Function @(
    'Test-EmailProcessingPipeline',
    'Test-EmailChunkingQuality',
    'Test-RAGSearchPerformance', 
    'Generate-TestReport',
    'Create-TestDataset',
    'Invoke-ComprehensiveRAGTest'
)

function Test-EmailProcessingPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TestDataPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$false)]
        [int]$SampleSize = 10,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath = "./test-results"
    )
    
    try {
        Write-Host "🧪 Starting Comprehensive Email Processing Pipeline Test" -ForegroundColor Cyan
        $testStartTime = Get-Date
        
        $testResults = @{
            TestName = "Email Processing Pipeline"
            StartTime = $testStartTime
            Configuration = $Configuration.Metadata?.Name ?? "Default Configuration"
            TestDataPath = $TestDataPath
            Results = @{}
            Summary = @{}
            Recommendations = [List[string]]::new()
        }
        
        # Test 1: Configuration Validation
        Write-Host "📋 Phase 1: Configuration Validation" -ForegroundColor Yellow
        $configTest = Test-ConfigurationIntegrity -Configuration $Configuration
        $testResults.Results.Configuration = $configTest
        
        # Test 2: File Processing Performance
        Write-Host "⚡ Phase 2: File Processing Performance" -ForegroundColor Yellow
        $processingTest = Test-ProcessingPerformance -TestDataPath $TestDataPath -Configuration $Configuration -SampleSize $SampleSize
        $testResults.Results.Processing = $processingTest
        
        # Test 3: Content Quality Assessment
        Write-Host "✨ Phase 3: Content Quality Assessment" -ForegroundColor Yellow
        $qualityTest = Test-ContentQuality -ProcessingResults $processingTest -Configuration $Configuration
        $testResults.Results.Quality = $qualityTest
        
        # Test 4: Search Integration Testing
        if ($Configuration.AzureSearch?.ServiceName) {
            Write-Host "🔍 Phase 4: Search Integration Testing" -ForegroundColor Yellow
            $searchTest = Test-SearchIntegration -Configuration $Configuration -SampleDocuments $processingTest.SampleDocuments
            $testResults.Results.Search = $searchTest
        }
        
        # Test 5: Performance Benchmarking
        Write-Host "📊 Phase 5: Performance Benchmarking" -ForegroundColor Yellow
        $performanceTest = Test-PerformanceBenchmarks -ProcessingResults $processingTest -Configuration $Configuration
        $testResults.Results.Performance = $performanceTest
        
        # Generate comprehensive summary
        $testResults.EndTime = Get-Date
        $testResults.TotalDuration = ($testResults.EndTime - $testResults.StartTime).TotalSeconds
        $testResults.Summary = Calculate-TestSummary -TestResults $testResults
        $testResults.Recommendations = Generate-TestRecommendations -TestResults $testResults
        
        # Generate detailed report
        if (-not [string]::IsNullOrEmpty($OutputPath)) {
            $reportPath = Join-Path $OutputPath "comprehensive-test-report.html"
            $reportResult = Generate-TestReport -TestResults $testResults -OutputPath $reportPath
            $testResults.ReportPath = $reportResult.ReportPath
        }
        
        # Display summary
        Display-TestSummary -TestResults $testResults
        
        return $testResults
        
    } catch {
        Write-Host "❌ Pipeline test failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Test-ConfigurationIntegrity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration
    )
    
    try {
        $configTest = @{
            Status = "Unknown"
            ValidationResults = @{}
            Issues = [List[string]]::new()
            Recommendations = [List[string]]::new()
        }
        
        # Import and test configuration module
        Import-Module "$PSScriptRoot\RAGConfigManager_v2_Fixed.psm1" -Force
        $validationResult = Test-RAGConfiguration -Configuration $Configuration -ValidateSettings $true -TestConnections $true
        
        $configTest.ValidationResults = $validationResult
        $configTest.Status = $validationResult.OverallStatus
        
        # Add specific recommendations based on validation
        if ($validationResult.TestsFailed -gt 0) {
            $configTest.Issues.Add("Configuration validation failed with $($validationResult.TestsFailed) failures")
        }
        
        if (-not $Configuration.OpenAI?.Enabled) {
            $configTest.Recommendations.Add("Consider enabling OpenAI integration for enhanced search capabilities")
        }
        
        return $configTest
        
    } catch {
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            Issues = @("Configuration test failed: $($_.Exception.Message)")
        }
    }
}

function Test-ProcessingPerformance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TestDataPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$true)]
        [int]$SampleSize
    )
    
    try {
        # Discover test files
        $testFiles = Get-ChildItem -Path $TestDataPath -Filter "*.msg" -Recurse | Select-Object -First $SampleSize
        
        if ($testFiles.Count -eq 0) {
            throw "No MSG files found in test data path: $TestDataPath"
        }
        
        Write-Host "   📧 Processing $($testFiles.Count) test files..." -ForegroundColor Gray
        
        # Import processing module
        Import-Module "$PSScriptRoot\EmailRAGProcessor_v2_Fixed.psm1" -Force
        
        # Run processing with timing
        $processingStartTime = Get-Date
        $processingResult = Invoke-EmailRAGProcessing -InputPath $TestDataPath -Configuration $Configuration -Parallel $true
        $processingEndTime = Get-Date
        
        $processingStats = Get-ProcessingStatistics -ProcessingResult $processingResult
        
        $performanceTest = @{
            Status = $processingResult.Status
            FilesProcessed = $processingStats.Summary.TotalFiles
            SuccessfulFiles = $processingStats.Summary.SuccessfulFiles
            FailedFiles = $processingStats.Summary.FailedFiles
            SuccessRate = $processingStats.Summary.SuccessRate
            TotalChunks = $processingStats.Summary.TotalChunks
            IndexedDocuments = $processingStats.Summary.IndexedDocuments
            ProcessingTime = $processingStats.Performance.TotalDuration
            AverageProcessingTime = $processingStats.Performance.AverageProcessingTime
            FilesPerSecond = $processingStats.Performance.FilesPerSecond
            SampleDocuments = $processingResult.ProcessedFiles ?? @()
            Errors = $processingStats.Errors
        }
        
        return $performanceTest
        
    } catch {
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            FilesProcessed = 0
            SuccessfulFiles = 0
            FailedFiles = 0
        }
    }
}

function Test-ContentQuality {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$ProcessingResults,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration
    )
    
    try {
        $qualityMetrics = @{
            TotalDocuments = $ProcessingResults.SampleDocuments.Count
            DocumentsWithContent = 0
            AverageContentLength = 0
            MinContentLength = [int]::MaxValue
            MaxContentLength = 0
            AverageChunkCount = 0
            OptimalSizeCount = 0
            QualityScore = 0
        }
        
        if ($qualityMetrics.TotalDocuments -eq 0) {
            return @{
                Status = "No Data"
                QualityMetrics = $qualityMetrics
                Recommendations = @("No processed documents available for quality assessment")
            }
        }
        
        $totalContentLength = 0
        $totalChunks = 0
        $qualityScores = [List[double]]::new()
        
        foreach ($doc in $ProcessingResults.SampleDocuments) {
            if ($doc.Status -eq "Success" -and $doc.ChunkCount -gt 0) {
                $qualityMetrics.DocumentsWithContent++
                $contentLength = $doc.DocumentSize ?? 0
                
                $totalContentLength += $contentLength
                $totalChunks += $doc.ChunkCount
                
                $qualityMetrics.MinContentLength = [Math]::Min($qualityMetrics.MinContentLength, $contentLength)
                $qualityMetrics.MaxContentLength = [Math]::Max($qualityMetrics.MaxContentLength, $contentLength)
                
                # Check if chunk count is optimal (simple heuristic)
                $targetChunks = [Math]::Max(1, [Math]::Ceiling($contentLength / 2000))
                if ([Math]::Abs($doc.ChunkCount - $targetChunks) -le 1) {
                    $qualityMetrics.OptimalSizeCount++
                }
                
                # Calculate quality score based on processing success and chunk generation
                $docQuality = ($doc.ChunkCount -gt 0 ? 0.5 : 0) + ($doc.Indexed ? 0.5 : 0)
                $qualityScores.Add($docQuality * 100)
            }
        }
        
        if ($qualityMetrics.DocumentsWithContent -gt 0) {
            $qualityMetrics.AverageContentLength = $totalContentLength / $qualityMetrics.DocumentsWithContent
            $qualityMetrics.AverageChunkCount = $totalChunks / $qualityMetrics.DocumentsWithContent
            $qualityMetrics.QualityScore = $qualityScores.Count -gt 0 ? ($qualityScores | Measure-Object -Average).Average : 0
        }
        
        $qualityMetrics.OptimalSizePercentage = $qualityMetrics.DocumentsWithContent -gt 0 ? 
            ($qualityMetrics.OptimalSizeCount / $qualityMetrics.DocumentsWithContent * 100) : 0
        
        $qualityMetrics.SearchReadinessPercentage = $ProcessingResults.IndexedDocuments -gt 0 ? 
            ($ProcessingResults.IndexedDocuments / $qualityMetrics.TotalDocuments * 100) : 0
        
        $qualityTest = @{
            Status = "Success"
            QualityMetrics = $qualityMetrics
            Recommendations = [List[string]]::new()
        }
        
        # Generate quality-based recommendations
        if ($qualityMetrics.QualityScore -lt 70) {
            $qualityTest.Recommendations.Add("Average quality score ($([Math]::Round($qualityMetrics.QualityScore, 1))) is below recommended threshold of 70")
        }
        
        if ($qualityMetrics.SearchReadinessPercentage -lt 90) {
            $qualityTest.Recommendations.Add("Search readiness ($([Math]::Round($qualityMetrics.SearchReadinessPercentage, 1))%) could be improved")  # Fixed % escaping
        }
        
        if ($qualityMetrics.OptimalSizePercentage -lt 80) {
            $qualityTest.Recommendations.Add("Consider adjusting chunk size parameters - only $([Math]::Round($qualityMetrics.OptimalSizePercentage, 1))% of documents have optimal chunk sizes")
        }
        
        return $qualityTest
        
    } catch {
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            QualityMetrics = @{}
        }
    }
}

function Test-SearchIntegration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$true)]
        [array]$SampleDocuments
    )
    
    try {
        $searchTest = @{
            Status = "Unknown"
            ConnectionTest = $false
            IndexExists = $false
            SearchQueries = [List[hashtable]]::new()
            Recommendations = [List[string]]::new()
        }
        
        # Test Azure Search connection
        $searchUrl = "$($Configuration.AzureSearch.ServiceUrl)/servicestats?api-version=$($Configuration.AzureSearch.ApiVersion)"
        $headers = @{ 'api-key' = $Configuration.AzureSearch.ApiKey; 'Content-Type' = 'application/json' }
        
        try {
            $connectionResponse = Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method GET -TimeoutSec 15
            $searchTest.ConnectionTest = $true
        } catch {
            $searchTest.Recommendations.Add("Azure Search connection failed: $($_.Exception.Message)")
        }
        
        # Test index existence
        if ($searchTest.ConnectionTest) {
            try {
                $indexUrl = "$($Configuration.AzureSearch.ServiceUrl)/indexes/$($Configuration.AzureSearch.IndexName)?api-version=$($Configuration.AzureSearch.ApiVersion)"
                $indexResponse = Invoke-RestMethod -Uri $indexUrl -Headers $headers -Method GET -TimeoutSec 15
                $searchTest.IndexExists = $true
            } catch {
                $searchTest.Recommendations.Add("Target index '$($Configuration.AzureSearch.IndexName)' does not exist or is not accessible")
            }
        }
        
        # Perform sample searches if index exists
        if ($searchTest.IndexExists) {
            $testQueries = @("email", "test", "message", "document")
            
            foreach ($query in $testQueries) {
                try {
                    $searchUrl = "$($Configuration.AzureSearch.ServiceUrl)/indexes/$($Configuration.AzureSearch.IndexName)/docs?api-version=$($Configuration.AzureSearch.ApiVersion)&search=$query&`$top=5"
                    $queryStart = Get-Date
                    $searchResponse = Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method GET -TimeoutSec 10
                    $queryEnd = Get-Date
                    
                    $searchTest.SearchQueries.Add(@{
                        Query = $query
                        ResultCount = $searchResponse.value.Count
                        ResponseTime = ($queryEnd - $queryStart).TotalMilliseconds
                        Success = $true
                    })
                } catch {
                    $searchTest.SearchQueries.Add(@{
                        Query = $query
                        ResultCount = 0
                        ResponseTime = 0
                        Success = $false
                        Error = $_.Exception.Message
                    })
                }
            }
        }
        
        # Calculate overall status
        $successfulQueries = ($searchTest.SearchQueries | Where-Object { $_.Success }).Count
        $searchTest.Status = ($searchTest.ConnectionTest -and $successfulQueries -gt 0) ? "Success" : "Failed"
        
        return $searchTest
        
    } catch {
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            ConnectionTest = $false
            IndexExists = $false
        }
    }
}

function Test-PerformanceBenchmarks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$ProcessingResults,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration
    )
    
    try {
        $performanceMetrics = @{
            ThroughputFilesPerSecond = $ProcessingResults.FilesPerSecond ?? 0
            AverageProcessingTimePerFile = $ProcessingResults.AverageProcessingTime ?? 0
            TotalProcessingTime = $ProcessingResults.ProcessingTime ?? 0
            MemoryUsageEstimate = 0
            ConcurrencyEfficiency = 0
        }
        
        # Estimate memory usage based on document sizes
        $totalDocumentSize = ($ProcessingResults.SampleDocuments | Where-Object { $_.Status -eq "Success" } | 
                             Measure-Object -Property DocumentSize -Sum).Sum ?? 0
        $performanceMetrics.MemoryUsageEstimate = [Math]::Max(100, $totalDocumentSize / 1MB * 2) # Rough estimate
        
        # Calculate concurrency efficiency (compare sequential vs parallel estimated time)
        $estimatedSequentialTime = $ProcessingResults.FilesProcessed * $performanceMetrics.AverageProcessingTimePerFile
        $performanceMetrics.ConcurrencyEfficiency = $estimatedSequentialTime -gt 0 ? 
            ($estimatedSequentialTime / [Math]::Max($performanceMetrics.TotalProcessingTime, 1)) : 1
        
        $performanceTest = @{
            Status = "Success"
            Metrics = $performanceMetrics
            Benchmarks = @{
                ThroughputRating = Get-PerformanceRating -Value $performanceMetrics.ThroughputFilesPerSecond -Type "Throughput"
                ResponseTimeRating = Get-PerformanceRating -Value $performanceMetrics.AverageProcessingTimePerFile -Type "ResponseTime"
                ResourceUtilization = Get-ResourceUtilizationRating -MemoryMB $performanceMetrics.MemoryUsageEstimate
            }
            Recommendations = [List[string]]::new()
        }
        
        # Generate performance recommendations
        if ($performanceMetrics.ThroughputFilesPerSecond -lt 1.0) {
            $performanceTest.Recommendations.Add("Low throughput detected ($([Math]::Round($performanceMetrics.ThroughputFilesPerSecond, 2)) files/sec) - consider optimizing processing pipeline")
        }
        
        if ($performanceMetrics.AverageProcessingTimePerFile -gt 10) {
            $performanceTest.Recommendations.Add("High average processing time ($([Math]::Round($performanceMetrics.AverageProcessingTimePerFile, 2)) seconds) - consider performance optimization")
        }
        
        if ($performanceMetrics.MemoryUsageEstimate -gt 1024) {
            $performanceTest.Recommendations.Add("High memory usage detected ($([Math]::Round($performanceMetrics.MemoryUsageEstimate, 0)) MB) - consider batch size optimization")  # Fixed string interpolation
        }
        
        if ($performanceMetrics.ConcurrencyEfficiency -lt 2.0) {
            $performanceTest.Recommendations.Add("Low concurrency efficiency detected - parallel processing may not be optimal for this workload")
        }
        
        return $performanceTest
        
    } catch {
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            Metrics = @{}
        }
    }
}

function Get-PerformanceRating {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [double]$Value,
        
        [Parameter(Mandatory=$true)]
        [string]$Type
    )
    
    switch ($Type) {
        "Throughput" {
            return $Value -ge 2.0 ? "Excellent" : 
                   $Value -ge 1.0 ? "Good" : 
                   $Value -ge 0.5 ? "Fair" : "Poor"
        }
        "ResponseTime" {
            return $Value -le 2.0 ? "Excellent" : 
                   $Value -le 5.0 ? "Good" : 
                   $Value -le 10.0 ? "Fair" : "Poor"
        }
        default { return "Unknown" }
    }
}

function Get-ResourceUtilizationRating {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [double]$MemoryMB
    )
    
    return $MemoryMB -le 256 ? "Low" : 
           $MemoryMB -le 512 ? "Moderate" : 
           $MemoryMB -le 1024 ? "High" : "Very High"
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
    <title>Email RAG Pipeline Test Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 20px; }
        .header h1 { margin: 0; font-size: 2.5em; }
        .header p { margin: 5px 0; opacity: 0.9; }
        .section { margin: 20px 0; padding: 20px; border: 1px solid #ddd; border-radius: 8px; }
        .success { background-color: #d4edda; border-color: #c3e6cb; }
        .warning { background-color: #fff3cd; border-color: #ffeaa7; }
        .error { background-color: #f8d7da; border-color: #f5c6cb; }
        .info { background-color: #d1ecf1; border-color: #bee5eb; }
        .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .metric { background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center; border-left: 4px solid #007bff; }
        .metric-value { font-size: 2em; font-weight: bold; color: #495057; }
        .metric-label { font-size: 0.9em; color: #6c757d; margin-top: 5px; }
        .recommendation { background: #e3f2fd; padding: 15px; margin: 10px 0; border-left: 4px solid #2196f3; border-radius: 4px; }
        .test-phase { border-left: 4px solid #28a745; padding-left: 15px; margin: 15px 0; }
        .status-badge { padding: 5px 12px; border-radius: 15px; color: white; font-weight: bold; display: inline-block; }
        .status-success { background-color: #28a745; }
        .status-warning { background-color: #ffc107; color: #212529; }
        .status-error { background-color: #dc3545; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f8f9fa; font-weight: 600; }
        .progress-bar { width: 100%; height: 20px; background-color: #e9ecef; border-radius: 10px; overflow: hidden; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #28a745, #20c997); transition: width 0.3s ease; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📧 Email RAG Pipeline Test Report</h1>
            <p><strong>Configuration:</strong> $($TestResults.Configuration)</p>
            <p><strong>Generated:</strong> $($TestResults.EndTime.ToString("yyyy-MM-dd HH:mm:ss"))</p>
            <p><strong>Total Duration:</strong> $([Math]::Round($TestResults.TotalDuration, 1)) seconds</p>
        </div>

        <div class="section success">
            <h2>📊 Executive Summary</h2>
            <div class="metric-grid">
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
        </div>
"@

        # Add test phase sections
        foreach ($phaseName in $TestResults.Results.Keys) {
            $phaseResult = $TestResults.Results[$phaseName]
            $sectionClass = switch ($phaseResult.Status) {
                "Success" { "success" }
                "Passed" { "success" }
                { $_ -in @("Partial", "Warning") } { "warning" }
                default { "error" }
            }
            
            $statusBadgeClass = switch ($phaseResult.Status) {
                "Success" { "status-success" }
                "Passed" { "status-success" }
                { $_ -in @("Partial", "Warning") } { "status-warning" }
                default { "status-error" }
            }
            
            $reportHtml += @"
        <div class="section $sectionClass">
            <div class="test-phase">
                <h2>$phaseName Test Results <span class="status-badge $statusBadgeClass">$($phaseResult.Status)</span></h2>
"@
            
            # Add phase-specific details
            if ($phaseName -eq "Processing" -and $phaseResult.FilesProcessed) {
                $reportHtml += @"
                <div class="metric-grid">
                    <div class="metric">
                        <div class="metric-value">$($phaseResult.FilesProcessed)</div>
                        <div class="metric-label">Files Processed</div>
                    </div>
                    <div class="metric">
                        <div class="metric-value">$($phaseResult.SuccessfulFiles)</div>
                        <div class="metric-label">Successful</div>
                    </div>
                    <div class="metric">
                        <div class="metric-value">$([Math]::Round($phaseResult.ProcessingTime, 1))s</div>
                        <div class="metric-label">Processing Time</div>
                    </div>
                    <div class="metric">
                        <div class="metric-value">$($phaseResult.TotalChunks)</div>
                        <div class="metric-label">Total Chunks</div>
                    </div>
                </div>
"@
            }
            
            if ($phaseName -eq "Quality" -and $phaseResult.QualityMetrics) {
                $reportHtml += @"
                <div class="metric-grid">
                    <div class="metric">
                        <div class="metric-value">$([Math]::Round($phaseResult.QualityMetrics.AverageContentLength / 1KB, 1))KB</div>
                        <div class="metric-label">Avg Content Size</div>
                    </div>
                    <div class="metric">
                        <div class="metric-value">$([Math]::Round($phaseResult.QualityMetrics.QualityScore, 1))</div>
                        <div class="metric-label">Quality Score</div>
                    </div>
                    <div class="metric">
                        <div class="metric-value">$([Math]::Round($phaseResult.QualityMetrics.OptimalSizePercentage, 1))%</div>
                        <div class="metric-label">Optimal Size</div>
                    </div>
                </div>
"@
            }
            
            # Add recommendations for this phase
            if ($phaseResult.Recommendations -and $phaseResult.Recommendations.Count -gt 0) {
                $reportHtml += "<h3>💡 Recommendations</h3>"
                foreach ($recommendation in $phaseResult.Recommendations) {
                    $reportHtml += "<div class='recommendation'>$recommendation</div>"
                }
            }
            
            $reportHtml += @"
            </div>
        </div>
"@
        }
        
        # Add overall recommendations
        if ($TestResults.Recommendations -and $TestResults.Recommendations.Count -gt 0) {
            $reportHtml += @"
        <div class="section info">
            <h2>🎯 Overall Recommendations</h2>
"@
            foreach ($recommendation in $TestResults.Recommendations) {
                $reportHtml += "<div class='recommendation'>$recommendation</div>"
            }
            $reportHtml += "</div>"
        }
        
        $reportHtml += @"
        <div class="section info">
            <h2>ℹ️ Test Environment</h2>
            <table>
                <tr><th>Configuration</th><td>$($TestResults.Configuration)</td></tr>
                <tr><th>Test Data Path</th><td>$($TestResults.TestDataPath)</td></tr>
                <tr><th>Start Time</th><td>$($TestResults.StartTime.ToString("yyyy-MM-dd HH:mm:ss"))</td></tr>
                <tr><th>End Time</th><td>$($TestResults.EndTime.ToString("yyyy-MM-dd HH:mm:ss"))</td></tr>
                <tr><th>Total Duration</th><td>$([Math]::Round($TestResults.TotalDuration, 2)) seconds</td></tr>
                <tr><th>PowerShell Version</th><td>$($PSVersionTable.PSVersion)</td></tr>
            </table>
        </div>
    </div>
</body>
</html>
"@
        
        # Ensure output directory exists
        $outputDir = Split-Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
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
    
    $overallStatus = $failedTests -eq 0 ? 
        ($passedTests -eq $totalTests ? "Passed" : "Warning") : 
        "Failed"
    
    return @{
        TotalTests = $totalTests
        TestsPassed = $passedTests
        TestsFailed = $failedTests
        SuccessRate = $totalTests -gt 0 ? [Math]::Round(($passedTests / $totalTests) * 100, 1) : 0
        OverallStatus = $overallStatus
    }
}

function Generate-TestRecommendations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$TestResults
    )
    
    $recommendations = [List[string]]::new()
    
    # Configuration recommendations
    if ($TestResults.Results.Configuration?.Status -ne "Passed") {
        $recommendations.Add("🔧 Review and fix configuration issues before production deployment")
    }
    
    # Processing recommendations
    if ($TestResults.Results.Processing) {
        $processing = $TestResults.Results.Processing
        if ($processing.AverageProcessingTime -gt 30) {
            $recommendations.Add("⚡ Consider optimizing email processing pipeline - current average time is $([Math]::Round($processing.AverageProcessingTime, 1)) seconds")
        }
        if ($processing.FailedFiles -gt 0) {
            $recommendations.Add("🐛 Investigate and resolve issues causing $($processing.FailedFiles) file processing failures")
        }
    }
    
    # Quality recommendations
    if ($TestResults.Results.Quality?.QualityMetrics) {
        $quality = $TestResults.Results.Quality.QualityMetrics
        if ($quality.OptimalSizePercentage -lt 70) {
            $recommendations.Add("📏 Adjust chunk size parameters to improve optimal size percentage (currently $([Math]::Round($quality.OptimalSizePercentage, 1))%)")
        }
        if ($quality.QualityScore -lt 70) {
            $recommendations.Add("✨ Improve content quality preprocessing to increase average quality score (currently $([Math]::Round($quality.QualityScore, 1)))")
        }
    }
    
    # Performance recommendations
    if ($TestResults.Results.Performance?.Metrics) {
        $perf = $TestResults.Results.Performance.Metrics
        if ($perf.ThroughputFilesPerSecond -lt 1.0) {
            $recommendations.Add("🚀 Optimize processing throughput - current rate is $([Math]::Round($perf.ThroughputFilesPerSecond, 2)) files/second")
        }
    }
    
    return $recommendations.ToArray()
}

function Display-TestSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$TestResults
    )
    
    Write-Host "`n" + "="*60 -ForegroundColor Green
    Write-Host "📧 EMAIL RAG PIPELINE TEST SUMMARY" -ForegroundColor Green
    Write-Host "="*60 -ForegroundColor Green
    
    $statusColor = $TestResults.Summary.OverallStatus -eq "Passed" ? "Green" : 
                   $TestResults.Summary.OverallStatus -eq "Warning" ? "Yellow" : "Red"
    
    Write-Host "🎯 Overall Status: $($TestResults.Summary.OverallStatus)" -ForegroundColor $statusColor
    Write-Host "📊 Tests Passed: $($TestResults.Summary.TestsPassed)/$($TestResults.Summary.TotalTests) ($($TestResults.Summary.SuccessRate)%)" -ForegroundColor Cyan
    Write-Host "⏱️  Total Duration: $([Math]::Round($TestResults.TotalDuration, 1)) seconds" -ForegroundColor Cyan
    
    if ($TestResults.ReportPath) {
        Write-Host "📄 Detailed Report: $($TestResults.ReportPath)" -ForegroundColor Magenta
    }
    
    if ($TestResults.Recommendations -and $TestResults.Recommendations.Count -gt 0) {
        Write-Host "`n💡 Top Recommendations:" -ForegroundColor Yellow
        for ($i = 0; $i -lt [Math]::Min(3, $TestResults.Recommendations.Count); $i++) {
            Write-Host "  $($i + 1). $($TestResults.Recommendations[$i])" -ForegroundColor White
        }
    }
    
    Write-Host "`n" + "="*60 -ForegroundColor Green
}

function Create-TestDataset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [int]$EmailCount = 10
    )
    
    try {
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        Write-Host "🏗️  Creating test dataset with $EmailCount emails..." -ForegroundColor Cyan
        
        $testEmails = @()
        for ($i = 1; $i -le $EmailCount; $i++) {
            $testEmails += @{
                Subject = "Test Email $i - Sample Content"
                Sender = "testuser$i@example.com"
                Body = "This is a test email message number $i. It contains sample content for testing the RAG processing pipeline. The email includes various elements like dates, names, and technical terms to test extraction capabilities."
                SentDate = (Get-Date).AddDays(-$i)
            }
        }
        
        # Create mock MSG files (as text files for testing)
        foreach ($email in $testEmails) {
            $fileName = "test_email_$($testEmails.IndexOf($email) + 1).msg"
            $filePath = Join-Path $OutputPath $fileName
            
            $mockMsgContent = @"
Subject: $($email.Subject)
From: $($email.Sender)
Date: $($email.SentDate.ToString("yyyy-MM-dd HH:mm:ss"))

$($email.Body)
"@
            
            $mockMsgContent | Out-File -FilePath $filePath -Encoding UTF8
        }
        
        Write-Host "✅ Created $($testEmails.Count) test email files in: $OutputPath" -ForegroundColor Green
        
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

function Invoke-ComprehensiveRAGTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$false)]
        [string]$TestDataPath,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath = "./test-results",
        
        [Parameter(Mandatory=$false)]
        [int]$SampleSize = 10
    )
    
    try {
        Write-Host "🚀 Starting Comprehensive RAG System Test" -ForegroundColor Green
        
        # Create test data if not provided
        if ([string]::IsNullOrEmpty($TestDataPath)) {
            $testDataPath = Join-Path $OutputPath "test-data"
            $datasetResult = Create-TestDataset -OutputPath $testDataPath -EmailCount $SampleSize
            $TestDataPath = $datasetResult.OutputPath
        }
        
        # Run comprehensive pipeline test
        $testResults = Test-EmailProcessingPipeline -TestDataPath $TestDataPath -Configuration $Configuration -SampleSize $SampleSize -OutputPath $OutputPath
        
        return $testResults
        
    } catch {
        Write-Host "❌ Comprehensive RAG test failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

Write-Verbose "RAGTestFramework_v2 (Fixed) module loaded successfully"