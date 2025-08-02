# EmailRAGProcessor_v2.psm1 - Core Email Processing Pipeline for RAG System
# Enhanced for PowerShell 7 with real MSG processing and Azure AI Search integration

using namespace System.Collections.Generic
using namespace System.Collections.Concurrent

Export-ModuleMember -Function @(
    'Invoke-EmailRAGProcessing',
    'Process-EmailBatch',
    'Process-MSGFile', 
    'Convert-EmailToRAGDocument',
    'Test-RAGPipeline',
    'Get-ProcessingStatistics',
    'Start-BackgroundProcessing'
)

function Invoke-EmailRAGProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressCallback,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$StatusCallback,
        
        [Parameter(Mandatory=$false)]
        [switch]$Parallel = $true,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxConcurrency = 5
    )
    
    try {
        Write-Host "🚀 Starting Email RAG Processing Pipeline" -ForegroundColor Green
        $startTime = Get-Date
        
        # Initialize statistics
        $stats = @{
            StartTime = $startTime
            TotalFiles = 0
            ProcessedFiles = 0
            SuccessfulFiles = 0
            FailedFiles = 0
            TotalChunks = 0
            ProcessedChunks = 0
            IndexedDocuments = 0
            Errors = [List[string]]::new()
        }
        
        # Discover MSG files
        Write-Host "📂 Discovering MSG files in: $InputPath" -ForegroundColor Cyan
        $StatusCallback?.Invoke("Discovering files...")
        
        $msgFiles = Get-ChildItem -Path $InputPath -Filter "*.msg" -Recurse -ErrorAction SilentlyContinue
        $stats.TotalFiles = $msgFiles.Count
        
        if ($stats.TotalFiles -eq 0) {
            throw "No MSG files found in the specified path: $InputPath"
        }
        
        Write-Host "📧 Found $($stats.TotalFiles) MSG files to process" -ForegroundColor Yellow
        $ProgressCallback?.Invoke(@{ Total = $stats.TotalFiles; Current = 0; Status = "Starting processing..." })
        
        # Process files based on parallel flag
        if ($Parallel -and $stats.TotalFiles -gt 1) {
            Write-Host "⚡ Processing files in parallel (Max Concurrency: $MaxConcurrency)" -ForegroundColor Magenta
            $processedResults = Process-FilesParallel -Files $msgFiles -Configuration $Configuration -MaxConcurrency $MaxConcurrency -Stats $stats -ProgressCallback $ProgressCallback -StatusCallback $StatusCallback
        } else {
            Write-Host "⏳ Processing files sequentially" -ForegroundColor Yellow
            $processedResults = Process-FilesSequential -Files $msgFiles -Configuration $Configuration -Stats $stats -ProgressCallback $ProgressCallback -StatusCallback $StatusCallback
        }
        
        # Calculate final statistics
        $stats.EndTime = Get-Date
        $stats.TotalDuration = ($stats.EndTime - $stats.StartTime).TotalSeconds
        $stats.AverageProcessingTime = $stats.TotalDuration / [Math]::Max($stats.ProcessedFiles, 1)
        $stats.SuccessRate = ($stats.SuccessfulFiles / [Math]::Max($stats.TotalFiles, 1)) * 100
        
        # Final status update
        $StatusCallback?.Invoke("Processing completed!")
        $ProgressCallback?.Invoke(@{ Total = $stats.TotalFiles; Current = $stats.TotalFiles; Status = "Complete!" })
        
        Write-Host "`n✅ Email RAG Processing Completed!" -ForegroundColor Green
        Write-Host "📊 Statistics:" -ForegroundColor Cyan
        Write-Host "   • Total Files: $($stats.TotalFiles)" -ForegroundColor White
        Write-Host "   • Processed Successfully: $($stats.SuccessfulFiles)" -ForegroundColor Green
        Write-Host "   • Failed: $($stats.FailedFiles)" -ForegroundColor Red
        Write-Host "   • Total Chunks Generated: $($stats.TotalChunks)" -ForegroundColor Yellow
        Write-Host "   • Documents Indexed: $($stats.IndexedDocuments)" -ForegroundColor Magenta
        Write-Host "   • Success Rate: $([Math]::Round($stats.SuccessRate, 2))%" -ForegroundColor Cyan
        Write-Host "   • Total Duration: $([Math]::Round($stats.TotalDuration, 2)) seconds" -ForegroundColor Gray
        Write-Host "   • Average Processing Time: $([Math]::Round($stats.AverageProcessingTime, 2)) seconds/file" -ForegroundColor Gray
        
        return @{
            Status = "Success"
            Statistics = $stats
            ProcessedFiles = $processedResults
            Configuration = $Configuration
        }
        
    } catch {
        $errorMessage = "Email RAG processing failed: $($_.Exception.Message)"
        Write-Host "❌ $errorMessage" -ForegroundColor Red
        $StatusCallback?.Invoke("Processing failed: $($_.Exception.Message)")
        
        $stats.Errors.Add($errorMessage)
        $stats.EndTime = Get-Date
        $stats.TotalDuration = ($stats.EndTime - $stats.StartTime).TotalSeconds
        
        return @{
            Status = "Failed"
            Error = $errorMessage
            Statistics = $stats
            Configuration = $Configuration
        }
    }
}

function Process-FilesParallel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Files,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$true)]
        [int]$MaxConcurrency,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Stats,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressCallback,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$StatusCallback
    )
    
    $processedFiles = [ConcurrentBag[hashtable]]::new()
    $processedCount = 0
    
    $Files | ForEach-Object -Parallel {
        $file = $_
        $config = $using:Configuration
        $stats = $using:Stats
        $progressCallback = $using:ProgressCallback
        $statusCallback = $using:StatusCallback
        $processedFiles = $using:processedFiles
        
        try {
            # Process single MSG file
            $result = Process-MSGFile -FilePath $file.FullName -Configuration $config
            $processedFiles.Add($result)
            
            # Thread-safe counter update
            $currentCount = [System.Threading.Interlocked]::Increment([ref]$using:processedCount)
            
            if ($result.Status -eq "Success") {
                [System.Threading.Interlocked]::Increment([ref]$stats.SuccessfulFiles)
                [System.Threading.Interlocked]::Add([ref]$stats.TotalChunks, $result.ChunkCount)
                $result.Indexed && [System.Threading.Interlocked]::Increment([ref]$stats.IndexedDocuments)
            } else {
                [System.Threading.Interlocked]::Increment([ref]$stats.FailedFiles)
                $stats.Errors.Add("Failed to process $($file.Name): $($result.Error)")
            }
            
            [System.Threading.Interlocked]::Increment([ref]$stats.ProcessedFiles)
            
            # Update progress
            $progressCallback?.Invoke(@{ 
                Total = $stats.TotalFiles; 
                Current = $currentCount; 
                Status = "Processed: $($file.Name)" 
            })
            
        } catch {
            [System.Threading.Interlocked]::Increment([ref]$stats.FailedFiles)
            [System.Threading.Interlocked]::Increment([ref]$stats.ProcessedFiles)
            $stats.Errors.Add("Exception processing $($file.Name): $($_.Exception.Message)")
        }
        
    } -ThrottleLimit $MaxConcurrency
    
    return $processedFiles.ToArray()
}

function Process-FilesSequential {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Files,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Stats,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressCallback,
        
        [Parameter(Mandatory=$false)]
        [scriptblock]$StatusCallback
    )
    
    $processedFiles = [List[hashtable]]::new()
    
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $file = $Files[$i]
        
        try {
            $StatusCallback?.Invoke("Processing: $($file.Name)")
            
            # Process single MSG file
            $result = Process-MSGFile -FilePath $file.FullName -Configuration $Configuration
            $processedFiles.Add($result)
            
            if ($result.Status -eq "Success") {
                $Stats.SuccessfulFiles++
                $Stats.TotalChunks += $result.ChunkCount
                $result.Indexed && $Stats.IndexedDocuments++
            } else {
                $Stats.FailedFiles++
                $Stats.Errors.Add("Failed to process $($file.Name): $($result.Error)")
            }
            
            $Stats.ProcessedFiles++
            
            # Update progress
            $ProgressCallback?.Invoke(@{ 
                Total = $Stats.TotalFiles; 
                Current = $i + 1; 
                Status = "Processed: $($file.Name)" 
            })
            
        } catch {
            $Stats.FailedFiles++
            $Stats.ProcessedFiles++
            $Stats.Errors.Add("Exception processing $($file.Name): $($_.Exception.Message)")
        }
    }
    
    return $processedFiles.ToArray()
}

function Process-MSGFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration
    )
    
    try {
        Write-Verbose "Processing MSG file: $FilePath"
        
        $result = @{
            FilePath = $FilePath
            FileName = Split-Path $FilePath -Leaf
            Status = "Unknown"
            ProcessingTime = 0
            ChunkCount = 0
            DocumentSize = 0
            Indexed = $false
            Error = $null
            StartTime = Get-Date
        }
        
        # Check if file exists and is accessible
        if (-not (Test-Path $FilePath)) {
            throw "File not found: $FilePath"
        }
        
        $fileInfo = Get-Item $FilePath
        $result.DocumentSize = $fileInfo.Length
        
        # Extract email content using Outlook COM or alternative method
        $emailContent = Extract-EmailContent -FilePath $FilePath
        
        if ([string]::IsNullOrWhiteSpace($emailContent.Body)) {
            throw "No content extracted from MSG file"
        }
        
        # Convert to RAG document format
        $ragDocument = Convert-EmailToRAGDocument -EmailContent $emailContent -Configuration $Configuration
        
        # Generate chunks for RAG processing
        if ($Configuration.Processing?.Chunking?.Enabled -ne $false) {
            $chunks = Split-EmailContent -Content $ragDocument -ChunkingConfig $Configuration.Processing.Chunking
            $result.ChunkCount = $chunks.Count
        } else {
            $chunks = @($ragDocument)
            $result.ChunkCount = 1
        }
        
        # Index to Azure AI Search if configured
        if ($Configuration.AzureSearch -and -not [string]::IsNullOrEmpty($Configuration.AzureSearch.ServiceName)) {
            try {
                $indexResult = Index-EmailDocument -Document $ragDocument -Chunks $chunks -AzureConfig $Configuration.AzureSearch
                $result.Indexed = $indexResult.Success
            } catch {
                Write-Warning "Failed to index document: $($_.Exception.Message)"
                # Continue processing even if indexing fails
            }
        }
        
        $result.EndTime = Get-Date
        $result.ProcessingTime = ($result.EndTime - $result.StartTime).TotalSeconds
        $result.Status = "Success"
        
        Write-Verbose "Successfully processed: $($result.FileName) ($($result.ChunkCount) chunks, $([Math]::Round($result.ProcessingTime, 2))s)"
        
        return $result
        
    } catch {
        $result.Status = "Failed"
        $result.Error = $_.Exception.Message
        $result.EndTime = Get-Date
        $result.ProcessingTime = ($result.EndTime - $result.StartTime).TotalSeconds
        
        Write-Warning "Failed to process MSG file $FilePath`: $($_.Exception.Message)"
        return $result
    }
}

function Extract-EmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        # Try to use Outlook COM object first
        if (Get-Command -Name "New-Object" -ParameterName "ComObject" -ErrorAction SilentlyContinue) {
            try {
                $outlook = New-Object -ComObject Outlook.Application
                $msg = $outlook.Session.OpenSharedItem($FilePath)
                
                $emailContent = @{
                    Subject = $msg.Subject ?? "No Subject"
                    Body = $msg.Body ?? ""
                    Sender = $msg.SenderName ?? "Unknown Sender"
                    SentOn = $msg.SentOn ?? (Get-Date)
                    Recipients = ($msg.Recipients | ForEach-Object { $_.Name }) -join "; "
                    HasAttachments = $msg.Attachments.Count -gt 0
                    Importance = $msg.Importance ?? 1
                    Size = $msg.Size ?? 0
                }
                
                # Clean up COM objects
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
                
                return $emailContent
                
            } catch {
                Write-Warning "Outlook COM failed, falling back to file-based extraction: $($_.Exception.Message)"
            }
        }
        
        # Fallback: Extract basic information from file properties and content
        $fileInfo = Get-Item $FilePath
        $emailContent = @{
            Subject = $fileInfo.BaseName
            Body = "Email content from: $($fileInfo.Name) (Size: $($fileInfo.Length) bytes, Modified: $($fileInfo.LastWriteTime))"
            Sender = "Unknown (extracted from file)"
            SentOn = $fileInfo.LastWriteTime
            Recipients = "Unknown"
            HasAttachments = $false
            Importance = 1
            Size = $fileInfo.Length
        }
        
        # Try to read some content if it's text-readable
        try {
            $rawContent = Get-Content $FilePath -Encoding UTF8 -TotalCount 50 -ErrorAction SilentlyContinue | Where-Object { $_ -match "[a-zA-Z]" }
            if ($rawContent) {
                $emailContent.Body = ($rawContent -join " ").Substring(0, [Math]::Min(1000, ($rawContent -join " ").Length))
            }
        } catch {
            # Keep default body if reading fails
        }
        
        return $emailContent
        
    } catch {
        throw "Failed to extract email content: $($_.Exception.Message)"
    }
}

function Convert-EmailToRAGDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailContent,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration
    )
    
    # Create RAG-optimized document structure
    $ragDocument = @{
        id = [Guid]::NewGuid().ToString()
        title = $EmailContent.Subject
        content = $EmailContent.Body
        sender_name = $EmailContent.Sender
        sent_date = $EmailContent.SentOn.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        recipients = $EmailContent.Recipients
        has_attachments = $EmailContent.HasAttachments
        importance = $EmailContent.Importance
        document_type = "email"
        processed_date = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        content_length = $EmailContent.Body.Length
        metadata = @{
            original_size = $EmailContent.Size
            processing_version = "2.0"
            extraction_method = "outlook_com"
        }
    }
    
    # Apply content cleaning if configured
    if ($Configuration.Processing?.ContentCleaning?.Enabled -ne $false) {
        $ragDocument.content = Clean-EmailContent -Content $ragDocument.content -CleaningConfig $Configuration.Processing.ContentCleaning
    }
    
    return $ragDocument
}

function Clean-EmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$CleaningConfig = @{}
    )
    
    $cleanedContent = $Content
    
    # Remove email signatures if configured
    if ($CleaningConfig.RemoveSignatures -eq $true) {
        $cleanedContent = $cleanedContent -replace "(?s)--\s*\r?\n.*$", ""
        $cleanedContent = $cleanedContent -replace "(?s)Best regards.*$", ""
        $cleanedContent = $cleanedContent -replace "(?s)Sincerely.*$", ""
    }
    
    # Remove quoted text if configured
    if ($CleaningConfig.RemoveQuotedText -eq $true) {
        $cleanedContent = $cleanedContent -replace "(?m)^>.*$", ""
        $cleanedContent = $cleanedContent -replace "(?s)-----Original Message-----.*$", ""
    }
    
    # Normalize whitespace
    if ($CleaningConfig.NormalizeWhitespace -ne $false) {
        $cleanedContent = $cleanedContent -replace "\s+", " "
        $cleanedContent = $cleanedContent.Trim()
    }
    
    return $cleanedContent
}

function Split-EmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Content,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$ChunkingConfig = @{}
    )
    
    $targetTokens = $ChunkingConfig.TargetTokens ?? 384
    $maxTokens = $ChunkingConfig.MaxTokens ?? 512
    $overlapTokens = $ChunkingConfig.OverlapTokens ?? 32
    
    # Simple chunking implementation (can be enhanced with more sophisticated algorithms)
    $text = $Content.content
    $chunks = [List[hashtable]]::new()
    
    if ($text.Length -le $targetTokens * 4) { # Rough estimate: 4 chars per token
        # Content is small enough for single chunk
        $chunks.Add($Content)
    } else {
        # Split into multiple chunks
        $chunkSize = $targetTokens * 4
        $overlapSize = $overlapTokens * 4
        $position = 0
        $chunkIndex = 0
        
        while ($position -lt $text.Length) {
            $endPosition = [Math]::Min($position + $chunkSize, $text.Length)
            $chunkText = $text.Substring($position, $endPosition - $position)
            
            $chunk = $Content.Clone()
            $chunk.id = "$($Content.id)_chunk_$chunkIndex"
            $chunk.content = $chunkText
            $chunk.chunk_index = $chunkIndex
            $chunk.chunk_total = -1 # Will be updated after all chunks are created
            $chunk.content_length = $chunkText.Length
            
            $chunks.Add($chunk)
            
            $position += ($chunkSize - $overlapSize)
            $chunkIndex++
        }
        
        # Update total chunk count
        foreach ($chunk in $chunks) {
            $chunk.chunk_total = $chunks.Count
        }
    }
    
    return $chunks.ToArray()
}

function Index-EmailDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Document,
        
        [Parameter(Mandatory=$true)]
        [array]$Chunks,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$AzureConfig
    )
    
    try {
        $indexUrl = "$($AzureConfig.ServiceUrl)/indexes/$($AzureConfig.IndexName)/docs/index?api-version=$($AzureConfig.ApiVersion)"
        $headers = @{
            'api-key' = $AzureConfig.ApiKey
            'Content-Type' = 'application/json'
        }
        
        # Prepare documents for indexing
        $indexDocuments = @{
            value = $Chunks | ForEach-Object {
                @{
                    '@search.action' = 'upload'
                    id = $_.id
                    title = $_.title
                    content = $_.content
                    sender_name = $_.sender_name
                    sent_date = $_.sent_date
                    recipients = $_.recipients
                    has_attachments = $_.has_attachments
                    importance = $_.importance
                    document_type = $_.document_type
                    processed_date = $_.processed_date
                    content_length = $_.content_length
                }
            }
        }
        
        $body = $indexDocuments | ConvertTo-Json -Depth 10
        $response = Invoke-RestMethod -Uri $indexUrl -Headers $headers -Method POST -Body $body -TimeoutSec 30
        
        Write-Verbose "Successfully indexed $($Chunks.Count) chunks to Azure AI Search"
        
        return @{
            Success = $true
            IndexedDocuments = $Chunks.Count
            Response = $response
        }
        
    } catch {
        Write-Warning "Failed to index to Azure AI Search: $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

function Test-RAGPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PipelineConfig,
        
        [Parameter(Mandatory=$false)]
        [string]$TestQuery = "test email search"
    )
    
    try {
        Write-Host "🧪 Testing RAG Pipeline Configuration" -ForegroundColor Cyan
        
        $testResults = @{
            Status = "Unknown"
            TestsRun = 0
            TestsPassed = 0
            TestsFailed = 0
            Details = @{}
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
        # Test 1: Configuration validation
        Write-Host "1. Validating configuration schema..." -ForegroundColor Yellow
        $testResults.TestsRun++
        
        try {
            Import-Module "$PSScriptRoot\RAGConfigManager_v2_Fixed.psm1" -Force
            $configTest = Test-RAGConfiguration -Configuration $PipelineConfig
            $testResults.TestsPassed++
            $testResults.Details.ConfigurationTest = $configTest
            Write-Host "   ✅ Configuration validation passed" -ForegroundColor Green
        } catch {
            $testResults.TestsFailed++
            $testResults.Details.ConfigurationTest = @{ Status = "Failed"; Error = $_.Exception.Message }
            Write-Host "   ❌ Configuration validation failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Test 2: Azure Search connectivity
        if ($PipelineConfig.AzureSearch) {
            Write-Host "2. Testing Azure Search connectivity..." -ForegroundColor Yellow
            $testResults.TestsRun++
            
            try {
                $searchUrl = "$($PipelineConfig.AzureSearch.ServiceUrl)/servicestats?api-version=$($PipelineConfig.AzureSearch.ApiVersion)"
                $headers = @{ 'api-key' = $PipelineConfig.AzureSearch.ApiKey }
                $response = Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method GET -TimeoutSec 15
                
                $testResults.TestsPassed++
                $testResults.Details.AzureSearchTest = @{ Status = "Passed"; Response = $response }
                Write-Host "   ✅ Azure Search connectivity successful" -ForegroundColor Green
            } catch {
                $testResults.TestsFailed++
                $testResults.Details.AzureSearchTest = @{ Status = "Failed"; Error = $_.Exception.Message }
                Write-Host "   ❌ Azure Search connectivity failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Calculate overall status
        $testResults.Status = $testResults.TestsFailed -eq 0 ? "Success" : "Failed"
        
        Write-Host "`n📊 RAG Pipeline Test Results:" -ForegroundColor Cyan
        Write-Host "   • Status: $($testResults.Status)" -ForegroundColor ($testResults.Status -eq "Success" ? "Green" : "Red")
        Write-Host "   • Tests Run: $($testResults.TestsRun)" -ForegroundColor White
        Write-Host "   • Tests Passed: $($testResults.TestsPassed)" -ForegroundColor Green
        Write-Host "   • Tests Failed: $($testResults.TestsFailed)" -ForegroundColor Red
        
        return $testResults
        
    } catch {
        Write-Host "❌ RAG pipeline test failed: $($_.Exception.Message)" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
    }
}

function Get-ProcessingStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$ProcessingResult
    )
    
    $stats = $ProcessingResult.Statistics
    
    return @{
        Summary = @{
            TotalFiles = $stats.TotalFiles
            SuccessfulFiles = $stats.SuccessfulFiles
            FailedFiles = $stats.FailedFiles
            SuccessRate = [Math]::Round(($stats.SuccessfulFiles / [Math]::Max($stats.TotalFiles, 1)) * 100, 2)
            TotalChunks = $stats.TotalChunks
            IndexedDocuments = $stats.IndexedDocuments
        }
        Performance = @{
            TotalDuration = [Math]::Round($stats.TotalDuration, 2)
            AverageProcessingTime = [Math]::Round($stats.AverageProcessingTime, 2)
            FilesPerSecond = [Math]::Round($stats.SuccessfulFiles / [Math]::Max($stats.TotalDuration, 1), 2)
        }
        Timeline = @{
            StartTime = $stats.StartTime
            EndTime = $stats.EndTime
            Duration = "$([Math]::Floor($stats.TotalDuration / 60))m $([Math]::Round($stats.TotalDuration % 60, 0))s"
        }
        Errors = $stats.Errors.ToArray()
    }
}

function Start-BackgroundProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$true)]
        [scriptblock]$CompletionCallback
    )
    
    # Start processing in background job
    $job = Start-Job -ScriptBlock {
        param($InputPath, $Configuration, $ModulePath)
        
        Import-Module $ModulePath -Force
        
        return Invoke-EmailRAGProcessing -InputPath $InputPath -Configuration $Configuration -Parallel $true
        
    } -ArgumentList $InputPath, $Configuration, $PSScriptRoot
    
    # Monitor job and call completion callback when done
    Register-ObjectEvent -InputObject $job -EventName StateChanged -Action {
        if ($Event.Sender.State -eq "Completed") {
            $result = Receive-Job -Job $Event.Sender
            & $using:CompletionCallback $result
            Remove-Job -Job $Event.Sender
        }
    }
    
    return $job
}

Write-Verbose "EmailRAGProcessor_v2 (Fixed) module loaded successfully"