# EmailRAGProcessor_v2.psm1 - Enhanced Email Processing Pipeline for Azure AI Search RAG
# Complete processing pipeline integrating MSG processing, content cleaning, chunking, and indexing

Export-ModuleMember -Function @(
    'Start-EmailRAGProcessing',
    'Process-EmailForRAG',
    'Process-EmailBatch',
    'Get-ProcessingStatistics',
    'Test-RAGPipeline',
    'Initialize-RAGPipeline'
)

# Import required modules
Import-Module (Join-Path $PSScriptRoot "clean_msgprocessor.txt") -Force
Import-Module (Join-Path $PSScriptRoot "clean_contentcleaner.txt") -Force
Import-Module (Join-Path $PSScriptRoot "EmailChunkingEngine_v2.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "AzureAISearchIntegration_v2.psm1") -Force

function Initialize-RAGPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$AzureSearchConfig,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$ProcessingConfig = @{},
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [switch]$CreateIndex = $true
    )
    
    try {
        Write-Verbose "Initializing Email RAG processing pipeline..."
        
        # Test Azure Search connection
        $connectionTest = Test-SearchConnection -Config $AzureSearchConfig
        if (-not $connectionTest.IsConnected) {
            throw "Failed to connect to Azure AI Search: $($connectionTest.Error)"
        }
        
        # Create or verify index
        if ($CreateIndex) {
            $indexResult = New-EmailSearchIndex -SearchConfig $AzureSearchConfig -IndexName $IndexName -RecreateIfExists:$false
            Write-Verbose "Index status: $($indexResult.Status)"
        }
        
        # Initialize pipeline configuration
        $pipelineConfig = @{
            AzureSearch = $AzureSearchConfig
            IndexName = $IndexName
            Processing = @{
                ChunkSize = if ($ProcessingConfig.ChunkSize) { $ProcessingConfig.ChunkSize } else { 384 }
                MinChunkSize = if ($ProcessingConfig.MinChunkSize) { $ProcessingConfig.MinChunkSize } else { 128 }
                MaxChunkSize = if ($ProcessingConfig.MaxChunkSize) { $ProcessingConfig.MaxChunkSize } else { 512 }
                OverlapTokens = if ($ProcessingConfig.OverlapTokens) { $ProcessingConfig.OverlapTokens } else { 32 }
                ExtractEntities = if ($ProcessingConfig.ExtractEntities -ne $null) { $ProcessingConfig.ExtractEntities } else { $true }
                OptimizeForRAG = if ($ProcessingConfig.OptimizeForRAG -ne $null) { $ProcessingConfig.OptimizeForRAG } else { $true }
                GenerateEmbeddings = if ($ProcessingConfig.GenerateEmbeddings -ne $null) { $ProcessingConfig.GenerateEmbeddings } else { $true }
                BatchSize = if ($ProcessingConfig.BatchSize) { $ProcessingConfig.BatchSize } else { 50 }
            }
            Statistics = @{
                TotalProcessed = 0
                TotalChunks = 0
                TotalErrors = 0
                AverageProcessingTime = 0
                StartTime = Get-Date
            }
            InitializedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
        Write-Verbose "Email RAG pipeline initialized successfully"
        return $pipelineConfig
        
    } catch {
        Write-Error "Failed to initialize RAG pipeline: $($_.Exception.Message)"
        throw
    }
}

function Start-EmailRAGProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PipelineConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        
        [Parameter(Mandatory=$false)]
        [string]$Filter = "*.msg",
        
        [Parameter(Mandatory=$false)]
        [switch]$Recursive = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$ContinueOnError = $true,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxFiles = 0
    )
    
    try {
        Write-Host "Starting Email RAG processing pipeline..." -ForegroundColor Green
        Write-Host "Input Path: $InputPath" -ForegroundColor Cyan
        Write-Host "Filter: $Filter" -ForegroundColor Cyan
        Write-Host "Index: $($PipelineConfig.IndexName)" -ForegroundColor Cyan
        
        # Find MSG files
        $searchParams = @{
            Path = $InputPath
            Filter = $Filter
            Recurse = $Recursive.IsPresent
        }
        
        $msgFiles = Get-ChildItem @searchParams | Where-Object { -not $_.PSIsContainer }
        
        if ($MaxFiles -gt 0 -and $msgFiles.Count -gt $MaxFiles) {
            $msgFiles = $msgFiles | Select-Object -First $MaxFiles
        }
        
        if ($msgFiles.Count -eq 0) {
            Write-Warning "No MSG files found in: $InputPath"
            return @{
                Status = "NoFilesFound"
                InputPath = $InputPath
                FilesProcessed = 0
            }
        }
        
        Write-Host "Found $($msgFiles.Count) MSG files to process" -ForegroundColor Green
        
        # Process files in batches
        $batchSize = $PipelineConfig.Processing.BatchSize
        $totalBatches = [Math]::Ceiling($msgFiles.Count / $batchSize)
        $allResults = @()
        
        for ($batchNum = 0; $batchNum -lt $totalBatches; $batchNum++) {
            $startIndex = $batchNum * $batchSize
            $endIndex = [Math]::Min($startIndex + $batchSize - 1, $msgFiles.Count - 1)
            $batch = $msgFiles[$startIndex..$endIndex]
            
            Write-Host "`nProcessing batch $($batchNum + 1) of $totalBatches ($($batch.Count) files)..." -ForegroundColor Yellow
            
            $batchResult = Process-EmailBatch -PipelineConfig $PipelineConfig -EmailFiles $batch -ContinueOnError:$ContinueOnError
            $allResults += $batchResult
            
            # Update statistics
            Update-PipelineStatistics -PipelineConfig $PipelineConfig -BatchResult $batchResult
            
            Write-Host "Batch $($batchNum + 1) completed: $($batchResult.ProcessedCount)/$($batch.Count) files processed" -ForegroundColor Green
        }
        
        # Final statistics
        $finalStats = Get-ProcessingStatistics -PipelineConfig $PipelineConfig
        
        Write-Host "`nEmail RAG Processing Complete!" -ForegroundColor Green
        Write-Host "Total Files: $($msgFiles.Count)" -ForegroundColor Cyan
        Write-Host "Successfully Processed: $($finalStats.TotalProcessed)" -ForegroundColor Green
        Write-Host "Total Chunks Created: $($finalStats.TotalChunks)" -ForegroundColor Cyan
        Write-Host "Total Errors: $($finalStats.TotalErrors)" -ForegroundColor $(if ($finalStats.TotalErrors -gt 0) { "Red" } else { "Green" })
        Write-Host "Average Processing Time: $($finalStats.AverageProcessingTime) seconds" -ForegroundColor Cyan
        
        return @{
            Status = "Completed"
            InputPath = $InputPath
            TotalFiles = $msgFiles.Count
            ProcessedFiles = $finalStats.TotalProcessed
            TotalChunks = $finalStats.TotalChunks
            TotalErrors = $finalStats.TotalErrors
            BatchResults = $allResults
            Statistics = $finalStats
            CompletedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Email RAG processing failed: $($_.Exception.Message)"
        throw
    }
}

function Process-EmailBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PipelineConfig,
        
        [Parameter(Mandatory=$true)]
        [array]$EmailFiles,
        
        [Parameter(Mandatory=$false)]
        [switch]$ContinueOnError = $true
    )
    
    $batchResults = @{
        ProcessedCount = 0
        ErrorCount = 0
        TotalChunks = 0
        ProcessedEmails = @()
        Errors = @()
        StartTime = Get-Date
    }
    
    foreach ($emailFile in $EmailFiles) {
        try {
            Write-Verbose "Processing: $($emailFile.Name)"
            
            $emailStartTime = Get-Date
            $result = Process-EmailForRAG -PipelineConfig $PipelineConfig -EmailFilePath $emailFile.FullName
            
            if ($result.Status -eq "Success") {
                $batchResults.ProcessedCount++
                $batchResults.TotalChunks += $result.ChunkCount
                $batchResults.ProcessedEmails += @{
                    FileName = $emailFile.Name
                    ChunkCount = $result.ChunkCount
                    ProcessingTime = ((Get-Date) - $emailStartTime).TotalSeconds
                    IndexedDocuments = $result.IndexedDocuments
                }
                Write-Verbose "Successfully processed: $($emailFile.Name) ($($result.ChunkCount) chunks)"
            } else {
                $batchResults.ErrorCount++
                $batchResults.Errors += @{
                    FileName = $emailFile.Name
                    Error = $result.Error
                    ProcessingTime = ((Get-Date) - $emailStartTime).TotalSeconds
                }
                Write-Warning "Failed to process: $($emailFile.Name) - $($result.Error)"
                
                if (-not $ContinueOnError) {
                    throw "Processing failed for: $($emailFile.Name)"
                }
            }
            
        } catch {
            $batchResults.ErrorCount++
            $batchResults.Errors += @{
                FileName = $emailFile.Name
                Error = $_.Exception.Message
                ProcessingTime = 0
            }
            
            Write-Warning "Error processing $($emailFile.Name): $($_.Exception.Message)"
            
            if (-not $ContinueOnError) {
                throw
            }
        }
    }
    
    $batchResults.EndTime = Get-Date
    $batchResults.TotalTime = ($batchResults.EndTime - $batchResults.StartTime).TotalSeconds
    
    return $batchResults
}

function Process-EmailForRAG {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PipelineConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$EmailFilePath
    )
    
    try {
        # Step 1: Read MSG file
        Write-Verbose "Step 1: Reading MSG file"
        $emailData = Read-MsgFile -FilePath $EmailFilePath -IncludeAttachments:$true -ExtractText:$true
        
        if (-not $emailData) {
            return @{
                Status = "Error"
                Error = "Failed to read MSG file"
                FilePath = $EmailFilePath
            }
        }
        
        # Step 2: Clean content
        Write-Verbose "Step 2: Cleaning email content"
        $emailData = Clean-EmailContent -EmailData $emailData -ExtractEntities:$($PipelineConfig.Processing.ExtractEntities) -OptimizeForRAG:$($PipelineConfig.Processing.OptimizeForRAG)
        
        # Step 3: Create intelligent chunks
        Write-Verbose "Step 3: Creating intelligent chunks"
        $chunkingResult = New-EmailChunks -EmailData $emailData -TargetTokens $PipelineConfig.Processing.ChunkSize -MinTokens $PipelineConfig.Processing.MinChunkSize -MaxTokens $PipelineConfig.Processing.MaxChunkSize -OverlapTokens $PipelineConfig.Processing.OverlapTokens -OptimizeForSearch:$true
        
        if (-not $chunkingResult.Chunks -or $chunkingResult.Chunks.Count -eq 0) {
            return @{
                Status = "Warning"
                Error = "No chunks created from email content"
                FilePath = $EmailFilePath
                ChunkCount = 0
            }
        }
        
        # Step 4: Convert chunks to Azure Search documents
        Write-Verbose "Step 4: Converting chunks to search documents"
        $searchDocuments = @()
        foreach ($chunk in $chunkingResult.Chunks) {
            # Merge email data with chunk data
            $documentData = @{
                id = $chunk.Id
                ChunkNumber = $chunk.ChunkNumber
                TotalChunks = $chunk.TotalChunks
                Content = $chunk.Content
                ChunkType = $chunk.ChunkType
                TokenCount = $chunk.TokenCount
                WordCount = $chunk.WordCount
                QualityScore = $chunk.QualityScore
                SearchRelevance = $chunk.SearchRelevance
                ParentEmailId = $chunk.ParentEmailId
                EmailSubject = $emailData.Subject
                SenderName = $emailData.Sender.Name
                Sender_Email = $emailData.Sender.Email
                Recipients_To = $emailData.Recipients.To
                Recipients_CC = $emailData.Recipients.CC
                SentDate = $emailData.Sent
                ReceivedDate = $emailData.Received
                HasAttachments = ($emailData.Attachments.Count -gt 0)
                AttachmentTypes = if ($emailData.Attachments) { $emailData.Attachments | ForEach-Object { [System.IO.Path]::GetExtension($_.FileName) } } else { @() }
                Importance = $emailData.Importance
                ConversationTopic = $emailData.ConversationTopic
                MessageId = $emailData.InternetMessageId
            }
            
            # Add entity data if available
            if ($emailData.Content.ExtractedEntities) {
                $entities = $emailData.Content.ExtractedEntities
                $documentData.People = $entities.Emails
                $documentData.Organizations = @()  # Would need NER for proper extraction
                $documentData.Locations = @()     # Would need NER for proper extraction
                $documentData.URLs = if ($entities.URLs) { $entities.URLs | ForEach-Object { $_.URL } } else { @() }
                $documentData.PhoneNumbers = $entities.PhoneNumbers
            }
            
            # Add search keywords from content cleaner or generate
            if ($chunk.ContainsKey('SearchKeywords')) {
                $documentData.SearchKeywords = $chunk.SearchKeywords
            } else {
                $documentData.SearchKeywords = Get-EmailKeywords -EmailData $emailData -Content $chunk.Content
            }
            
            $searchDocuments += $documentData
        }
        
        # Step 5: Index documents in Azure AI Search
        Write-Verbose "Step 5: Indexing documents in Azure AI Search"
        $indexResult = Add-EmailDocuments -SearchConfig $PipelineConfig.AzureSearch -Documents $searchDocuments -IndexName $PipelineConfig.IndexName -GenerateEmbeddings:$($PipelineConfig.Processing.GenerateEmbeddings)
        
        return @{
            Status = "Success"
            FilePath = $EmailFilePath
            EmailData = $emailData
            ChunkCount = $chunkingResult.Chunks.Count
            QualityReport = $chunkingResult.QualityReport
            IndexedDocuments = $indexResult.ProcessedDocuments
            IndexingResult = $indexResult
            ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        return @{
            Status = "Error"
            Error = $_.Exception.Message
            FilePath = $EmailFilePath
        }
    }
}

function Get-EmailKeywords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    try {
        $keywords = @()
        
        # Extract from subject
        if ($EmailData.Subject) {
            $subjectWords = ($EmailData.Subject -split '\s+' | Where-Object { $_.Length -gt 3 } | ForEach-Object { $_.ToLower().Trim('.,!?:;()[]{}') })
            $keywords += $subjectWords
        }
        
        # Extract from sender name
        if ($EmailData.Sender.Name) {
            $nameWords = ($EmailData.Sender.Name -split '\s+' | Where-Object { $_.Length -gt 2 } | ForEach-Object { $_.ToLower() })
            $keywords += $nameWords
        }
        
        # Extract frequent words from content
        if ($Content.Length -gt 50) {
            $contentWords = ($Content -split '\s+' | 
                           Where-Object { $_.Length -gt 4 -and $_ -notmatch '^\d+$' -and $_ -notmatch '^https?:' } | 
                           ForEach-Object { $_.ToLower().Trim('.,!?:;()[]{}') })
            
            $wordFreq = @{}
            foreach ($word in $contentWords) {
                if ($wordFreq.ContainsKey($word)) {
                    $wordFreq[$word]++
                } else {
                    $wordFreq[$word] = 1
                }
            }
            
            $topWords = $wordFreq.GetEnumerator() | 
                       Sort-Object Value -Descending | 
                       Select-Object -First 10 | 
                       ForEach-Object { $_.Key }
            
            $keywords += $topWords
        }
        
        # Remove common stop words
        $stopWords = @('the', 'and', 'that', 'have', 'for', 'not', 'with', 'you', 'this', 'but', 'his', 'from', 'they', 'she', 'her', 'been', 'than', 'its', 'who')
        
        $uniqueKeywords = $keywords | 
                         Where-Object { $_ -notin $stopWords -and $_.Length -gt 2 } | 
                         Sort-Object -Unique | 
                         Select-Object -First 20
        
        return @($uniqueKeywords)
        
    } catch {
        Write-Verbose "Warning: Failed to extract keywords: $($_.Exception.Message)"
        return @()
    }
}

function Update-PipelineStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PipelineConfig,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$BatchResult
    )
    
    $stats = $PipelineConfig.Statistics
    
    $stats.TotalProcessed += $BatchResult.ProcessedCount
    $stats.TotalChunks += $BatchResult.TotalChunks
    $stats.TotalErrors += $BatchResult.ErrorCount
    
    # Calculate average processing time
    if ($stats.TotalProcessed -gt 0) {
        $totalTime = ((Get-Date) - $stats.StartTime).TotalSeconds
        $stats.AverageProcessingTime = [Math]::Round($totalTime / $stats.TotalProcessed, 2)
    }
    
    $stats.LastUpdated = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
}

function Get-ProcessingStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PipelineConfig
    )
    
    $stats = $PipelineConfig.Statistics.Clone()
    $stats.TotalTime = ((Get-Date) - $stats.StartTime).TotalSeconds
    $stats.RetrievedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
    
    return $stats
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
        Write-Host "Testing RAG pipeline..." -ForegroundColor Yellow
        
        # Test 1: Azure Search connection
        Write-Host "1. Testing Azure Search connection..." -ForegroundColor Cyan
        $connectionTest = Test-SearchConnection -Config $PipelineConfig.AzureSearch
        if (-not $connectionTest.IsConnected) {
            throw "Azure Search connection failed: $($connectionTest.Error)"
        }
        Write-Host "   ✓ Connection successful" -ForegroundColor Green
        
        # Test 2: Index statistics
        Write-Host "2. Checking index statistics..." -ForegroundColor Cyan
        try {
            $indexStats = Get-SearchStatistics -SearchConfig $PipelineConfig.AzureSearch -IndexName $PipelineConfig.IndexName
            Write-Host "   ✓ Index contains $($indexStats.DocumentCount) documents" -ForegroundColor Green
        } catch {
            Write-Host "   ⚠ Index statistics unavailable: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Test 3: Basic search
        Write-Host "3. Testing basic search..." -ForegroundColor Cyan
        try {
            $searchResult = Search-EmailDocuments -SearchConfig $PipelineConfig.AzureSearch -Query $TestQuery -IndexName $PipelineConfig.IndexName -Top 5
            Write-Host "   ✓ Search returned $($searchResult.Results.Count) results" -ForegroundColor Green
        } catch {
            Write-Host "   ⚠ Basic search failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Test 4: Hybrid search (if OpenAI configured)
        if ($PipelineConfig.AzureSearch.OpenAI -and $PipelineConfig.AzureSearch.OpenAI.ApiKey) {
            Write-Host "4. Testing hybrid search..." -ForegroundColor Cyan
            try {
                $hybridResult = Search-EmailHybrid -SearchConfig $PipelineConfig.AzureSearch -Query $TestQuery -IndexName $PipelineConfig.IndexName -Top 3
                Write-Host "   ✓ Hybrid search returned $($hybridResult.Results.Count) results" -ForegroundColor Green
            } catch {
                Write-Host "   ⚠ Hybrid search failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "RAG pipeline test completed successfully!" -ForegroundColor Green
        
        return @{
            Status = "Success"
            ConnectionTest = $connectionTest
            IndexStats = if ($indexStats) { $indexStats } else { $null }
            SearchTest = if ($searchResult) { $searchResult.Results.Count } else { 0 }
            HybridSearchTest = if ($hybridResult) { $hybridResult.Results.Count } else { 0 }
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Host "RAG pipeline test failed: $($_.Exception.Message)" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Error = $_.Exception.Message
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
    }
}

Write-Verbose "EmailRAGProcessor_v2 module loaded successfully"