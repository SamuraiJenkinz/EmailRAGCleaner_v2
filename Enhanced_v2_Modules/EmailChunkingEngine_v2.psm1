# EmailChunkingEngine_v2.psm1 - Intelligent Email Chunking for Azure AI Search RAG
# Advanced chunking algorithm optimized for email content and OpenAI embeddings

Export-ModuleMember -Function @(
    'New-EmailChunks',
    'Optimize-ChunkForEmbedding',
    'Test-ChunkQuality',
    'Get-OptimalChunkSize',
    'Merge-SmallChunks',
    'Split-LargeChunks'
)

function New-EmailChunks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$false)]
        [int]$TargetTokens = 384,
        
        [Parameter(Mandatory=$false)]
        [int]$MinTokens = 128,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxTokens = 512,
        
        [Parameter(Mandatory=$false)]
        [int]$OverlapTokens = 32,
        
        [Parameter(Mandatory=$false)]
        [switch]$PreserveStructure = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$OptimizeForSearch = $true
    )
    
    try {
        Write-Verbose "Starting intelligent email chunking for Azure AI Search..."
        
        $content = Get-EmailContentForChunking -EmailData $EmailData
        if (-not $content -or $content.Trim() -eq "") {
            Write-Warning "No content available for chunking"
            return @()
        }
        
        # Initialize chunking context
        $chunkingContext = @{
            EmailSubject = $EmailData.Subject
            SenderName = $EmailData.Sender.Name
            HasQuoteChain = $content -match '^>.*?$'
            HasSignature = $content -match '(?i)(?:best regards|sincerely|thanks)'
            ContentType = if ($EmailData.HTMLBody) { "HTML" } else { "PlainText" }
            OriginalLength = $content.Length
        }
        
        # Pre-process content for optimal chunking
        $processedContent = Preprocess-EmailContent -Content $content -Context $chunkingContext
        
        # Extract email structure elements
        $emailStructure = Extract-EmailStructure -Content $processedContent -EmailData $EmailData
        
        # Create intelligent chunks based on email structure
        $chunks = @()
        
        # Create header chunk with metadata
        $headerChunk = New-HeaderChunk -EmailData $EmailData -Structure $emailStructure
        if ($headerChunk) {
            $chunks += $headerChunk
        }
        
        # Process main content sections
        foreach ($section in $emailStructure.Sections) {
            $sectionChunks = New-SectionChunks -Section $section -TargetTokens $TargetTokens -MinTokens $MinTokens -MaxTokens $MaxTokens -OverlapTokens $OverlapTokens
            $chunks += $sectionChunks
        }
        
        # Post-process chunks for optimization
        if ($OptimizeForSearch) {
            $chunks = Optimize-ChunksForSearch -Chunks $chunks -EmailData $EmailData
        }
        
        # Add chunk relationships and metadata
        $chunks = Add-ChunkMetadata -Chunks $chunks -EmailData $EmailData -Context $chunkingContext
        
        # Validate chunk quality
        $qualityReport = Test-ChunksQuality -Chunks $chunks
        
        Write-Verbose "Created $($chunks.Count) intelligent chunks with average quality score: $($qualityReport.AverageQualityScore)"
        
        return @{
            Chunks = $chunks
            QualityReport = $qualityReport
            ChunkingContext = $chunkingContext
            ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            ChunkingVersion = "2.0"
        }
        
    } catch {
        Write-Error "Failed to create intelligent email chunks: $($_.Exception.Message)"
        throw
    }
}

function New-HeaderChunk {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Structure
    )
    
    try {
        $headerElements = @()
        
        if ($EmailData.Subject) {
            $headerElements += "Subject: $($EmailData.Subject)"
        }
        
        if ($EmailData.Sender.Name) {
            $headerElements += "From: $($EmailData.Sender.Name)"
            if ($EmailData.Sender.Email -and $EmailData.Sender.Email -ne $EmailData.Sender.Name) {
                $headerElements[-1] += " <$($EmailData.Sender.Email)>"
            }
        }
        
        if ($EmailData.Recipients.To -and $EmailData.Recipients.To.Count -gt 0) {
            $toList = $EmailData.Recipients.To | ForEach-Object { 
                if ($_.Name -and $_.Email) { "$($_.Name) <$($_.Email)>" } 
                elseif ($_.Email) { $_.Email } 
                elseif ($_.Name) { $_.Name } 
            }
            $headerElements += "To: $($toList -join ', ')"
        }
        
        if ($EmailData.Recipients.CC -and $EmailData.Recipients.CC.Count -gt 0) {
            $ccList = $EmailData.Recipients.CC | ForEach-Object { 
                if ($_.Name -and $_.Email) { "$($_.Name) <$($_.Email)>" } 
                elseif ($_.Email) { $_.Email } 
                elseif ($_.Name) { $_.Name } 
            }
            $headerElements += "CC: $($ccList -join ', ')"
        }
        
        if ($EmailData.Sent) {
            $headerElements += "Date: $($EmailData.Sent)"
        }
        
        if ($EmailData.Attachments -and $EmailData.Attachments.Count -gt 0) {
            $attachmentNames = $EmailData.Attachments | ForEach-Object { $_.FileName }
            $headerElements += "Attachments: $($attachmentNames -join ', ')"
        }
        
        $headerContent = $headerElements -join "`n"
        
        return @{
            ChunkNumber = 0
            ChunkType = "Header"
            Content = $headerContent
            TokenCount = Get-TokenCount -Text $headerContent
            WordCount = ($headerContent -split '\s+' | Where-Object { $_ -ne "" }).Count
            IsHeader = $true
            ContainsMetadata = $true
            SearchRelevance = "High"
            ProcessingPriority = "Critical"
        }
        
    } catch {
        Write-Verbose "Warning: Failed to create header chunk: $($_.Exception.Message)"
        return $null
    }
}

function New-SectionChunks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Section,
        
        [Parameter(Mandatory=$true)]
        [int]$TargetTokens,
        
        [Parameter(Mandatory=$true)]
        [int]$MinTokens,
        
        [Parameter(Mandatory=$true)]
        [int]$MaxTokens,
        
        [Parameter(Mandatory=$true)]
        [int]$OverlapTokens
    )
    
    $chunks = @()
    $content = $Section.Content
    
    if (-not $content -or $content.Trim() -eq "") {
        return $chunks
    }
    
    $tokenCount = Get-TokenCount -Text $content
    
    # If section fits in target size, create single chunk
    if ($tokenCount -le $TargetTokens) {
        $chunks += @{
            ChunkType = $Section.Type
            Content = $content
            TokenCount = $tokenCount
            WordCount = ($content -split '\s+' | Where-Object { $_ -ne "" }).Count
            SectionType = $Section.Type
            ContainsQuote = $Section.IsQuoted
            IsSignature = $Section.Type -eq "Signature"
            SearchRelevance = Get-SectionSearchRelevance -Section $Section
        }
        return $chunks
    }
    
    # Split large sections intelligently
    $sentences = Split-IntoSentences -Text $content
    $currentChunk = ""
    $currentTokens = 0
    $chunkNumber = 1
    
    foreach ($sentence in $sentences) {
        $sentenceTokens = Get-TokenCount -Text $sentence
        
        # Check if adding this sentence would exceed max tokens
        if (($currentTokens + $sentenceTokens) -gt $MaxTokens -and $currentChunk -ne "") {
            # Create chunk with current content
            $chunks += @{
                ChunkType = $Section.Type
                Content = $currentChunk.Trim()
                TokenCount = $currentTokens
                WordCount = ($currentChunk -split '\s+' | Where-Object { $_ -ne "" }).Count
                SectionType = $Section.Type
                ContainsQuote = $Section.IsQuoted
                IsSignature = $Section.Type -eq "Signature"
                SearchRelevance = Get-SectionSearchRelevance -Section $Section
                ChunkNumber = $chunkNumber
            }
            
            # Start new chunk with overlap
            $overlapText = Get-ChunkOverlap -Text $currentChunk -TargetTokens $OverlapTokens
            $currentChunk = $overlapText + " " + $sentence
            $currentTokens = Get-TokenCount -Text $currentChunk
            $chunkNumber++
        } else {
            # Add sentence to current chunk
            if ($currentChunk -ne "") {
                $currentChunk += " " + $sentence
            } else {
                $currentChunk = $sentence
            }
            $currentTokens += $sentenceTokens
        }
    }
    
    # Add final chunk if content remains
    if ($currentChunk.Trim() -ne "" -and $currentTokens -ge $MinTokens) {
        $chunks += @{
            ChunkType = $Section.Type
            Content = $currentChunk.Trim()
            TokenCount = $currentTokens
            WordCount = ($currentChunk -split '\s+' | Where-Object { $_ -ne "" }).Count
            SectionType = $Section.Type
            ContainsQuote = $Section.IsQuoted
            IsSignature = $Section.Type -eq "Signature"
            SearchRelevance = Get-SectionSearchRelevance -Section $Section
            ChunkNumber = $chunkNumber
        }
    }
    
    return $chunks
}

function Extract-EmailStructure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData
    )
    
    try {
        $sections = @()
        $lines = $Content -split "`n"
        
        $currentSection = @{
            Type = "Body"
            Content = ""
            StartLine = 0
            IsQuoted = $false
        }
        
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            
            # Detect quote chain
            if ($line -match '^>.*' -or $line -match '^On.*wrote:' -or $line -match '^From:.*Sent:') {
                if ($currentSection.Content.Trim() -ne "") {
                    $sections += $currentSection.Clone()
                }
                
                $currentSection = @{
                    Type = "Quote"
                    Content = $line
                    StartLine = $i
                    IsQuoted = $true
                }
                continue
            }
            
            # Detect signature
            if ($line -match '^--\s*$' -or 
                $line -match '(?i)^(best regards|sincerely|thanks|kind regards)' -or 
                $line -match '(?i)sent from my (iphone|android|mobile)') {
                
                if ($currentSection.Content.Trim() -ne "") {
                    $sections += $currentSection.Clone()
                }
                
                # Collect remaining lines as signature
                $signatureLines = $lines[$i..($lines.Count-1)]
                $sections += @{
                    Type = "Signature"
                    Content = ($signatureLines -join "`n").Trim()
                    StartLine = $i
                    IsQuoted = $false
                }
                break
            }
            
            # Add line to current section
            if ($currentSection.Content -ne "") {
                $currentSection.Content += "`n" + $line
            } else {
                $currentSection.Content = $line
            }
        }
        
        # Add final section if content remains
        if ($currentSection.Content.Trim() -ne "") {
            $sections += $currentSection
        }
        
        return @{
            Sections = $sections
            HasQuoteChain = ($sections | Where-Object { $_.Type -eq "Quote" }).Count -gt 0
            HasSignature = ($sections | Where-Object { $_.Type -eq "Signature" }).Count -gt 0
            SectionCount = $sections.Count
        }
        
    } catch {
        Write-Verbose "Warning: Failed to extract email structure: $($_.Exception.Message)"
        return @{
            Sections = @(@{
                Type = "Body"
                Content = $Content
                StartLine = 0
                IsQuoted = $false
            })
            HasQuoteChain = $false
            HasSignature = $false
            SectionCount = 1
        }
    }
}

function Get-TokenCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )
    
    if (-not $Text) {
        return 0
    }
    
    # Approximate token count using GPT-3.5/4 tokenization rules
    # Average: ~0.75 tokens per word, adjusting for punctuation and spaces
    $wordCount = ($Text -split '\s+' | Where-Object { $_ -ne "" }).Count
    $punctuationCount = ([regex]::Matches($Text, '[^\w\s]')).Count
    $approximateTokens = [math]::Ceiling($wordCount * 0.75 + $punctuationCount * 0.25)
    
    return $approximateTokens
}

function Get-SectionSearchRelevance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Section
    )
    
    switch ($Section.Type) {
        "Body" { return "High" }
        "Quote" { return "Medium" }
        "Signature" { return "Low" }
        default { return "Medium" }
    }
}

function Split-IntoSentences {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text
    )
    
    if (-not $Text) {
        return @()
    }
    
    # Split on sentence boundaries while preserving context
    $sentences = $Text -split '(?<=[.!?])\s+(?=[A-Z])' | Where-Object { $_.Trim() -ne "" }
    
    # Handle edge cases where splits are too aggressive
    $refinedSentences = @()
    $currentSentence = ""
    
    foreach ($sentence in $sentences) {
        $sentence = $sentence.Trim()
        if ($sentence.Length -lt 10 -and $currentSentence -ne "") {
            # Merge very short sentences with previous
            $currentSentence += " " + $sentence
        } else {
            if ($currentSentence -ne "") {
                $refinedSentences += $currentSentence
            }
            $currentSentence = $sentence
        }
    }
    
    if ($currentSentence -ne "") {
        $refinedSentences += $currentSentence
    }
    
    return $refinedSentences
}

function Get-ChunkOverlap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text,
        
        [Parameter(Mandatory=$true)]
        [int]$TargetTokens
    )
    
    if (-not $Text -or $TargetTokens -le 0) {
        return ""
    }
    
    $words = $Text -split '\s+' | Where-Object { $_ -ne "" }
    $targetWords = [math]::Ceiling($TargetTokens * 1.33) # Approximate conversion
    
    if ($words.Count -le $targetWords) {
        return $Text
    }
    
    # Take last N words for overlap context
    $overlapWords = $words[($words.Count - $targetWords)..($words.Count - 1)]
    
    # Find last complete sentence in overlap
    $overlapText = $overlapWords -join " "
    $lastSentenceMatch = [regex]::Match($overlapText, '.*[.!?]\s*')
    
    if ($lastSentenceMatch.Success) {
        return $lastSentenceMatch.Value.Trim()
    }
    
    return $overlapText
}

function Optimize-ChunksForSearch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Chunks,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData
    )
    
    $optimizedChunks = @()
    
    foreach ($chunk in $Chunks) {
        $optimizedChunk = $chunk.Clone()
        
        # Add context from email metadata
        $contextElements = @()
        
        if ($EmailData.Subject -and $chunk.ChunkType -ne "Header") {
            $contextElements += "Email Subject: $($EmailData.Subject)"
        }
        
        if ($EmailData.Sender.Name -and $chunk.ChunkType -ne "Header") {
            $contextElements += "From: $($EmailData.Sender.Name)"
        }
        
        if ($contextElements.Count -gt 0) {
            $contextString = $contextElements -join " | "
            $optimizedChunk.Content = $contextString + "`n`n" + $optimizedChunk.Content
            $optimizedChunk.TokenCount = Get-TokenCount -Text $optimizedChunk.Content
            $optimizedChunk.WordCount = ($optimizedChunk.Content -split '\s+' | Where-Object { $_ -ne "" }).Count
            $optimizedChunk.HasContext = $true
        }
        
        # Add search optimization flags
        $optimizedChunk.OptimizedForSearch = $true
        $optimizedChunk.SearchWeight = switch ($chunk.SearchRelevance) {
            "High" { 1.0 }
            "Medium" { 0.7 }
            "Low" { 0.3 }
            default { 0.5 }
        }
        
        $optimizedChunks += $optimizedChunk
    }
    
    return $optimizedChunks
}

function Add-ChunkMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Chunks,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Context
    )
    
    $totalChunks = $Chunks.Count
    
    for ($i = 0; $i -lt $totalChunks; $i++) {
        $chunk = $Chunks[$i]
        
        # Add standard metadata
        $chunk.Id = "$($EmailData.Metadata.FileName)_chunk_$($i + 1)"
        $chunk.ChunkNumber = $i + 1
        $chunk.TotalChunks = $totalChunks
        $chunk.IsFirst = ($i -eq 0)
        $chunk.IsLast = ($i -eq ($totalChunks - 1))
        
        # Add relationships
        if ($i -gt 0) {
            $chunk.PreviousChunkId = $Chunks[$i - 1].Id
        }
        if ($i -lt ($totalChunks - 1)) {
            $chunk.NextChunkId = $Chunks[$i + 1].Id
        }
        
        # Add email context
        $chunk.ParentEmailId = $EmailData.Metadata.FileName -replace '\.msg$', ''
        $chunk.EmailSubject = $EmailData.Subject
        $chunk.SenderName = $EmailData.Sender.Name
        $chunk.ProcessedAt = $Context.ProcessedAt
        
        # Add quality metrics
        $chunk.QualityScore = Get-ChunkQualityScore -Chunk $chunk
        $chunk.SearchReadiness = Test-SearchReadiness -Chunk $chunk
    }
    
    return $Chunks
}

function Get-ChunkQualityScore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Chunk
    )
    
    $score = 0
    
    # Token count optimization (384 is optimal)
    if ($Chunk.TokenCount -ge 128 -and $Chunk.TokenCount -le 512) {
        $tokenScore = 1.0 - [math]::Abs($Chunk.TokenCount - 384) / 384.0
        $score += $tokenScore * 30
    }
    
    # Content completeness
    if ($Chunk.Content -match '\.$') { $score += 10 }
    if ($Chunk.WordCount -gt 10) { $score += 10 }
    if (-not ($Chunk.Content -match '^\s*$')) { $score += 20 }
    
    # Structure preservation
    if ($Chunk.ChunkType -eq "Header") { $score += 15 }
    if ($Chunk.HasContext) { $score += 10 }
    
    # Search relevance
    switch ($Chunk.SearchRelevance) {
        "High" { $score += 5 }
        "Medium" { $score += 3 }
        "Low" { $score += 1 }
    }
    
    return [math]::Round([math]::Min(100, [math]::Max(0, $score)), 1)
}

function Test-SearchReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Chunk
    )
    
    $issues = @()
    
    if ($Chunk.TokenCount -lt 32) {
        $issues += "Token count too low for effective embeddings"
    }
    
    if ($Chunk.TokenCount -gt 512) {
        $issues += "Token count exceeds OpenAI embedding limit"
    }
    
    if ($Chunk.Content.Length -lt 10) {
        $issues += "Content too short for meaningful search"
    }
    
    if ($Chunk.Content -match '^\s*$') {
        $issues += "Empty or whitespace-only content"
    }
    
    return @{
        IsReady = $issues.Count -eq 0
        Issues = $issues
        ReadinessScore = if ($issues.Count -eq 0) { 100 } else { [math]::Max(0, 100 - ($issues.Count * 25)) }
    }
}

function Test-ChunksQuality {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Chunks
    )
    
    if ($Chunks.Count -eq 0) {
        return @{
            TotalChunks = 0
            AverageQualityScore = 0
            AverageTokenCount = 0
            SearchReadyChunks = 0
            Issues = @("No chunks created")
        }
    }
    
    $qualityScores = $Chunks | ForEach-Object { $_.QualityScore }
    $tokenCounts = $Chunks | ForEach-Object { $_.TokenCount }
    $readyChunks = ($Chunks | Where-Object { $_.SearchReadiness.IsReady }).Count
    
    return @{
        TotalChunks = $Chunks.Count
        AverageQualityScore = [math]::Round(($qualityScores | Measure-Object -Average).Average, 1)
        AverageTokenCount = [math]::Round(($tokenCounts | Measure-Object -Average).Average, 0)
        MinTokenCount = ($tokenCounts | Measure-Object -Minimum).Minimum
        MaxTokenCount = ($tokenCounts | Measure-Object -Maximum).Maximum
        SearchReadyChunks = $readyChunks
        SearchReadinessPercentage = [math]::Round(($readyChunks / $Chunks.Count) * 100, 1)
        OptimalTokenRangePercentage = [math]::Round((($Chunks | Where-Object { $_.TokenCount -ge 256 -and $_.TokenCount -le 512 }).Count / $Chunks.Count) * 100, 1)
    }
}

function Get-EmailContentForChunking {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData
    )
    
    # Priority order for content selection
    if ($EmailData.Content -and $EmailData.Content.CleanedText) {
        return $EmailData.Content.CleanedText
    }
    
    if ($EmailData.ProcessedContent -and $EmailData.ProcessedContent.ExtractedText) {
        return $EmailData.ProcessedContent.ExtractedText
    }
    
    if ($EmailData.ProcessedContent -and $EmailData.ProcessedContent.PlainText) {
        return $EmailData.ProcessedContent.PlainText
    }
    
    if ($EmailData.Body -and $EmailData.Body.Trim() -ne "") {
        return $EmailData.Body
    }
    
    if ($EmailData.HTMLBody -and $EmailData.HTMLBody.Trim() -ne "") {
        # Quick HTML to text conversion
        $text = $EmailData.HTMLBody -replace '<[^>]*>', ''
        $text = $text -replace '&nbsp;', ' '
        $text = $text -replace '&amp;', '&'
        $text = $text -replace '&lt;', '<'
        $text = $text -replace '&gt;', '>'
        return $text.Trim()
    }
    
    return ""
}

function Preprocess-EmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Context
    )
    
    if (-not $Content) {
        return ""
    }
    
    # Normalize line endings
    $processed = $Content -replace '\r\n', "`n"
    $processed = $processed -replace '\r', "`n"
    
    # Clean up excessive whitespace while preserving structure
    $processed = $processed -replace ' +', ' '
    $processed = $processed -replace '\n\s*\n\s*\n+', "`n`n"
    
    # Trim lines
    $lines = $processed -split "`n"
    $lines = $lines | ForEach-Object { $_.Trim() }
    $processed = $lines -join "`n"
    
    return $processed.Trim()
}

Write-Verbose "EmailChunkingEngine_v2 module loaded successfully"