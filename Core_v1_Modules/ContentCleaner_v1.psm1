# ContentCleaner.psm1 - Email Content Cleaning and RAG Optimization Module (Clean Version)
# Professional content processing with entity extraction and RAG chunking

Export-ModuleMember -Function @(
    'Clean-EmailContent',
    'Add-RAGChunks',
    'Extract-Entities',
    'Remove-EmailSignatures',
    'Optimize-ContentForRAG',
    'Get-ContentQualityScore'
)

function Clean-EmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$false)]
        [switch]$ExtractEntities = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$RemoveSignatures = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$OptimizeForRAG = $true
    )
    
    try {
        Write-Verbose "Starting content cleaning for: $($EmailData.Subject)"
        
        if (-not $EmailData.ContainsKey('Content')) {
            $EmailData.Content = @{}
        }
        
        $originalContent = Get-BestContent -EmailData $EmailData
        if (-not $originalContent) {
            Write-Warning "No content found to clean"
            return $EmailData
        }
        
        $EmailData.Content.OriginalContent = $originalContent
        $EmailData.Content.OriginalLength = $originalContent.Length
        
        $cleanedContent = $originalContent
        if ($EmailData.HTMLBody -and $EmailData.HTMLBody.Trim() -ne "") {
            Write-Verbose "Sanitizing HTML content..."
            $cleanedContent = Remove-HtmlTracking -HtmlContent $cleanedContent
            $cleanedContent = Sanitize-HtmlContent -HtmlContent $cleanedContent
            $cleanedContent = ConvertFrom-HtmlToText -HtmlContent $cleanedContent
        }
        
        if ($RemoveSignatures) {
            Write-Verbose "Removing email signatures..."
            $cleanedContent = Remove-EmailSignatures -Content $cleanedContent
        }
        
        Write-Verbose "Normalizing content formatting..."
        $cleanedContent = Normalize-ContentFormatting -Content $cleanedContent
        
        if ($ExtractEntities) {
            Write-Verbose "Extracting entities from content..."
            $EmailData.Content.ExtractedEntities = Extract-Entities -Content $cleanedContent
        }
        
        if ($OptimizeForRAG) {
            Write-Verbose "Optimizing content for RAG..."
            $cleanedContent = Optimize-ContentForRAG -Content $cleanedContent
        }
        
        $EmailData.Content.CleanedText = $cleanedContent
        $EmailData.Content.CleanedLength = $cleanedContent.Length
        
        $EmailData.Content.QualityScore = Get-ContentQualityScore -Content $cleanedContent
        
        $EmailData.Content.ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        $EmailData.Content.ProcessingVersion = "1.0"
        $EmailData.Content.ReductionRatio = if ($EmailData.Content.OriginalLength -gt 0) {
            [math]::Round((1 - ($EmailData.Content.CleanedLength / $EmailData.Content.OriginalLength)) * 100, 2)
        } else { 0 }
        
        Write-Verbose "Content cleaning completed. Reduction: $($EmailData.Content.ReductionRatio)%"
        return $EmailData
        
    } catch {
        Write-Error "Failed to clean email content: $($_.Exception.Message)"
        throw
    }
}

function Add-RAGChunks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$false)]
        [int]$ChunkSize = 512,
        
        [Parameter(Mandatory=$false)]
        [int]$Overlap = 50
    )
    
    try {
        Write-Verbose "Creating RAG chunks with size: $ChunkSize, overlap: $Overlap"
        
        $content = ""
        if ($EmailData.Content -and $EmailData.Content.CleanedText) {
            $content = $EmailData.Content.CleanedText
        } else {
            $content = Get-BestContent -EmailData $EmailData
        }
        
        if (-not $content -or $content.Trim() -eq "") {
            Write-Warning "No content available for chunking"
            $EmailData.Content.Chunks = @()
            return $EmailData
        }
        
        if (-not $EmailData.ContainsKey('Content')) {
            $EmailData.Content = @{}
        }
        
        $chunks = @()
        $contentLength = $content.Length
        
        if ($contentLength -le $ChunkSize) {
            $chunks += @{
                ChunkNumber = 1
                Content = $content
                StartPosition = 0
                EndPosition = $contentLength - 1
                Length = $contentLength
                WordCount = ($content -split '\s+' | Where-Object { $_ -ne "" }).Count
                OverlapWithNext = $false
                OverlapWithPrevious = $false
            }
        } else {
            $position = 0
            $chunkNumber = 1
            
            while ($position -lt $contentLength) {
                $endPosition = [Math]::Min($position + $ChunkSize - 1, $contentLength - 1)
                
                if ($endPosition -lt $contentLength - 1) {
                    $nextSpace = $content.IndexOf(' ', $endPosition)
                    $prevSpace = $content.LastIndexOf(' ', $endPosition)
                    
                    if ($nextSpace -ne -1 -and $prevSpace -ne -1) {
                        if (($nextSpace - $endPosition) -le ($endPosition - $prevSpace)) {
                            $endPosition = $nextSpace
                        } else {
                            $endPosition = $prevSpace
                        }
                    } elseif ($prevSpace -ne -1) {
                        $endPosition = $prevSpace
                    }
                }
                
                $chunkContent = $content.Substring($position, $endPosition - $position + 1).Trim()
                
                if ($chunkContent -ne "") {
                    $chunks += @{
                        ChunkNumber = $chunkNumber
                        Content = $chunkContent
                        StartPosition = $position
                        EndPosition = $endPosition
                        Length = $chunkContent.Length
                        WordCount = ($chunkContent -split '\s+' | Where-Object { $_ -ne "" }).Count
                        OverlapWithNext = ($endPosition + 1) -lt $contentLength
                        OverlapWithPrevious = $position -gt 0
                    }
                    $chunkNumber++
                }
                
                $position = [Math]::Max($endPosition + 1 - $Overlap, $position + 1)
                
                if ($position -ge $contentLength) {
                    break
                }
            }
        }
        
        for ($i = 0; $i -lt $chunks.Count; $i++) {
            $chunks[$i].TotalChunks = $chunks.Count
            $chunks[$i].IsFirst = ($i -eq 0)
            $chunks[$i].IsLast = ($i -eq ($chunks.Count - 1))
            $chunks[$i].ChunkId = "$($EmailData.Metadata.FileName)_chunk_$($chunks[$i].ChunkNumber)"
        }
        
        $EmailData.Content.Chunks = $chunks
        $EmailData.Content.ChunkCount = $chunks.Count
        $EmailData.Content.AverageChunkSize = if ($chunks.Count -gt 0) {
            [math]::Round(($chunks | Measure-Object -Property Length -Average).Average, 0)
        } else { 0 }
        
        Write-Verbose "Created $($chunks.Count) RAG chunks with average size: $($EmailData.Content.AverageChunkSize)"
        return $EmailData
        
    } catch {
        Write-Error "Failed to create RAG chunks: $($_.Exception.Message)"
        throw
    }
}

function Extract-Entities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Content
    )
    
    try {
        if (-not $Content -or $Content.Trim() -eq "") {
            return @{
                Emails = @()
                URLs = @()
                PhoneNumbers = @()
                Dates = @()
                IPAddresses = @()
                Numbers = @()
                EntityCount = 0
            }
        }
        
        Write-Verbose "Extracting entities from content..."
        
        # Email addresses - fixed regex pattern
        $emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        $emails = ([regex]::Matches($Content, $emailPattern) | ForEach-Object { $_.Value } | Sort-Object -Unique)
        
        # URLs - simplified pattern
        $urlPattern = 'https?://[^\s<>"{}|\\\^\[\]`;/?:@=&]*'
        $urls = ([regex]::Matches($Content, $urlPattern) | ForEach-Object { 
            @{
                URL = $_.Value
                Domain = try { 
                    $uri = [System.Uri]$_.Value
                    $uri.Host 
                } catch { 
                    "Unknown" 
                }
                IsSecure = $_.Value.StartsWith("https://")
            }
        })
        
        # Phone numbers - simplified patterns
        $phonePatterns = @(
            '\b\d{3}[-.]?\d{3}[-.]?\d{4}\b',
            '\(\d{3}\)\s?\d{3}[-.]?\d{4}',
            '\+\d{1,3}[-.\s]?\d{3,4}[-.\s]?\d{3,4}[-.\s]?\d{3,4}',
            '\b\d{3}\s\d{3}\s\d{4}\b'
        )
        
        $phoneNumbers = @()
        foreach ($pattern in $phonePatterns) {
            $phoneNumbers += ([regex]::Matches($Content, $pattern) | ForEach-Object { $_.Value })
        }
        $phoneNumbers = $phoneNumbers | Sort-Object -Unique
        
        # Dates - simplified patterns
        $datePatterns = @(
            '\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b',
            '\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b',
            '\b\w+\s+\d{1,2},?\s+\d{4}\b',
            '\b\d{1,2}\s+\w+\s+\d{4}\b'
        )
        
        $dates = @()
        foreach ($pattern in $datePatterns) {
            $dates += ([regex]::Matches($Content, $pattern) | ForEach-Object { $_.Value })
        }
        $dates = $dates | Sort-Object -Unique
        
        # IP Addresses
        $ipPattern = '\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
        $ipAddresses = ([regex]::Matches($Content, $ipPattern) | ForEach-Object { $_.Value } | Sort-Object -Unique)
        
        # Numbers - simplified patterns
        $numberPatterns = @(
            '\$\d{1,3}(?:,\d{3})*(?:\.\d{2})?',
            '\d+(?:\.\d+)?%',
            '\b\d{1,3}(?:,\d{3})+\b',
            '\b\d+\.\d+\b'
        )
        
        $numbers = @()
        foreach ($pattern in $numberPatterns) {
            $numbers += ([regex]::Matches($Content, $pattern) | ForEach-Object { $_.Value })
        }
        $numbers = $numbers | Sort-Object -Unique
        
        $entities = @{
            Emails = @($emails)
            URLs = @($urls)
            PhoneNumbers = @($phoneNumbers)
            Dates = @($dates)
            IPAddresses = @($ipAddresses)
            Numbers = @($numbers)
            EntityCount = $emails.Count + $urls.Count + $phoneNumbers.Count + $dates.Count + $ipAddresses.Count + $numbers.Count
            ExtractedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
        Write-Verbose "Extracted $($entities.EntityCount) entities: $($emails.Count) emails, $($urls.Count) URLs, $($phoneNumbers.Count) phones, $($dates.Count) dates"
        return $entities
        
    } catch {
        Write-Error "Failed to extract entities: $($_.Exception.Message)"
        return @{
            Emails = @()
            URLs = @()
            PhoneNumbers = @()
            Dates = @()
            IPAddresses = @()
            Numbers = @()
            EntityCount = 0
            Error = $_.Exception.Message
        }
    }
}

function Remove-EmailSignatures {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Content
    )
    
    try {
        if (-not $Content -or $Content.Trim() -eq "") {
            return $Content
        }
        
        $signaturePatterns = @(
            '(?s)^--\s*$.*',
            '(?s)^_{3,}.*',
            '(?s)^-{3,}.*',
            '(?s)^={3,}.*',
            '(?s)\b(?:best\s+regards?|sincerely|thanks?\s+(?:and\s+)?regards?|kind\s+regards?)\b.*$',
            '(?s)\b(?:sent\s+from\s+my\s+(?:iphone|ipad|android|blackberry|mobile))\b.*$',
            '(?s)\b(?:get\s+outlook\s+for\s+(?:ios|android))\b.*$',
            '(?s)\b(?:phone|tel|mobile|email|fax):\s*[+\d\s\-\(\)\.]+.*$',
            '(?s)\b\w+@\w+\.\w+\b.*(?:\n.*){0,5}$',
            '(?s)\b(?:confidential|disclaimer|privileged|proprietary)\b.*$',
            '(?s)\b(?:this\s+email\s+(?:and\s+any\s+)?(?:attachments?\s+)?(?:are?\s+)?(?:confidential|privileged))\b.*$'
        )
        
        $cleanedContent = $Content
        
        foreach ($pattern in $signaturePatterns) {
            $cleanedContent = $cleanedContent -replace $pattern, ''
        }
        
        $cleanedContent = $cleanedContent.TrimEnd()
        $cleanedContent = $cleanedContent -replace '\n\s*\n\s*\n', "`n`n"
        
        return $cleanedContent
        
    } catch {
        Write-Verbose "Warning: Failed to remove signatures: $($_.Exception.Message)"
        return $Content
    }
}

function Optimize-ContentForRAG {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Content
    )
    
    try {
        if (-not $Content -or $Content.Trim() -eq "") {
            return $Content
        }
        
        $optimizedContent = $Content
        
        $optimizedContent = $optimizedContent -replace '\s+', ' '
        $optimizedContent = $optimizedContent -replace '\r\n', "`n"
        $optimizedContent = $optimizedContent -replace '\r', "`n"
        $optimizedContent = $optimizedContent -replace '\n{3,}', "`n`n"
        
        $lines = $optimizedContent -split '\n'
        $lines = $lines | ForEach-Object { $_.Trim() }
        $optimizedContent = $lines -join "`n"
        
        $optimizedContent = $optimizedContent.Trim()
        
        if ($optimizedContent -and -not ($optimizedContent -match '[.!?]$')) {
            $optimizedContent += '.'
        }
        
        return $optimizedContent
        
    } catch {
        Write-Verbose "Warning: Failed to optimize content for RAG: $($_.Exception.Message)"
        return $Content
    }
}

function Get-ContentQualityScore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Content
    )
    
    try {
        if (-not $Content -or $Content.Trim() -eq "") {
            return @{
                OverallScore = 0
                Length = 0
                WordCount = 0
                SentenceCount = 0
                ParagraphCount = 0
                ReadabilityScore = 0
                HasMeaningfulContent = $false
            }
        }
        
        $length = $Content.Length
        $words = ($Content -split '\s+' | Where-Object { $_ -ne "" })
        $wordCount = $words.Count
        
        $sentences = ($Content -split '[.!?]+' | Where-Object { $_.Trim() -ne "" })
        $sentenceCount = $sentences.Count
        
        $paragraphs = ($Content -split '\n\s*\n' | Where-Object { $_.Trim() -ne "" })
        $paragraphCount = $paragraphs.Count
        
        $avgWordsPerSentence = if ($sentenceCount -gt 0) { $wordCount / $sentenceCount } else { 0 }
        $avgSentencesPerParagraph = if ($paragraphCount -gt 0) { $sentenceCount / $paragraphCount } else { 0 }
        
        $readabilityScore = 0
        if ($avgWordsPerSentence -gt 0) {
            $readabilityScore = [Math]::Max(0, [Math]::Min(100, 100 - ($avgWordsPerSentence * 2)))
        }
        
        $hasMeaningfulContent = (
            $wordCount -gt 10 -and
            $sentenceCount -gt 1 -and
            ($Content -match '\w{3,}') -and
            -not ($Content -match '^[^a-zA-Z]*$')
        )
        
        $overallScore = 0
        if ($hasMeaningfulContent) {
            $lengthScore = [Math]::Min(100, ($length / 10))
            $wordScore = [Math]::Min(100, ($wordCount * 2))
            $structureScore = [Math]::Min(100, ($paragraphCount * 20))
            
            $overallScore = ($lengthScore * 0.3) + ($wordScore * 0.4) + ($structureScore * 0.2) + ($readabilityScore * 0.1)
            $overallScore = [Math]::Round([Math]::Min(100, $overallScore), 1)
        }
        
        return @{
            OverallScore = $overallScore
            Length = $length
            WordCount = $wordCount
            SentenceCount = $sentenceCount
            ParagraphCount = $paragraphCount
            ReadabilityScore = [Math]::Round($readabilityScore, 1)
            AverageWordsPerSentence = [Math]::Round($avgWordsPerSentence, 1)
            AverageSentencesPerParagraph = [Math]::Round($avgSentencesPerParagraph, 1)
            HasMeaningfulContent = $hasMeaningfulContent
            AnalyzedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Verbose "Warning: Failed to calculate quality score: $($_.Exception.Message)"
        return @{
            OverallScore = 0
            Error = $_.Exception.Message
        }
    }
}

# Helper functions
function Get-BestContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData
    )
    
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
        return ConvertFrom-HtmlToText -HtmlContent $EmailData.HTMLBody
    }
    
    return ""
}

function Remove-HtmlTracking {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )
    
    # Fixed tracking pixel regex - using backticks to escape quotes properly
    $cleaned = $HtmlContent -replace '<img[^>]*width\s*=\s*["`'']*1["`'']*[^>]*height\s*=\s*["`'']*1["`'']*[^>]*>', ''
    
    $cleaned = $cleaned -replace '<script[^>]*google-analytics[^>]*>.*?</script>', ''
    $cleaned = $cleaned -replace '<script[^>]*gtag[^>]*>.*?</script>', ''
    
    # Fixed URL tracking parameter regex - escape ampersand properly
    $cleaned = $cleaned -replace 'utm_[^=]*=[^"&\s'']*[&]?', ''
    
    return $cleaned
}

function Sanitize-HtmlContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )
    
    $sanitized = $HtmlContent -replace '<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', ''
    $sanitized = $sanitized -replace '<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>', ''
    $sanitized = $sanitized -replace '<embed[^>]*>', ''
    $sanitized = $sanitized -replace '<applet\b[^<]*(?:(?!<\/applet>)<[^<]*)*<\/applet>', ''
    $sanitized = $sanitized -replace '<form\b[^<]*(?:(?!<\/form>)<[^<]*)*<\/form>', ''
    
    # Fixed event handler regex
    $sanitized = $sanitized -replace '\s+on\w+\s*=\s*["`''][^"`'']*["`'']', ''
    
    return $sanitized
}

function ConvertFrom-HtmlToText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )
    
    try {
        if (-not $HtmlContent) {
            return ""
        }
        
        $text = $HtmlContent -replace '<!--.*?-->', ''
        $text = $text -replace '<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', ''
        $text = $text -replace '<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>', ''
        
        $text = $text -replace '<br\s*/?>', "`n"
        $text = $text -replace '</?p[^>]*>', "`n`n"
        $text = $text -replace '</?div[^>]*>', "`n"
        $text = $text -replace '</?h[1-6][^>]*>', "`n`n"
        
        $text = $text -replace '<li[^>]*>', "`n• "
        $text = $text -replace '</li>', ""
        $text = $text -replace '</?[uo]l[^>]*>', "`n"
        
        $text = $text -replace '<[^>]*>', ''
        
        $text = $text -replace '&amp;', '&'
        $text = $text -replace '&lt;', '<'
        $text = $text -replace '&gt;', '>'
        $text = $text -replace '&quot;', '"'
        $text = $text -replace '&#39;', "'"
        $text = $text -replace '&nbsp;', ' '
        
        $text = $text -replace '\s+', ' '
        $text = $text -replace '\n\s*\n\s*\n+', "`n`n"
        $text = $text.Trim()
        
        return $text
        
    } catch {
        Write-Verbose "Warning: Failed to convert HTML to text: $($_.Exception.Message)"
        return $HtmlContent
    }
}

function Normalize-ContentFormatting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    if (-not $Content) {
        return ""
    }
    
    $normalized = $Content -replace '\r\n', "`n"
    $normalized = $normalized -replace '\r', "`n"
    $normalized = $normalized -replace '[ \t]+', ' '
    $normalized = $normalized -replace '\n{3,}', "`n`n"
    
    $lines = $normalized -split '\n'
    $lines = $lines | ForEach-Object { $_.Trim() }
    $normalized = $lines -join "`n"
    
    return $normalized.Trim()
}

Write-Verbose "ContentCleaner module loaded successfully"