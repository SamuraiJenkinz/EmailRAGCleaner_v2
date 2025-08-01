# AzureFlattener.psm1 - Azure AI Search Document Flattening Module (Clean Version)
# Converts email data to flattened JSON format for vector databases and Azure AI Search

Export-ModuleMember -Function @(
    'ConvertTo-AzureSearchFormat',
    'ConvertTo-VectorDatabaseFormat',
    'Test-AzureSearchCompatibility',
    'Get-SearchableFields',
    'Optimize-ForEmbedding'
)

function ConvertTo-AzureSearchFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeChunks = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$OptimizeForSearch = $true
    )
    
    try {
        Write-Verbose "Converting email data to Azure Search format..."
        
        $documents = @()
        $baseId = Get-SafeId -InputString $EmailData.Subject -Prefix "email"
        
        $mainContent = Get-BestEmailContent -EmailData $EmailData
        if ($mainContent -and $mainContent.Trim() -ne "") {
            $mainDoc = Create-BaseSearchDocument -EmailData $EmailData -BaseId $baseId -Content $mainContent
            
            if ($OptimizeForSearch) {
                $mainDoc = Optimize-SearchDocument -Document $mainDoc
            }
            
            $documents += $mainDoc
            Write-Verbose "Created main document: $($mainDoc.id)"
        }
        
        if ($IncludeChunks -and $EmailData.Content -and $EmailData.Content.Chunks) {
            Write-Verbose "Creating chunk documents..."
            
            foreach ($chunk in $EmailData.Content.Chunks) {
                if ($chunk.Content -and $chunk.Content.Trim() -ne "") {
                    $chunkDoc = Create-ChunkSearchDocument -EmailData $EmailData -Chunk $chunk -BaseId $baseId
                    
                    if ($OptimizeForSearch) {
                        $chunkDoc = Optimize-SearchDocument -Document $chunkDoc
                    }
                    
                    $documents += $chunkDoc
                }
            }
            
            Write-Verbose "Created $($EmailData.Content.Chunks.Count) chunk documents"
        }
        
        Write-Verbose "Total documents created: $($documents.Count)"
        return $documents
        
    } catch {
        Write-Error "Failed to convert to Azure Search format: $($_.Exception.Message)"
        throw
    }
}

function ConvertTo-VectorDatabaseFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Pinecone", "Weaviate", "Qdrant", "Chroma", "Generic")]
        [string]$DatabaseType = "Generic",
        
        [Parameter(Mandatory=$false)]
        [switch]$ChunkOnly = $false
    )
    
    try {
        Write-Verbose "Converting email data to $DatabaseType vector database format..."
        
        $documents = @()
        $baseId = Get-SafeId -InputString $EmailData.Subject -Prefix "email"
        
        if ($EmailData.Content -and $EmailData.Content.Chunks) {
            foreach ($chunk in $EmailData.Content.Chunks) {
                if ($chunk.Content -and $chunk.Content.Trim() -ne "") {
                    $vectorDoc = Create-VectorDocument -EmailData $EmailData -Chunk $chunk -BaseId $baseId -DatabaseType $DatabaseType
                    $documents += $vectorDoc
                }
            }
        } elseif (-not $ChunkOnly) {
            $mainContent = Get-BestEmailContent -EmailData $EmailData
            if ($mainContent -and $mainContent.Trim() -ne "") {
                $vectorDoc = Create-VectorDocument -EmailData $EmailData -Content $mainContent -BaseId $baseId -DatabaseType $DatabaseType
                $documents += $vectorDoc
            }
        }
        
        Write-Verbose "Created $($documents.Count) vector database documents"
        return $documents
        
    } catch {
        Write-Error "Failed to convert to vector database format: $($_.Exception.Message)"
        throw
    }
}

function Test-AzureSearchCompatibility {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Documents
    )
    
    try {
        Write-Verbose "Testing Azure Search compatibility for $($Documents.Count) documents..."
        
        $issues = @()
        $warnings = @()
        
        foreach ($doc in $Documents) {
            $docId = if ($doc.id) { $doc.id } else { "Unknown" }
            
            $requiredFields = @('id', 'DocumentType')
            foreach ($field in $requiredFields) {
                if (-not $doc.ContainsKey($field) -or [string]::IsNullOrWhiteSpace($doc.$field)) {
                    $issues += "Document $docId missing required field: $field"
                }
            }
            
            foreach ($fieldName in $doc.Keys) {
                $value = $doc.$fieldName
                
                if ($value -is [string]) {
                    if ($fieldName -eq 'id' -and $value.Length -gt 1024) {
                        $issues += "Document $docId has ID longer than 1024 characters"
                    } elseif ($fieldName -notin @('AllText', 'ChunkContent') -and $value.Length -gt 8000) {
                        $warnings += "Document $docId field $fieldName exceeds recommended length of 8000 characters"
                    }
                }
                
                if ($fieldName -eq 'SearchKeywords') {
                    if ($value -isnot [array] -and $value -isnot [System.Collections.ArrayList]) {
                        $issues += "Document $docId SearchKeywords must be an array"
                    } else {
                        foreach ($keyword in $value) {
                            if ($keyword -isnot [string]) {
                                $issues += "Document $docId SearchKeywords contains non-string value: $keyword"
                            }
                        }
                    }
                }
                
                if ($fieldName -match '[^a-zA-Z0-9_]') {
                    $issues += "Document $docId has invalid field name: $fieldName (contains special characters)"
                }
            }
        }
        
        $isCompatible = $issues.Count -eq 0
        
        Write-Verbose "Compatibility test completed. Compatible: $isCompatible, Issues: $($issues.Count), Warnings: $($warnings.Count)"
        
        return @{
            IsCompatible = $isCompatible
            Issues = $issues
            Warnings = $warnings
            TestedDocuments = $Documents.Count
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to test Azure Search compatibility: $($_.Exception.Message)"
        
        return @{
            IsCompatible = $false
            Issues = @("Compatibility test error: $($_.Exception.Message)")
            Warnings = @()
            TestedDocuments = 0
        }
    }
}

function Get-SearchableFields {
    [CmdletBinding()]
    param()
    
    return @{
        Fields = @(
            @{
                Name = "id"
                Type = "Edm.String"
                Key = $true
                Searchable = $false
                Filterable = $true
                Sortable = $false
                Facetable = $false
                Retrievable = $true
            },
            @{
                Name = "DocumentType"
                Type = "Edm.String"
                Key = $false
                Searchable = $false
                Filterable = $true
                Sortable = $true
                Facetable = $true
                Retrievable = $true
            },
            @{
                Name = "Subject"
                Type = "Edm.String"
                Key = $false
                Searchable = $true
                Filterable = $true
                Sortable = $true
                Facetable = $false
                Retrievable = $true
            },
            @{
                Name = "Sender_Name"
                Type = "Edm.String"
                Key = $false
                Searchable = $true
                Filterable = $true
                Sortable = $true
                Facetable = $true
                Retrievable = $true
            },
            @{
                Name = "Sender_Email"
                Type = "Edm.String"
                Key = $false
                Searchable = $true
                Filterable = $true
                Sortable = $false
                Facetable = $true
                Retrievable = $true
            },
            @{
                Name = "ChunkContent"
                Type = "Edm.String"
                Key = $false
                Searchable = $true
                Filterable = $false
                Sortable = $false
                Facetable = $false
                Retrievable = $true
            },
            @{
                Name = "SearchKeywords"
                Type = "Collection(Edm.String)"
                Key = $false
                Searchable = $true
                Filterable = $true
                Sortable = $false
                Facetable = $true
                Retrievable = $true
            },
            @{
                Name = "ProcessedAt"
                Type = "Edm.DateTimeOffset"
                Key = $false
                Searchable = $false
                Filterable = $true
                Sortable = $true
                Facetable = $true
                Retrievable = $true
            },
            @{
                Name = "AllText"
                Type = "Edm.String"
                Key = $false
                Searchable = $true
                Filterable = $false
                Sortable = $false
                Facetable = $false
                Retrievable = $true
            }
        )
    }
}

function Optimize-ForEmbedding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Document
    )
    
    try {
        $optimized = $Document.Clone()
        
        $textFields = @()
        
        if ($optimized.Subject) {
            $textFields += "Subject: $($optimized.Subject)"
        }
        
        if ($optimized.Sender_Name) {
            $textFields += "From: $($optimized.Sender_Name)"
        }
        
        if ($optimized.ChunkContent) {
            $textFields += $optimized.ChunkContent
        } elseif ($optimized.AllText) {
            $textFields += $optimized.AllText
        }
        
        if ($optimized.SearchKeywords -and $optimized.SearchKeywords.Count -gt 0) {
            $textFields += "Keywords: $($optimized.SearchKeywords -join ', ')"
        }
        
        $optimized.EmbeddingText = ($textFields -join "`n`n").Trim()
        
        if ($optimized.EmbeddingText.Length -gt 30000) {
            $optimized.EmbeddingText = $optimized.EmbeddingText.Substring(0, 30000) + "..."
            $optimized.IsTruncated = $true
        } else {
            $optimized.IsTruncated = $false
        }
        
        $optimized.EmbeddingOptimizedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        $optimized.EmbeddingLength = $optimized.EmbeddingText.Length
        
        return $optimized
        
    } catch {
        Write-Warning "Failed to optimize document for embedding: $($_.Exception.Message)"
        return $Document
    }
}

# Helper functions
function Create-BaseSearchDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$true)]
        [string]$BaseId,
        
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    $subject = Get-SafeStringValue -Value $EmailData.Subject -DefaultValue "No Subject"
    $senderName = Get-SafeStringValue -Value $EmailData.Sender.Name -DefaultValue "Unknown Sender"
    $senderEmail = Get-SafeStringValue -Value $EmailData.Sender.Email -DefaultValue "unknown@unknown.com"
    
    $keywords = Get-SearchKeywords -EmailData $EmailData -Content $Content
    
    return @{
        id = $BaseId
        DocumentType = "Email"
        Subject = $subject
        Sender_Name = $senderName
        Sender_Email = $senderEmail
        Recipients_To = Get-RecipientList -Recipients $EmailData.Recipients.To
        Recipients_CC = Get-RecipientList -Recipients $EmailData.Recipients.CC
        Recipients_BCC = Get-RecipientList -Recipients $EmailData.Recipients.BCC
        SentDate = Get-SafeStringValue -Value $EmailData.Sent -DefaultValue (Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ")
        ReceivedDate = Get-SafeStringValue -Value $EmailData.Received -DefaultValue (Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ")
        HasAttachments = if ($EmailData.Attachments) { $EmailData.Attachments.Count -gt 0 } else { $false }
        AttachmentCount = if ($EmailData.Attachments) { $EmailData.Attachments.Count } else { 0 }
        MessageSize = if ($EmailData.Size) { [int]$EmailData.Size } else { 0 }
        Importance = Get-SafeStringValue -Value $EmailData.Importance -DefaultValue "Normal"
        SearchKeywords = $keywords
        ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        AllText = $Content
        FileName = Get-SafeStringValue -Value $EmailData.Metadata.FileName -DefaultValue "unknown.msg"
        ProcessorVersion = "1.0"
    }
}

function Create-ChunkSearchDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Chunk,
        
        [Parameter(Mandatory=$true)]
        [string]$BaseId
    )
    
    $chunkId = "$BaseId-chunk-$($Chunk.ChunkNumber)"
    
    $subject = Get-SafeStringValue -Value $EmailData.Subject -DefaultValue "No Subject"
    $senderName = Get-SafeStringValue -Value $EmailData.Sender.Name -DefaultValue "Unknown Sender"
    $senderEmail = Get-SafeStringValue -Value $EmailData.Sender.Email -DefaultValue "unknown@unknown.com"
    
    $keywords = Get-SearchKeywords -EmailData $EmailData -Content $Chunk.Content
    
    return @{
        id = $chunkId
        DocumentType = "EmailChunk"
        Subject = $subject
        Sender_Name = $senderName
        Sender_Email = $senderEmail
        ChunkContent = $Chunk.Content
        ChunkNumber = [int]$Chunk.ChunkNumber
        TotalChunks = [int]$Chunk.TotalChunks
        ChunkLength = [int]$Chunk.Length
        WordCount = [int]$Chunk.WordCount
        IsFirst = $Chunk.IsFirst
        IsLast = $Chunk.IsLast
        SearchKeywords = $keywords
        ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        AllText = $Chunk.Content
        ParentEmailId = $BaseId
        FileName = Get-SafeStringValue -Value $EmailData.Metadata.FileName -DefaultValue "unknown.msg"
        ProcessorVersion = "1.0"
    }
}

function Create-VectorDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Chunk,
        
        [Parameter(Mandatory=$false)]
        [string]$Content,
        
        [Parameter(Mandatory=$true)]
        [string]$BaseId,
        
        [Parameter(Mandatory=$true)]
        [string]$DatabaseType
    )
    
    $docContent = if ($Chunk) { $Chunk.Content } else { $Content }
    $docId = if ($Chunk) { "$BaseId-chunk-$($Chunk.ChunkNumber)" } else { $BaseId }
    $docType = if ($Chunk) { "EmailChunk" } else { "Email" }
    
    $document = @{
        id = $docId
        type = $docType
        content = $docContent
        metadata = @{
            subject = Get-SafeStringValue -Value $EmailData.Subject -DefaultValue "No Subject"
            sender_name = Get-SafeStringValue -Value $EmailData.Sender.Name -DefaultValue "Unknown Sender"
            sender_email = Get-SafeStringValue -Value $EmailData.Sender.Email -DefaultValue "unknown@unknown.com"
            processed_at = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            file_name = Get-SafeStringValue -Value $EmailData.Metadata.FileName -DefaultValue "unknown.msg"
        }
    }
    
    if ($Chunk) {
        $document.metadata.chunk_number = [int]$Chunk.ChunkNumber
        $document.metadata.total_chunks = [int]$Chunk.TotalChunks
        $document.metadata.chunk_length = [int]$Chunk.Length
        $document.metadata.word_count = [int]$Chunk.WordCount
        $document.metadata.parent_email_id = $BaseId
    }
    
    switch ($DatabaseType) {
        "Pinecone" {
            $flatMetadata = @{}
            foreach ($key in $document.metadata.Keys) {
                $flatMetadata[$key] = $document.metadata[$key]
            }
            $document.metadata = $flatMetadata
        }
        
        "Weaviate" {
            $document.className = if ($Chunk) { "EmailChunk" } else { "Email" }
        }
        
        "Qdrant" {
            $document.payload = $document.metadata
            $document.Remove("metadata")
        }
    }
    
    return $document
}

function Get-SafeId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$InputString,
        
        [Parameter(Mandatory=$false)]
        [string]$Prefix = "doc"
    )
    
    if ([string]::IsNullOrWhiteSpace($InputString)) {
        return "$Prefix-$(New-Guid)"
    }
    
    $cleaned = $InputString -replace '[^\w\-_.]', '-'
    $cleaned = $cleaned -replace '-+', '-'
    $cleaned = $cleaned.Trim('-')
    
    if ($cleaned.Length -gt 50) {
        $cleaned = $cleaned.Substring(0, 50).TrimEnd('-')
    }
    
    if ([string]::IsNullOrWhiteSpace($cleaned)) {
        return "$Prefix-$(New-Guid)"
    }
    
    return "$Prefix-$cleaned"
}

function Get-SafeStringValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        $Value,
        
        [Parameter(Mandatory=$true)]
        [string]$DefaultValue
    )
    
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace($Value.ToString())) {
        return $DefaultValue
    }
    
    return $Value.ToString().Trim()
}

function Get-RecipientList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        $Recipients
    )
    
    if (-not $Recipients -or $Recipients.Count -eq 0) {
        return @()
    }
    
    $recipientList = @()
    foreach ($recipient in $Recipients) {
        if ($recipient.Email) {
            $recipientList += $recipient.Email
        } elseif ($recipient.Name) {
            $recipientList += $recipient.Name
        }
    }
    
    return $recipientList
}

function Get-SearchKeywords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData,
        
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    $keywords = @()
    
    if ($EmailData.Subject) {
        $subjectWords = ($EmailData.Subject -split '\s+' | 
                        Where-Object { $_.Length -gt 3 -and $_ -notmatch '^\d+ } | 
                        ForEach-Object { $_.ToLower().Trim('.,!?:;()[]{}') } |
                        Where-Object { $_ -ne "" })
        $keywords += $subjectWords
    }
    
    if ($EmailData.Sender -and $EmailData.Sender.Name) {
        $senderWords = ($EmailData.Sender.Name -split '\s+' | 
                       Where-Object { $_.Length -gt 2 } |
                       ForEach-Object { $_.ToLower().Trim('.,!?:;()[]{}') } |
                       Where-Object { $_ -ne "" })
        $keywords += $senderWords
    }
    
    if ($Content -and $Content.Length -gt 20) {
        $contentWords = ($Content -split '\s+' | 
                        Where-Object { $_.Length -gt 4 -and $_ -notmatch '^\d+ -and $_ -notmatch '^https?:' } | 
                        ForEach-Object { $_.ToLower().Trim('.,!?:;()[]{}') } |
                        Where-Object { $_ -ne "" })
        
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
                   Select-Object -First 15 | 
                   ForEach-Object { $_.Key }
        
        $keywords += $topWords
    }
    
    $stopWords = @('the', 'and', 'that', 'have', 'for', 'not', 'with', 'you', 'this', 'but', 'his', 'from', 'they', 'she', 'her', 'been', 'than', 'its', 'who', 'oil', 'sit', 'now', 'find', 'long', 'down', 'day', 'did', 'get', 'has', 'him', 'had', 'let', 'put', 'say', 'set', 'sun', 'try')
    
    $uniqueKeywords = $keywords | 
                     Where-Object { $_ -notin $stopWords -and $_.Length -gt 2 } | 
                     Sort-Object -Unique |
                     Select-Object -First 25
    
    if (-not $uniqueKeywords) {
        return @()
    }
    
    return @($uniqueKeywords)
}

function Get-BestEmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData
    )
    
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
        $text = $EmailData.HTMLBody -replace '<[^>]*>', ''
        $text = $text -replace '&nbsp;', ' '
        $text = $text -replace '&amp;', '&'
        $text = $text -replace '&lt;', '<'
        $text = $text -replace '&gt;', '>'
        return $text.Trim()
    }
    
    return ""
}

function Optimize-SearchDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Document
    )
    
    $optimized = $Document.Clone()
    
    if ($optimized.SearchKeywords) {
        if ($optimized.SearchKeywords -isnot [array]) {
            $optimized.SearchKeywords = @($optimized.SearchKeywords)
        }
        
        $optimized.SearchKeywords = $optimized.SearchKeywords | 
                                   ForEach-Object { 
                                       if ($_ -is [string] -and $_.Length -le 100) { 
                                           $_ 
                                       } 
                                   } | 
                                   Where-Object { $_ -ne $null }
    } else {
        $optimized.SearchKeywords = @()
    }
    
    $textFields = @('Subject', 'ChunkContent', 'AllText')
    foreach ($field in $textFields) {
        if ($optimized.ContainsKey($field) -and $optimized.$field) {
            if ($optimized.$field.Length -gt 32766) {
                $optimized.$field = $optimized.$field.Substring(0, 32760) + "..."
                $optimized["${field}_Truncated"] = $true
            }
        }
    }
    
    $numericFields = @('ChunkNumber', 'TotalChunks', 'ChunkLength', 'WordCount', 'AttachmentCount', 'MessageSize')
    foreach ($field in $numericFields) {
        if ($optimized.ContainsKey($field) -and $optimized.$field -ne $null) {
            try {
                $optimized.$field = [int]$optimized.$field
            } catch {
                $optimized.$field = 0
            }
        }
    }
    
    $booleanFields = @('HasAttachments', 'IsFirst', 'IsLast', 'IsTruncated')
    foreach ($field in $booleanFields) {
        if ($optimized.ContainsKey($field) -and $optimized.$field -ne $null) {
            try {
                $optimized.$field = [bool]$optimized.$field
            } catch {
                $optimized.$field = $false
            }
        }
    }
    
    return $optimized
}

Write-Verbose "AzureFlattener module loaded successfully with SearchKeywords data type fix"