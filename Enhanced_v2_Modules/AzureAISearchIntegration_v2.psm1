# AzureAISearchIntegration_v2.psm1 - Azure AI Search Integration for Email RAG Pipeline
# Complete integration module for indexing, searching, and managing email documents

Export-ModuleMember -Function @(
    'Initialize-AzureSearchService',
    'New-EmailSearchIndex',
    'Add-EmailDocuments',
    'Search-EmailDocuments',
    'Search-EmailHybrid',
    'Get-SearchStatistics',
    'Update-SearchIndex',
    'Test-SearchConnection'
)

function Initialize-AzureSearchService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
        
        [Parameter(Mandatory=$true)]
        [string]$ApiKey,
        
        [Parameter(Mandatory=$false)]
        [string]$ApiVersion = "2023-11-01",
        
        [Parameter(Mandatory=$false)]
        [hashtable]$OpenAIConfig = @{}
    )
    
    try {
        Write-Verbose "Initializing Azure AI Search service connection..."
        
        $searchConfig = @{
            ServiceName = $ServiceName
            ServiceUrl = "https://$ServiceName.search.windows.net"
            ApiKey = $ApiKey
            ApiVersion = $ApiVersion
            Headers = @{
                'api-key' = $ApiKey
                'Content-Type' = 'application/json'
            }
            OpenAI = $OpenAIConfig
            InitializedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
        # Test connection
        $connectionTest = Test-SearchConnection -Config $searchConfig
        if (-not $connectionTest.IsConnected) {
            throw "Failed to connect to Azure AI Search service: $($connectionTest.Error)"
        }
        
        Write-Verbose "Successfully initialized Azure AI Search service: $ServiceName"
        return $searchConfig
        
    } catch {
        Write-Error "Failed to initialize Azure AI Search service: $($_.Exception.Message)"
        throw
    }
}

function New-EmailSearchIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [string]$SchemaPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$RecreateIfExists = $false
    )
    
    try {
        Write-Verbose "Creating Azure AI Search index: $IndexName"
        
        $indexUrl = "$($SearchConfig.ServiceUrl)/indexes/$IndexName"
        
        # Check if index exists
        $existsResponse = try {
            Invoke-RestMethod -Uri "$indexUrl?api-version=$($SearchConfig.ApiVersion)" -Headers $SearchConfig.Headers -Method GET
            $true
        } catch {
            $false
        }
        
        if ($existsResponse -and -not $RecreateIfExists) {
            Write-Verbose "Index $IndexName already exists"
            return @{
                IndexName = $IndexName
                Status = "AlreadyExists"
                Url = $indexUrl
            }
        }
        
        # Delete existing index if recreating
        if ($existsResponse -and $RecreateIfExists) {
            Write-Verbose "Deleting existing index: $IndexName"
            try {
                Invoke-RestMethod -Uri "$indexUrl?api-version=$($SearchConfig.ApiVersion)" -Headers $SearchConfig.Headers -Method DELETE
                Start-Sleep -Seconds 2
            } catch {
                Write-Warning "Failed to delete existing index: $($_.Exception.Message)"
            }
        }
        
        # Load schema
        $indexSchema = if ($SchemaPath -and (Test-Path $SchemaPath)) {
            Get-Content $SchemaPath -Raw | ConvertFrom-Json
        } else {
            Get-DefaultEmailIndexSchema -IndexName $IndexName -OpenAIConfig $SearchConfig.OpenAI
        }
        
        # Create index
        $createUrl = "$($SearchConfig.ServiceUrl)/indexes?api-version=$($SearchConfig.ApiVersion)"
        $schemaJson = $indexSchema | ConvertTo-Json -Depth 10
        
        $response = Invoke-RestMethod -Uri $createUrl -Headers $SearchConfig.Headers -Method POST -Body $schemaJson
        
        Write-Verbose "Successfully created Azure AI Search index: $IndexName"
        return @{
            IndexName = $IndexName
            Status = "Created"
            Url = $indexUrl
            Schema = $indexSchema
            Response = $response
        }
        
    } catch {
        Write-Error "Failed to create Azure AI Search index: $($_.Exception.Message)"
        throw
    }
}

function Add-EmailDocuments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [array]$Documents,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [int]$BatchSize = 100,
        
        [Parameter(Mandatory=$false)]
        [switch]$GenerateEmbeddings = $true
    )
    
    try {
        Write-Verbose "Adding $($Documents.Count) documents to Azure AI Search index: $IndexName"
        
        $indexUrl = "$($SearchConfig.ServiceUrl)/indexes/$IndexName/docs/index"
        $totalProcessed = 0
        $batches = @()
        
        # Process documents in batches
        for ($i = 0; $i -lt $Documents.Count; $i += $BatchSize) {
            $endIndex = [Math]::Min($i + $BatchSize - 1, $Documents.Count - 1)
            $batch = $Documents[$i..$endIndex]
            
            Write-Verbose "Processing batch $([Math]::Floor($i / $BatchSize) + 1): Documents $($i + 1) to $($endIndex + 1)"
            
            # Prepare documents for indexing
            $indexDocuments = @()
            foreach ($doc in $batch) {
                $indexDoc = Convert-ToAzureSearchDocument -Document $doc -GenerateEmbeddings:$GenerateEmbeddings -OpenAIConfig $SearchConfig.OpenAI
                $indexDocuments += $indexDoc
            }
            
            # Create batch request
            $batchRequest = @{
                value = $indexDocuments | ForEach-Object {
                    @{
                        "@search.action" = "upload"
                    } + $_
                }
            }
            
            $requestJson = $batchRequest | ConvertTo-Json -Depth 10
            
            # Send batch to Azure AI Search
            $response = Invoke-RestMethod -Uri "$indexUrl?api-version=$($SearchConfig.ApiVersion)" -Headers $SearchConfig.Headers -Method POST -Body $requestJson
            
            $batchResult = @{
                BatchNumber = [Math]::Floor($i / $BatchSize) + 1
                DocumentsInBatch = $batch.Count
                Response = $response
                ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            }
            
            $batches += $batchResult
            $totalProcessed += $batch.Count
            
            Write-Verbose "Batch $($batchResult.BatchNumber) completed: $($batch.Count) documents processed"
            
            # Small delay between batches to avoid throttling
            if ($i + $BatchSize -lt $Documents.Count) {
                Start-Sleep -Milliseconds 500
            }
        }
        
        Write-Verbose "Successfully added $totalProcessed documents to Azure AI Search index"
        
        return @{
            TotalDocuments = $Documents.Count
            ProcessedDocuments = $totalProcessed
            BatchCount = $batches.Count
            BatchResults = $batches
            IndexName = $IndexName
            CompletedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to add documents to Azure AI Search: $($_.Exception.Message)"
        throw
    }
}

function Search-EmailDocuments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$Query,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [int]$Top = 10,
        
        [Parameter(Mandatory=$false)]
        [int]$Skip = 0,
        
        [Parameter(Mandatory=$false)]
        [string]$Filter,
        
        [Parameter(Mandatory=$false)]
        [string[]]$SearchFields,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Select,
        
        [Parameter(Mandatory=$false)]
        [string]$OrderBy,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeTotalCount = $true,
        
        [Parameter(Mandatory=$false)]
        [string]$ScoringProfile = "email-relevance-profile"
    )
    
    try {
        Write-Verbose "Searching Azure AI Search index: $IndexName for query: '$Query'"
        
        $searchUrl = "$($SearchConfig.ServiceUrl)/indexes/$IndexName/docs/search"
        
        # Build search request
        $searchRequest = @{
            search = $Query
            top = $Top
            skip = $Skip
            includeTotalCount = $IncludeTotalCount.IsPresent
        }
        
        if ($Filter) { $searchRequest.filter = $Filter }
        if ($SearchFields) { $searchRequest.searchFields = $SearchFields -join "," }
        if ($Select) { $searchRequest.select = $Select -join "," }
        if ($OrderBy) { $searchRequest.orderby = $OrderBy }
        if ($ScoringProfile) { $searchRequest.scoringProfile = $ScoringProfile }
        
        $requestJson = $searchRequest | ConvertTo-Json -Depth 5
        
        # Execute search
        $response = Invoke-RestMethod -Uri "$searchUrl?api-version=$($SearchConfig.ApiVersion)" -Headers $SearchConfig.Headers -Method POST -Body $requestJson
        
        $searchResult = @{
            Query = $Query
            TotalCount = if ($response.'@odata.count') { $response.'@odata.count' } else { $response.value.Count }
            Results = $response.value
            SearchDuration = if ($response.'@search.facets') { "Available in response" } else { "Not available" }
            ExecutedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            IndexName = $IndexName
        }
        
        Write-Verbose "Search completed: $($searchResult.TotalCount) total results, returning top $($searchResult.Results.Count)"
        
        return $searchResult
        
    } catch {
        Write-Error "Failed to search Azure AI Search index: $($_.Exception.Message)"
        throw
    }
}

function Search-EmailHybrid {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$Query,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [int]$Top = 10,
        
        [Parameter(Mandatory=$false)]
        [string]$Filter,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseSemanticSearch = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$GenerateQueryEmbedding = $true
    )
    
    try {
        Write-Verbose "Performing hybrid search on Azure AI Search index: $IndexName"
        
        $searchUrl = "$($SearchConfig.ServiceUrl)/indexes/$IndexName/docs/search"
        
        # Build hybrid search request
        $searchRequest = @{
            search = $Query
            top = $Top
            includeTotalCount = $true
            queryType = "semantic"
            semanticConfiguration = "email-semantic-config"
            searchFields = "title,content,sender_name,people,organizations,keywords"
            select = "id,title,content,sender_name,sender_email,sent_date,chunk_number,total_chunks,people,organizations,keywords"
        }
        
        if ($Filter) { $searchRequest.filter = $Filter }
        
        # Add vector search if embeddings are configured
        if ($GenerateQueryEmbedding -and $SearchConfig.OpenAI -and $SearchConfig.OpenAI.ApiKey) {
            $queryEmbedding = Get-QueryEmbedding -Query $Query -OpenAIConfig $SearchConfig.OpenAI
            if ($queryEmbedding) {
                $searchRequest.vectors = @(@{
                    value = $queryEmbedding
                    fields = "content_vector"
                    k = $Top
                })
            }
        }
        
        $requestJson = $searchRequest | ConvertTo-Json -Depth 5
        
        # Execute hybrid search
        $response = Invoke-RestMethod -Uri "$searchUrl?api-version=$($SearchConfig.ApiVersion)" -Headers $SearchConfig.Headers -Method POST -Body $requestJson
        
        # Process and enrich results
        $enrichedResults = @()
        foreach ($result in $response.value) {
            $enrichedResult = $result.PSObject.Copy()
            
            # Add relevance scoring
            $enrichedResult | Add-Member -NotePropertyName "HybridScore" -NotePropertyValue $result.'@search.score'
            $enrichedResult | Add-Member -NotePropertyName "SearchType" -NotePropertyValue "Hybrid"
            
            # Add semantic search captions if available
            if ($result.'@search.captions') {
                $enrichedResult | Add-Member -NotePropertyName "SemanticCaptions" -NotePropertyValue $result.'@search.captions'
            }
            
            $enrichedResults += $enrichedResult
        }
        
        $hybridResult = @{
            Query = $Query
            SearchType = "Hybrid"
            TotalCount = $response.'@odata.count'
            Results = $enrichedResults
            HasSemanticSearch = $UseSemanticSearch.IsPresent
            HasVectorSearch = ($GenerateQueryEmbedding -and $SearchConfig.OpenAI)
            ExecutedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            IndexName = $IndexName
        }
        
        Write-Verbose "Hybrid search completed: $($hybridResult.TotalCount) total results"
        
        return $hybridResult
        
    } catch {
        Write-Error "Failed to perform hybrid search: $($_.Exception.Message)"
        throw
    }
}

function Convert-ToAzureSearchDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Document,
        
        [Parameter(Mandatory=$false)]
        [switch]$GenerateEmbeddings = $true,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$OpenAIConfig = @{}
    )
    
    try {
        # Map document fields to Azure Search schema
        $azureDoc = @{
            id = $Document.id
            parent_id = $Document.ParentEmailId
            document_type = if ($Document.ChunkType) { $Document.ChunkType } else { "Email" }
            title = $Document.EmailSubject
            content = $Document.Content
            chunk_number = if ($Document.ChunkNumber) { [int]$Document.ChunkNumber } else { 1 }
            total_chunks = if ($Document.TotalChunks) { [int]$Document.TotalChunks } else { 1 }
            sender_name = $Document.SenderName
            sender_email = $Document.Sender_Email
            sent_date = $Document.SentDate
            received_date = $Document.ReceivedDate
            importance = if ($Document.Importance) { $Document.Importance } else { "Normal" }
            has_attachments = if ($Document.HasAttachments) { [bool]$Document.HasAttachments } else { $false }
            processed_at = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            processor_version = "2.0"
            content_quality_score = if ($Document.QualityScore) { [double]$Document.QualityScore / 100 } else { 0.5 }
            word_count = if ($Document.WordCount) { [int]$Document.WordCount } else { 0 }
            language_code = "en"
        }
        
        # Add arrays safely
        $azureDoc.recipients_to = if ($Document.Recipients_To) { @($Document.Recipients_To) } else { @() }
        $azureDoc.recipients_cc = if ($Document.Recipients_CC) { @($Document.Recipients_CC) } else { @() }
        $azureDoc.attachment_types = if ($Document.AttachmentTypes) { @($Document.AttachmentTypes) } else { @() }
        $azureDoc.people = if ($Document.People) { @($Document.People) } else { @() }
        $azureDoc.organizations = if ($Document.Organizations) { @($Document.Organizations) } else { @() }
        $azureDoc.locations = if ($Document.Locations) { @($Document.Locations) } else { @() }
        $azureDoc.urls = if ($Document.URLs) { @($Document.URLs) } else { @() }
        $azureDoc.phone_numbers = if ($Document.PhoneNumbers) { @($Document.PhoneNumbers) } else { @() }
        $azureDoc.keywords = if ($Document.SearchKeywords) { @($Document.SearchKeywords) } else { @() }
        
        # Generate content vector embedding if requested
        if ($GenerateEmbeddings -and $OpenAIConfig -and $OpenAIConfig.ApiKey -and $Document.Content) {
            $embedding = Get-ContentEmbedding -Content $Document.Content -OpenAIConfig $OpenAIConfig
            if ($embedding) {
                $azureDoc.content_vector = $embedding
            }
        }
        
        # Add conversation topic
        $azureDoc.conversation_topic = $Document.ConversationTopic
        
        # Add message ID
        $azureDoc.message_id = $Document.MessageId
        
        # Add sentiment score (placeholder - would integrate with Azure Cognitive Services)
        $azureDoc.sentiment_score = 0.0
        
        return $azureDoc
        
    } catch {
        Write-Warning "Failed to convert document to Azure Search format: $($_.Exception.Message)"
        return $null
    }
}

function Get-ContentEmbedding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$OpenAIConfig
    )
    
    try {
        if (-not $OpenAIConfig.ApiKey -or -not $OpenAIConfig.Endpoint) {
            Write-Verbose "OpenAI configuration incomplete - skipping embedding generation"
            return $null
        }
        
        $embeddingUrl = "$($OpenAIConfig.Endpoint)/openai/deployments/$($OpenAIConfig.EmbeddingModel)/embeddings?api-version=$($OpenAIConfig.ApiVersion)"
        
        $requestBody = @{
            input = $Content
        } | ConvertTo-Json
        
        $headers = @{
            'api-key' = $OpenAIConfig.ApiKey
            'Content-Type' = 'application/json'
        }
        
        $response = Invoke-RestMethod -Uri $embeddingUrl -Headers $headers -Method POST -Body $requestBody
        
        if ($response.data -and $response.data[0].embedding) {
            return $response.data[0].embedding
        }
        
        return $null
        
    } catch {
        Write-Verbose "Failed to generate embedding: $($_.Exception.Message)"
        return $null
    }
}

function Get-QueryEmbedding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Query,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$OpenAIConfig
    )
    
    return Get-ContentEmbedding -Content $Query -OpenAIConfig $OpenAIConfig
}

function Get-DefaultEmailIndexSchema {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$IndexName,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$OpenAIConfig = @{}
    )
    
    # Load and customize the schema from the JSON file
    $schemaPath = Join-Path (Split-Path $PSScriptRoot) "AzureAISearchSchema_v2.json"
    
    if (Test-Path $schemaPath) {
        $schema = Get-Content $schemaPath -Raw | ConvertFrom-Json
        $schema.name = $IndexName
        
        # Update OpenAI configuration if provided
        if ($OpenAIConfig.Endpoint -and $OpenAIConfig.ApiKey) {
            $schema.vectorSearch.vectorizers[0].azureOpenAIParameters.resourceUri = $OpenAIConfig.Endpoint
            $schema.vectorSearch.vectorizers[0].azureOpenAIParameters.apiKey = $OpenAIConfig.ApiKey
            if ($OpenAIConfig.EmbeddingModel) {
                $schema.vectorSearch.vectorizers[0].azureOpenAIParameters.deploymentId = $OpenAIConfig.EmbeddingModel
            }
        }
        
        return $schema
    }
    
    # Fallback basic schema if file not found
    return @{
        name = $IndexName
        fields = @(
            @{
                name = "id"
                type = "Edm.String"
                key = $true
                searchable = $false
                filterable = $true
                retrievable = $true
            },
            @{
                name = "title"
                type = "Edm.String"
                searchable = $true
                filterable = $false
                retrievable = $true
            },
            @{
                name = "content"
                type = "Edm.String"
                searchable = $true
                filterable = $false
                retrievable = $true
            }
        )
    }
}

function Get-SearchStatistics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index"
    )
    
    try {
        $statsUrl = "$($SearchConfig.ServiceUrl)/indexes/$IndexName/stats"
        $response = Invoke-RestMethod -Uri "$statsUrl?api-version=$($SearchConfig.ApiVersion)" -Headers $SearchConfig.Headers -Method GET
        
        return @{
            IndexName = $IndexName
            DocumentCount = $response.documentCount
            StorageSize = $response.storageSize
            RetrievedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to get search statistics: $($_.Exception.Message)"
        throw
    }
}

function Test-SearchConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        $testUrl = "$($Config.ServiceUrl)/servicestats"
        $response = Invoke-RestMethod -Uri "$testUrl?api-version=$($Config.ApiVersion)" -Headers $Config.Headers -Method GET
        
        return @{
            IsConnected = $true
            ServiceName = $Config.ServiceName
            ServiceStats = $response
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        return @{
            IsConnected = $false
            Error = $_.Exception.Message
            TestedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
    }
}

Write-Verbose "AzureAISearchIntegration_v2 module loaded successfully"