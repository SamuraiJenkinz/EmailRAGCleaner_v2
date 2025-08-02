# RAGConfigManager_v2.psm1 - Configuration Management for Email RAG System
# Enhanced for PowerShell 7 with modern syntax and robust error handling

using namespace System.Collections.Generic

Export-ModuleMember -Function @(
    'New-RAGConfiguration',
    'Import-RAGConfiguration', 
    'Export-RAGConfiguration',
    'Update-RAGConfiguration',
    'Test-RAGConfiguration',
    'Get-DefaultConfiguration',
    'Merge-Configurations',
    'Validate-ConfigurationSchema'
)

function New-RAGConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigurationName,
        
        [Parameter(Mandatory=$true)]
        [string]$AzureSearchServiceName,
        
        [Parameter(Mandatory=$true)]
        [string]$AzureSearchApiKey,
        
        [Parameter(Mandatory=$false)]
        [string]$OpenAIEndpoint,
        
        [Parameter(Mandatory=$false)]
        [string]$OpenAIApiKey,
        
        [Parameter(Mandatory=$false)]
        [string]$EmbeddingModel = "text-embedding-ada-002",
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [hashtable]$ProcessingSettings = @{},
        
        [Parameter(Mandatory=$false)]
        [hashtable]$SearchSettings = @{},
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath
    )
    
    try {
        Write-Verbose "Creating new RAG configuration: $ConfigurationName"
        
        $configuration = @{
            Metadata = @{
                Name = $ConfigurationName
                Version = "2.0"
                CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
                CreatedBy = $env:USERNAME
                Description = "Email RAG system configuration for Azure AI Search integration"
            }
            
            AzureSearch = @{
                ServiceName = $AzureSearchServiceName
                ServiceUrl = "https://$AzureSearchServiceName.search.windows.net"
                ApiKey = $AzureSearchApiKey
                ApiVersion = "2023-11-01"
                IndexName = $IndexName
                Timeout = 120
                RetryAttempts = 3
                BatchSize = 50
            }
            
            OpenAI = @{
                Endpoint = $OpenAIEndpoint ?? ""
                ApiKey = $OpenAIApiKey ?? ""
                ApiVersion = "2023-05-15"
                EmbeddingModel = $EmbeddingModel
                MaxTokens = 8191
                Timeout = 60
                RetryAttempts = 3
                EmbeddingDimensions = 1536
                Enabled = (-not [string]::IsNullOrEmpty($OpenAIEndpoint) -and -not [string]::IsNullOrEmpty($OpenAIApiKey))
            }
            
            Processing = Get-DefaultProcessingSettings -CustomSettings $ProcessingSettings
            Search = Get-DefaultSearchSettings -CustomSettings $SearchSettings
            
            Logging = @{
                Level = "Information"
                EnableFileLogging = $true
                LogPath = "Logs"
                MaxLogFiles = 10
                MaxLogSizeMB = 50
            }
            
            Performance = @{
                MaxConcurrentProcessing = 5
                ChunkProcessingBatchSize = 100
                IndexingBatchSize = 50
                MemoryThresholdMB = 1024
                EnableProgressReporting = $true
            }
        }
        
        # Validate configuration
        $validationResult = Validate-ConfigurationSchema -Configuration $configuration
        if (-not $validationResult.IsValid) {
            throw "Configuration validation failed: $($validationResult.Errors -join ', ')"
        }
        
        # Export configuration if path specified
        if ($OutputPath) {
            Export-RAGConfiguration -Configuration $configuration -OutputPath $OutputPath
            Write-Verbose "Configuration exported to: $OutputPath"
        }
        
        Write-Verbose "RAG configuration created successfully: $ConfigurationName"
        return $configuration
        
    } catch {
        Write-Error "Failed to create RAG configuration: $($_.Exception.Message)"
        throw
    }
}

function Import-RAGConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigurationPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$ValidateOnImport = $true,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$EnvironmentOverrides = @{}
    )
    
    try {
        Write-Verbose "Importing RAG configuration from: $ConfigurationPath"
        
        if (-not (Test-Path $ConfigurationPath)) {
            throw "Configuration file not found: $ConfigurationPath"
        }
        
        # Read and parse configuration
        $configContent = Get-Content $ConfigurationPath -Raw
        $configuration = $configContent | ConvertFrom-Json -AsHashtable
        
        if (-not $configuration) {
            throw "Failed to parse configuration file"
        }
        
        # Apply environment overrides using PowerShell 7 syntax
        if ($EnvironmentOverrides.Count -gt 0) {
            $configuration = Merge-Configurations -BaseConfiguration $configuration -OverrideConfiguration $EnvironmentOverrides
            Write-Verbose "Applied $($EnvironmentOverrides.Count) environment overrides"
        }
        
        # Validate if requested
        if ($ValidateOnImport) {
            $validationResult = Validate-ConfigurationSchema -Configuration $configuration
            if (-not $validationResult.IsValid) {
                throw "Configuration validation failed: $($validationResult.Errors -join ', ')"
            }
        }
        
        # Add import metadata
        $configuration.Metadata ??= @{}
        $configuration.Metadata.ImportedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        $configuration.Metadata.ImportedBy = $env:USERNAME
        $configuration.Metadata.ImportedFrom = $ConfigurationPath
        
        Write-Verbose "RAG configuration imported successfully: $($configuration.Metadata.Name)"
        return $configuration
        
    } catch {
        Write-Error "Failed to import RAG configuration: $($_.Exception.Message)"
        throw
    }
}

function Export-RAGConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeSecrets = $false,
        
        [Parameter(Mandatory=$false)]
        [switch]$PrettyFormat = $true
    )
    
    try {
        Write-Verbose "Exporting RAG configuration to: $OutputPath"
        
        # Deep copy configuration
        $exportConfig = $Configuration | ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable
        
        # Mask secrets if not including them
        if (-not $IncludeSecrets) {
            $exportConfig.AzureSearch?.ApiKey && ($exportConfig.AzureSearch.ApiKey = "***MASKED***")
            $exportConfig.OpenAI?.ApiKey && ($exportConfig.OpenAI.ApiKey = "***MASKED***")
        }
        
        # Add export metadata
        $exportConfig.Metadata ??= @{}
        $exportConfig.Metadata.ExportedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        $exportConfig.Metadata.ExportedBy = $env:USERNAME
        $exportConfig.Metadata.SecretsIncluded = $IncludeSecrets.IsPresent
        
        # Ensure output directory exists
        $outputDir = Split-Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # Convert to JSON and export
        $jsonOptions = @{ Depth = 10 }
        $configJson = $PrettyFormat ? 
            ($exportConfig | ConvertTo-Json @jsonOptions) :
            ($exportConfig | ConvertTo-Json @jsonOptions -Compress)
        
        $configJson | Out-File -FilePath $OutputPath -Encoding UTF8
        
        Write-Verbose "RAG configuration exported successfully to: $OutputPath"
        
        return @{
            OutputPath = $OutputPath
            ConfigurationName = $exportConfig.Metadata.Name
            SecretsIncluded = $IncludeSecrets.IsPresent
            FileSize = (Get-Item $OutputPath).Length
            ExportedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to export RAG configuration: $($_.Exception.Message)"
        throw
    }
}

function Update-RAGConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Updates = @{},
        
        [Parameter(Mandatory=$false)]
        [string]$UpdateDescription = "Configuration update",
        
        [Parameter(Mandatory=$false)]
        [switch]$ValidateAfterUpdate = $true
    )
    
    try {
        Write-Verbose "Updating RAG configuration: $UpdateDescription"
        
        # Create backup using PowerShell 7 pipeline
        $originalConfig = $Configuration | ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable
        
        # Apply updates recursively
        $updatedConfig = Merge-Configurations -BaseConfiguration $Configuration -OverrideConfiguration $Updates
        
        # Update metadata
        $updatedConfig.Metadata ??= @{}
        $updatedConfig.Metadata.UpdateHistory ??= @()
        
        $updatedConfig.Metadata.UpdateHistory += @{
            UpdatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            UpdatedBy = $env:USERNAME
            Description = $UpdateDescription
            ChangesApplied = $Updates.Keys.Count
        }
        
        $updatedConfig.Metadata.LastModified = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        $updatedConfig.Metadata.Version = Get-NextVersion -CurrentVersion ($updatedConfig.Metadata.Version ?? "1.0")
        
        # Validate updated configuration
        if ($ValidateAfterUpdate) {
            $validationResult = Validate-ConfigurationSchema -Configuration $updatedConfig
            if (-not $validationResult.IsValid) {
                throw "Updated configuration validation failed: $($validationResult.Errors -join ', ')"
            }
        }
        
        Write-Verbose "RAG configuration updated successfully. New version: $($updatedConfig.Metadata.Version)"
        
        return @{
            Configuration = $updatedConfig
            UpdateSummary = @{
                ChangesApplied = $Updates.Keys.Count
                NewVersion = $updatedConfig.Metadata.Version
                UpdateDescription = $UpdateDescription
                ValidationPassed = $ValidateAfterUpdate ? $validationResult.IsValid : $null
            }
            OriginalConfiguration = $originalConfig
        }
        
    } catch {
        Write-Error "Failed to update RAG configuration: $($_.Exception.Message)"
        throw
    }
}

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
        
        # Test 2: Azure Search connection using modern PowerShell 7 features
        if ($TestConnections -and $Configuration.AzureSearch) {
            Write-Host "2. Testing Azure Search connection..." -ForegroundColor Cyan
            $testResults.TestsRun++
            
            try {
                $searchConfig = $Configuration.AzureSearch
                $testUrl = "$($searchConfig.ServiceUrl)/servicestats?api-version=$($searchConfig.ApiVersion)"
                $headers = @{
                    'api-key' = $searchConfig.ApiKey
                    'Content-Type' = 'application/json'
                }
                
                $response = Invoke-RestMethod -Uri $testUrl -Headers $headers -Method GET -TimeoutSec 30
                
                Write-Host "   ✓ Azure Search connection successful" -ForegroundColor Green
                $testResults.TestsPassed++
                $testResults.Details.AzureSearchConnection = @{ 
                    Status = "Passed" 
                    ServiceName = $searchConfig.ServiceName
                    Response = $response
                }
            } catch {
                Write-Host "   ✗ Azure Search connection failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.TestsFailed++
                $testResults.Details.AzureSearchConnection = @{ Status = "Failed"; Error = $_.Exception.Message }
            }
        }
        
        # Test 3: OpenAI connection (if configured)
        if ($TestConnections -and $Configuration.OpenAI?.Enabled -and $Configuration.OpenAI?.ApiKey) {
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
                    AvailableModels = $response.data?.Count ?? 0
                }
            } catch {
                Write-Host "   ✗ OpenAI connection failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.TestsFailed++
                $testResults.Details.OpenAIConnection = @{ Status = "Failed"; Error = $_.Exception.Message }
            }
        }
        
        # Calculate overall status using PowerShell 7 ternary operators
        $testResults.OverallStatus = $testResults.TestsFailed -eq 0 ? 
            ($testResults.TestsPassed -eq $testResults.TestsRun ? "Passed" : "Warning") : 
            "Failed"
        
        # Display summary
        Write-Host "`nConfiguration Test Summary:" -ForegroundColor Green
        $statusColor = $testResults.OverallStatus -eq "Passed" ? "Green" : 
                      $testResults.OverallStatus -eq "Warning" ? "Yellow" : "Red"
        Write-Host "Overall Status: $($testResults.OverallStatus)" -ForegroundColor $statusColor
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

function Get-DefaultConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ConfigurationName = "Default RAG Configuration"
    )
    
    return @{
        Metadata = @{
            Name = $ConfigurationName
            Version = "2.0"
            CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            Description = "Default configuration template for Email RAG system"
        }
        
        AzureSearch = @{
            ServiceName = ""
            ServiceUrl = ""
            ApiKey = ""
            ApiVersion = "2023-11-01"
            IndexName = "email-rag-index"
            Timeout = 120
            RetryAttempts = 3
            BatchSize = 50
        }
        
        OpenAI = @{
            Endpoint = ""
            ApiKey = ""
            ApiVersion = "2023-05-15"
            EmbeddingModel = "text-embedding-ada-002"
            MaxTokens = 8191
            Timeout = 60
            RetryAttempts = 3
            EmbeddingDimensions = 1536
            Enabled = $false
        }
        
        Processing = Get-DefaultProcessingSettings
        Search = Get-DefaultSearchSettings
        
        Logging = @{
            Level = "Information"
            EnableFileLogging = $true
            LogPath = "Logs"
            MaxLogFiles = 10
            MaxLogSizeMB = 50
        }
        
        Performance = @{
            MaxConcurrentProcessing = 5
            ChunkProcessingBatchSize = 100
            IndexingBatchSize = 50
            MemoryThresholdMB = 1024
            EnableProgressReporting = $true
        }
    }
}

function Get-DefaultProcessingSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [hashtable]$CustomSettings = @{}
    )
    
    $defaultSettings = @{
        Chunking = @{
            TargetTokens = 384
            MinTokens = 128
            MaxTokens = 512
            OverlapTokens = 32
            PreserveStructure = $true
            OptimizeForSearch = $true
        }
        
        ContentCleaning = @{
            RemoveSignatures = $true
            RemoveQuotedText = $false
            NormalizeWhitespace = $true
            ExtractEntities = $true
            OptimizeForRAG = $true
        }
        
        EntityExtraction = @{
            IncludeBusinessEntities = $true
            IncludePersonalInfo = $true
            IncludeTechnicalEntities = $true
            MinConfidenceScore = 0.6
        }
        
        Quality = @{
            MinContentLength = 10
            MaxContentLength = 50000
            MinQualityScore = 0.3
            ValidateEmail = $true
        }
        
        Batch = @{
            BatchSize = 50
            MaxConcurrentBatches = 3
            ContinueOnError = $true
            EnableProgressReporting = $true
        }
    }
    
    return Merge-Configurations -BaseConfiguration $defaultSettings -OverrideConfiguration $CustomSettings
}

function Get-DefaultSearchSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [hashtable]$CustomSettings = @{}
    )
    
    $defaultSettings = @{
        DefaultSearchType = "Hybrid"
        MaxResults = 50
        MinimumScore = 0.0
        EnableSemanticSearch = $true
        EnableVectorSearch = $true
        
        Facets = @(
            "sender_name",
            "sent_date", 
            "has_attachments",
            "importance",
            "document_type"
        )
        
        SortOptions = @(
            "Relevance",
            "Date",
            "Sender"
        )
        
        DefaultScoringProfile = "email-relevance-profile"
        
        QueryExpansion = @{
            EnableSynonyms = $true
            EnableSpellCheck = $true
            ExpandAcronyms = $true
        }
    }
    
    return Merge-Configurations -BaseConfiguration $defaultSettings -OverrideConfiguration $CustomSettings
}

function Merge-Configurations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$BaseConfiguration,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$OverrideConfiguration
    )
    
    # Deep copy using PowerShell 7 features
    $mergedConfig = $BaseConfiguration | ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable
    
    foreach ($key in $OverrideConfiguration.Keys) {
        if ($mergedConfig.ContainsKey($key) -and 
            $mergedConfig[$key] -is [hashtable] -and 
            $OverrideConfiguration[$key] -is [hashtable]) {
            # Recursively merge nested hashtables
            $mergedConfig[$key] = Merge-Configurations -BaseConfiguration $mergedConfig[$key] -OverrideConfiguration $OverrideConfiguration[$key]
        } else {
            # Override or add new value
            $mergedConfig[$key] = $OverrideConfiguration[$key]
        }
    }
    
    return $mergedConfig
}

function Validate-ConfigurationSchema {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Configuration
    )
    
    $validationResult = @{
        IsValid = $true
        Errors = [List[string]]::new()
        Warnings = [List[string]]::new()
    }
    
    try {
        # Required sections
        $requiredSections = @('Metadata', 'AzureSearch', 'Processing')
        foreach ($section in $requiredSections) {
            if (-not $Configuration.ContainsKey($section)) {
                $validationResult.Errors.Add("Missing required section: $section")
                $validationResult.IsValid = $false
            }
        }
        
        # Azure Search validation using null-conditional operators
        if ($Configuration.AzureSearch) {
            $azureSearch = $Configuration.AzureSearch
            
            [string]::IsNullOrEmpty($azureSearch.ServiceName) && $validationResult.Errors.Add("AzureSearch.ServiceName is required")
            [string]::IsNullOrEmpty($azureSearch.ApiKey) && $validationResult.Errors.Add("AzureSearch.ApiKey is required")
            [string]::IsNullOrEmpty($azureSearch.IndexName) && $validationResult.Errors.Add("AzureSearch.IndexName is required")
        }
        
        # OpenAI validation (if enabled)
        if ($Configuration.OpenAI?.Enabled) {
            $openAI = $Configuration.OpenAI
            
            [string]::IsNullOrEmpty($openAI.Endpoint) && $validationResult.Errors.Add("OpenAI.Endpoint is required when OpenAI is enabled")
            [string]::IsNullOrEmpty($openAI.ApiKey) && $validationResult.Errors.Add("OpenAI.ApiKey is required when OpenAI is enabled")
        }
        
        # Processing settings validation
        if ($Configuration.Processing?.Chunking) {
            $chunking = $Configuration.Processing.Chunking
            
            $chunking.TargetTokens -le 0 && $validationResult.Warnings.Add("Processing.Chunking.TargetTokens should be greater than 0")
            $chunking.MaxTokens -lt $chunking.MinTokens && $validationResult.Errors.Add("Processing.Chunking.MaxTokens should be greater than MinTokens")
        }
        
        # Update IsValid based on errors
        $validationResult.IsValid = $validationResult.Errors.Count -eq 0
        
        return $validationResult
        
    } catch {
        $validationResult.IsValid = $false
        $validationResult.Errors.Add("Configuration validation error: $($_.Exception.Message)")
        return $validationResult
    }
}

function Get-NextVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$CurrentVersion = "1.0"
    )
    
    try {
        $versionParts = $CurrentVersion.Split('.')
        $major = [int]$versionParts[0]
        $minor = $versionParts.Length -gt 1 ? [int]$versionParts[1] : 0
        
        return "$major.$($minor + 1)"
    } catch {
        return "1.1"
    }
}

Write-Verbose "RAGConfigManager_v2 (Fixed) module loaded successfully"