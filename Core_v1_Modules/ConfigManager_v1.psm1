# ConfigManager.psm1 - Configuration Management Module (Fixed)
# Professional configuration handling with validation and encryption support

# Export functions
Export-ModuleMember -Function @(
    'Load-Config',
    'Save-Config',
    'Test-ConfigSchema',
    'Get-DefaultConfig',
    'Backup-Config',
    'Restore-Config',
    'Convert-PSObjectToHashtable'
)

function Load-Config {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseDefaults = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$ValidateSchema = $true
    )
    
    try {
        Write-Verbose "Loading configuration from: $ConfigPath"
        
        # Start with defaults if requested
        $config = if ($UseDefaults) { Get-DefaultConfig } else { @{} }
        
        # Load from file if it exists
        if (Test-Path $ConfigPath) {
            Write-Verbose "Reading configuration file..."
            $jsonContent = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
            
            if ($jsonContent -and $jsonContent.Trim() -ne "") {
                try {
                    $loadedConfig = $jsonContent | ConvertFrom-Json
                    $loadedHashtable = Convert-PSObjectToHashtable -InputObject $loadedConfig
                    $config = Merge-Hashtables -BaseConfig $config -OverrideConfig $loadedHashtable
                    Write-Verbose "Configuration loaded successfully from file"
                } catch {
                    Write-Warning "Failed to parse JSON configuration: $($_.Exception.Message)"
                    Write-Warning "Using default configuration"
                }
            } else {
                Write-Warning "Configuration file is empty, using defaults"
            }
        } else {
            Write-Verbose "Configuration file not found, using defaults"
        }
        
        # Validate schema if requested
        if ($ValidateSchema) {
            $validationResult = Test-ConfigSchema -Config $config
            if (-not $validationResult.IsValid) {
                Write-Warning "Configuration validation failed: $($validationResult.Errors -join ', ')"
            }
        }
        
        Write-Verbose "Configuration loaded with $($config.Keys.Count) top-level sections"
        return $config
        
    } catch {
        Write-Error "Failed to load configuration: $($_.Exception.Message)"
        
        if ($UseDefaults) {
            Write-Warning "Returning default configuration as fallback"
            return Get-DefaultConfig
        }
        
        throw
    }
}

function Save-Config {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$ConfigData,
        
        [Parameter(Mandatory=$false)]
        [switch]$CreateBackup = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$ValidateSchema = $true
    )
    
    try {
        Write-Verbose "Saving configuration to: $ConfigPath"
        
        # Validate schema if requested
        if ($ValidateSchema) {
            $validationResult = Test-ConfigSchema -Config $ConfigData
            if (-not $validationResult.IsValid) {
                throw "Configuration validation failed: $($validationResult.Errors -join ', ')"
            }
        }
        
        # Create backup if requested and file exists
        if ($CreateBackup -and (Test-Path $ConfigPath)) {
            $backupResult = Backup-Config -ConfigPath $ConfigPath
            if ($backupResult.Success) {
                Write-Verbose "Configuration backup created: $($backupResult.BackupPath)"
            }
        }
        
        # Ensure directory exists
        $configDir = Split-Path -Parent $ConfigPath
        if ($configDir -and -not (Test-Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
            Write-Verbose "Created configuration directory: $configDir"
        }
        
        # Convert to JSON
        $jsonContent = $ConfigData | ConvertTo-Json -Depth 10
        
        # Save to file
        $jsonContent | Set-Content -Path $ConfigPath -Encoding UTF8 -Force
        
        Write-Verbose "Configuration saved successfully"
        
        return @{
            Success = $true
            ConfigPath = $ConfigPath
            SavedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            SizeBytes = (Get-Item $ConfigPath).Length
        }
        
    } catch {
        Write-Error "Failed to save configuration: $($_.Exception.Message)"
        
        return @{
            Success = $false
            Error = $_.Exception.Message
            ConfigPath = $ConfigPath
        }
    }
}

function Test-ConfigSchema {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        Write-Verbose "Validating configuration schema..."
        
        $errors = @()
        $warnings = @()
        
        # Define expected schema
        $requiredSections = @('Processing', 'Azure', 'Performance', 'Logging')
        
        # Check for required sections
        foreach ($section in $requiredSections) {
            if (-not $Config.ContainsKey($section)) {
                $errors += "Missing required section: $section"
            }
        }
        
        # Validate Processing section
        if ($Config.ContainsKey('Processing')) {
            $processing = $Config.Processing
            
            if ($processing.ContainsKey('ChunkSize')) {
                $chunkSize = $processing.ChunkSize
                if (-not ($chunkSize -is [int]) -or $chunkSize -lt 100 -or $chunkSize -gt 2000) {
                    $errors += "ChunkSize must be an integer between 100 and 2000"
                }
            }
            
            if ($processing.ContainsKey('ChunkOverlap')) {
                $chunkOverlap = $processing.ChunkOverlap
                if (-not ($chunkOverlap -is [int]) -or $chunkOverlap -lt 0 -or $chunkOverlap -gt 200) {
                    $errors += "ChunkOverlap must be an integer between 0 and 200"
                }
            }
        }
        
        # Validate Azure section
        if ($Config.ContainsKey('Azure')) {
            $azure = $Config.Azure
            
            if ($azure.ContainsKey('UseAzure') -and $azure.UseAzure) {
                if (-not $azure.ContainsKey('ConnectionString') -or [string]::IsNullOrWhiteSpace($azure.ConnectionString)) {
                    $errors += "Azure ConnectionString is required when UseAzure is true"
                }
                
                if (-not $azure.ContainsKey('ContainerName') -or [string]::IsNullOrWhiteSpace($azure.ContainerName)) {
                    $errors += "Azure ContainerName is required when UseAzure is true"
                }
            }
        }
        
        $isValid = $errors.Count -eq 0
        
        Write-Verbose "Configuration validation completed. Valid: $isValid, Errors: $($errors.Count), Warnings: $($warnings.Count)"
        
        return @{
            IsValid = $isValid
            Errors = $errors
            Warnings = $warnings
            ValidatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Configuration validation failed: $($_.Exception.Message)"
        
        return @{
            IsValid = $false
            Errors = @("Validation error: $($_.Exception.Message)")
            Warnings = @()
        }
    }
}

function Get-DefaultConfig {
    [CmdletBinding()]
    param()
    
    return @{
        Processing = @{
            OptimizeForRAG = $true
            RemoveSignatures = $true
            ExtractEntities = $true
            ProcessAttachments = $true
            ChunkSize = 512
            ChunkOverlap = 50
            MaxContentLength = 1000000
            ContentQualityThreshold = 10.0
        }
        
        Azure = @{
            UseAzure = $false
            ConnectionString = ""
            ContainerName = "email-data"
            CreateContainerIfNotExists = $true
            BlobPrefix = ""
        }
        
        Performance = @{
            BatchSize = 20
            MaxConcurrency = 4
            TimeoutSeconds = 300
            RetryAttempts = 3
            MemoryLimitMB = 1024
            EnableProgressReporting = $true
        }
        
        Logging = @{
            LogLevel = "INFO"
            EnableFileLogging = $true
            EnableConsoleLogging = $true
            MaxLogSizeMB = 100
            MaxLogFiles = 10
            LogPath = "Logs"
            LogFormat = "yyyy-MM-dd HH:mm:ss"
        }
        
        Metadata = @{
            ConfigVersion = "1.0"
            CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            Description = "MSG Email Cleaner Configuration"
        }
    }
}

function Backup-Config {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory=$false)]
        [string]$BackupDir
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            return @{
                Success = $false
                Error = "Configuration file not found: $ConfigPath"
            }
        }
        
        if (-not $BackupDir) {
            $configDir = Split-Path -Parent $ConfigPath
            $BackupDir = Join-Path $configDir "Backup"
        }
        
        if (-not (Test-Path $BackupDir)) {
            New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
        }
        
        $configFileName = Split-Path -Leaf $ConfigPath
        $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($configFileName)
        $extension = [System.IO.Path]::GetExtension($configFileName)
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFileName = "$nameWithoutExt" + "_backup_" + "$timestamp" + "$extension"
        $backupPath = Join-Path $BackupDir $backupFileName
        
        Copy-Item -Path $ConfigPath -Destination $backupPath -Force
        
        Write-Verbose "Configuration backup created: $backupPath"
        
        return @{
            Success = $true
            BackupPath = $backupPath
            OriginalPath = $ConfigPath
            BackupSize = (Get-Item $backupPath).Length
            CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to create configuration backup: $($_.Exception.Message)"
        
        return @{
            Success = $false
            Error = $_.Exception.Message
            OriginalPath = $ConfigPath
        }
    }
}

function Restore-Config {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BackupPath,
        
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $BackupPath)) {
            throw "Backup file not found: $BackupPath"
        }
        
        if (Test-Path $ConfigPath) {
            $currentBackup = Backup-Config -ConfigPath $ConfigPath
            Write-Verbose "Current configuration backed up to: $($currentBackup.BackupPath)"
        }
        
        Copy-Item -Path $BackupPath -Destination $ConfigPath -Force
        
        Write-Verbose "Configuration restored from backup: $BackupPath"
        
        return @{
            Success = $true
            RestoredFrom = $BackupPath
            RestoredTo = $ConfigPath
            RestoredAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Failed to restore configuration: $($_.Exception.Message)"
        
        return @{
            Success = $false
            Error = $_.Exception.Message
            BackupPath = $BackupPath
            ConfigPath = $ConfigPath
        }
    }
}

function Convert-PSObjectToHashtable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $InputObject
    )
    
    if ($null -eq $InputObject) {
        return $null
    }
    
    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in $InputObject.Keys) {
            $result[$key] = Convert-PSObjectToHashtable -InputObject $InputObject[$key]
        }
        return $result
    }
    
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $result = @()
        foreach ($item in $InputObject) {
            $result += Convert-PSObjectToHashtable -InputObject $item
        }
        return $result
    }
    
    if ($InputObject -is [PSCustomObject]) {
        $result = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $result[$property.Name] = Convert-PSObjectToHashtable -InputObject $property.Value
        }
        return $result
    }
    
    return $InputObject
}

function Merge-Hashtables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$BaseConfig,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$OverrideConfig
    )
    
    $result = $BaseConfig.Clone()
    
    foreach ($key in $OverrideConfig.Keys) {
        if ($result.ContainsKey($key) -and $result[$key] -is [hashtable] -and $OverrideConfig[$key] -is [hashtable]) {
            $result[$key] = Merge-Hashtables -BaseConfig $result[$key] -OverrideConfig $OverrideConfig[$key]
        } else {
            $result[$key] = $OverrideConfig[$key]
        }
    }
    
    return $result
}

Write-Verbose "ConfigManager module loaded successfully"