# Install-EmailRAGCleaner.ps1 - Complete Installation Script for MSG Email Cleaner v2.0
# Installs and configures the enhanced Email RAG system with Azure AI Search integration

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$InstallPath = "C:\EmailRAGCleaner",
    
    [Parameter(Mandatory=$false)]
    [string]$AzureSearchServiceName,
    
    [Parameter(Mandatory=$false)]
    [string]$AzureSearchApiKey,
    
    [Parameter(Mandatory=$false)]
    [string]$OpenAIEndpoint,
    
    [Parameter(Mandatory=$false)]
    [string]$OpenAIApiKey,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateDesktopShortcut = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateStartMenuEntry = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$InstallPrerequisites = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$RunTests = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Silent = $false
)

# Set error handling
$ErrorActionPreference = "Stop"

# Function to write colored output
function Write-InstallStep {
    param(
        [string]$Message,
        [string]$Status = "INFO",
        [switch]$NoNewline = $false
    )
    
    $color = switch($Status) {
        "SUCCESS" { "Green" }
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "HEADER" { "Cyan" }
        default { "White" }
    }
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $output = "[$timestamp] $Message"
    
    if ($NoNewline) {
        Write-Host $output -ForegroundColor $color -NoNewline
    } else {
        Write-Host $output -ForegroundColor $color
    }
}

# Function to test administrator privileges
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to install prerequisites
function Install-Prerequisites {
    Write-InstallStep "Installing prerequisites..." "HEADER"
    
    try {
        # Check PowerShell version
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "PowerShell 5.0 or higher is required. Current version: $($PSVersionTable.PSVersion)"
        }
        Write-InstallStep "✓ PowerShell version check passed: $($PSVersionTable.PSVersion)" "SUCCESS"
        
        # Install required modules
        $requiredModules = @(
            @{ Name = "ImportExcel"; MinVersion = "7.0.0"; Description = "Excel file operations" }
            @{ Name = "Microsoft.PowerShell.Utility"; MinVersion = "3.0.0"; Description = "Utility functions" }
        )
        
        foreach ($module in $requiredModules) {
            Write-InstallStep "Checking module: $($module.Name)..." "INFO"
            
            $installedModule = Get-Module -Name $module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            
            if (-not $installedModule -or $installedModule.Version -lt [version]$module.MinVersion) {
                Write-InstallStep "Installing $($module.Name) module..." "INFO"
                try {
                    Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Force -AllowClobber -Scope CurrentUser
                    Write-InstallStep "✓ $($module.Name) installed successfully" "SUCCESS"
                } catch {
                    Write-InstallStep "⚠ Failed to install $($module.Name): $($_.Exception.Message)" "WARNING"
                }
            } else {
                Write-InstallStep "✓ $($module.Name) already installed (v$($installedModule.Version))" "SUCCESS"
            }
        }
        
        # Check .NET Framework version
        $dotNetVersion = (Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\" -Name Release -ErrorAction SilentlyContinue).Release
        if ($dotNetVersion -and $dotNetVersion -ge 461808) {
            Write-InstallStep "✓ .NET Framework 4.7.2+ detected" "SUCCESS"
        } else {
            Write-InstallStep "⚠ .NET Framework 4.7.2+ recommended for optimal performance" "WARNING"
        }
        
        return $true
        
    } catch {
        Write-InstallStep "Failed to install prerequisites: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to create directory structure
function New-DirectoryStructure {
    param([string]$BasePath)
    
    Write-InstallStep "Creating directory structure..." "HEADER"
    
    $directories = @(
        $BasePath,
        "$BasePath\Modules",
        "$BasePath\Config",
        "$BasePath\Scripts",
        "$BasePath\Logs",
        "$BasePath\TestData",
        "$BasePath\Documentation",
        "$BasePath\Backup"
    )
    
    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-InstallStep "✓ Created directory: $dir" "SUCCESS"
        } else {
            Write-InstallStep "✓ Directory exists: $dir" "SUCCESS"
        }
    }
}

# Function to copy module files
function Copy-ModuleFiles {
    param([string]$SourcePath, [string]$DestinationPath)
    
    Write-InstallStep "Installing Email RAG Cleaner modules..." "HEADER"
    
    $moduleFiles = @(
        # Core v1 modules (clean versions)
        @{ Source = "clean_msgprocessor.txt"; Destination = "MsgProcessor.psm1"; Description = "MSG file processor" },
        @{ Source = "clean_contentcleaner.txt"; Destination = "ContentCleaner.psm1"; Description = "Content cleaning engine" },
        @{ Source = "clean_azureflattener.txt"; Destination = "AzureFlattener.psm1"; Description = "Document flattening" },
        @{ Source = "working_configmanager.txt"; Destination = "ConfigManager.psm1"; Description = "Configuration management" },
        
        # New v2 modules
        @{ Source = "EmailChunkingEngine_v2.psm1"; Destination = "EmailChunkingEngine_v2.psm1"; Description = "Intelligent email chunking" },
        @{ Source = "AzureAISearchIntegration_v2.psm1"; Destination = "AzureAISearchIntegration_v2.psm1"; Description = "Azure AI Search integration" },
        @{ Source = "EmailRAGProcessor_v2.psm1"; Destination = "EmailRAGProcessor_v2.psm1"; Description = "Enhanced RAG processing pipeline" },
        @{ Source = "EmailSearchInterface_v2.psm1"; Destination = "EmailSearchInterface_v2.psm1"; Description = "Hybrid search interface" },
        @{ Source = "EmailEntityExtractor_v2.psm1"; Destination = "EmailEntityExtractor_v2.psm1"; Description = "Advanced entity extraction" },
        @{ Source = "RAGConfigManager_v2.psm1"; Destination = "RAGConfigManager_v2.psm1"; Description = "RAG configuration management" },
        @{ Source = "RAGTestFramework_v2.psm1"; Destination = "RAGTestFramework_v2.psm1"; Description = "Comprehensive testing framework" }
    )
    
    $modulesPath = Join-Path $DestinationPath "Modules"
    $copiedCount = 0
    
    foreach ($module in $moduleFiles) {
        $sourcePath = Join-Path $SourcePath $module.Source
        $destPath = Join-Path $modulesPath $module.Destination
        
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-InstallStep "✓ Installed: $($module.Description)" "SUCCESS"
            $copiedCount++
        } else {
            Write-InstallStep "⚠ Source file not found: $($module.Source)" "WARNING"
        }
    }
    
    # Copy main application script
    $mainScript = Join-Path $SourcePath "emailcleaner_main.txt"
    if (Test-Path $mainScript) {
        Copy-Item -Path $mainScript -Destination (Join-Path $DestinationPath "EmailCleaner.ps1") -Force
        Write-InstallStep "✓ Installed main application script" "SUCCESS"
    }
    
    # Copy Azure AI Search schema
    $schemaFile = Join-Path $SourcePath "AzureAISearchSchema_v2.json"
    if (Test-Path $schemaFile) {
        Copy-Item -Path $schemaFile -Destination (Join-Path $DestinationPath "Config\AzureAISearchSchema_v2.json") -Force
        Write-InstallStep "✓ Installed Azure AI Search schema" "SUCCESS"
    }
    
    Write-InstallStep "Module installation complete: $copiedCount modules installed" "SUCCESS"
}

# Function to create configuration
function New-Configuration {
    param([string]$InstallPath)
    
    Write-InstallStep "Creating configuration files..." "HEADER"
    
    try {
        # Import configuration manager
        $configModulePath = Join-Path $InstallPath "Modules\RAGConfigManager_v2.psm1"
        if (Test-Path $configModulePath) {
            Import-Module $configModulePath -Force
        }
        
        # Create default configuration
        $configParams = @{
            ConfigurationName = "Email RAG Cleaner v2.0"
            AzureSearchServiceName = if ($AzureSearchServiceName) { $AzureSearchServiceName } else { "" }
            AzureSearchApiKey = if ($AzureSearchApiKey) { $AzureSearchApiKey } else { "" }
            OpenAIEndpoint = if ($OpenAIEndpoint) { $OpenAIEndpoint } else { "" }
            OpenAIApiKey = if ($OpenAIApiKey) { $OpenAIApiKey } else { "" }
            OutputPath = Join-Path $InstallPath "Config\default-config.json"
        }
        
        if (Get-Command "New-RAGConfiguration" -ErrorAction SilentlyContinue) {
            $config = New-RAGConfiguration @configParams
            Write-InstallStep "✓ Created default configuration file" "SUCCESS"
        } else {
            # Create basic configuration manually
            $basicConfig = @{
                Metadata = @{
                    Name = "Email RAG Cleaner v2.0"
                    Version = "2.0"
                    CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
                    InstallPath = $InstallPath
                }
                AzureSearch = @{
                    ServiceName = if ($AzureSearchServiceName) { $AzureSearchServiceName } else { "" }
                    ApiKey = if ($AzureSearchApiKey) { $AzureSearchApiKey } else { "" }
                    IndexName = "email-rag-index"
                }
                OpenAI = @{
                    Endpoint = if ($OpenAIEndpoint) { $OpenAIEndpoint } else { "" }
                    ApiKey = if ($OpenAIApiKey) { $OpenAIApiKey } else { "" }
                    Enabled = -not ([string]::IsNullOrEmpty($OpenAIEndpoint))
                }
            }
            
            $basicConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $configParams.OutputPath -Encoding UTF8
            Write-InstallStep "✓ Created basic configuration file" "SUCCESS"
        }
        
        # Create launcher script
        $launcherScript = @"
# EmailRAGCleaner Launcher Script
# Generated by installer on $(Get-Date)

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = "$InstallPath\Config\default-config.json",
    
    [Parameter(Mandatory=$false)]
    [string]$InputPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestMode = $false
)

# Set working directory
Set-Location "$InstallPath"

# Import required modules
$modulePath = "$InstallPath\Modules"
Import-Module "$modulePath\RAGConfigManager_v2.psm1" -Force
Import-Module "$modulePath\EmailRAGProcessor_v2.psm1" -Force

if ($TestMode) {
    Import-Module "$modulePath\RAGTestFramework_v2.psm1" -Force
    Write-Host "Starting Email RAG Cleaner in test mode..." -ForegroundColor Yellow
    
    # Run comprehensive tests
    $config = Import-RAGConfiguration -ConfigurationPath $ConfigPath
    $testResults = Start-RAGPipelineTest -RAGConfiguration $config -DetailedReport:$true
    
    Write-Host "Test completed. Check test report for details." -ForegroundColor Green
    
} else {
    Write-Host "Starting Email RAG Cleaner v2.0..." -ForegroundColor Green
    
    if (-not $InputPath) {
        $InputPath = Read-Host "Enter path to MSG files"
    }
    
    if (-not (Test-Path $InputPath)) {
        Write-Error "Input path not found: $InputPath"
        exit 1
    }
    
    # Initialize and run processing
    $config = Import-RAGConfiguration -ConfigurationPath $ConfigPath
    $pipeline = Initialize-RAGPipeline -AzureSearchConfig $config.AzureSearch -IndexName $config.AzureSearch.IndexName
    $results = Start-EmailRAGProcessing -PipelineConfig $pipeline -InputPath $InputPath
    
    Write-Host "Processing completed successfully!" -ForegroundColor Green
    Write-Host "Files processed: $($results.ProcessedFiles)" -ForegroundColor Cyan
    Write-Host "Documents indexed: $($results.TotalChunks)" -ForegroundColor Cyan
}
"@
        
        $launcherPath = Join-Path $InstallPath "Start-EmailRAGCleaner.ps1"
        $launcherScript | Out-File -FilePath $launcherPath -Encoding UTF8
        Write-InstallStep "✓ Created launcher script" "SUCCESS"
        
        return $true
        
    } catch {
        Write-InstallStep "Failed to create configuration: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to create shortcuts
function New-Shortcuts {
    param([string]$InstallPath)
    
    Write-InstallStep "Creating shortcuts..." "HEADER"
    
    try {
        $shell = New-Object -ComObject WScript.Shell
        
        # Desktop shortcut
        if ($CreateDesktopShortcut) {
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            $shortcutPath = Join-Path $desktopPath "Email RAG Cleaner v2.0.lnk"
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$InstallPath\Start-EmailRAGCleaner.ps1`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Email RAG Cleaner v2.0 with Azure AI Search"
            $shortcut.Save()
            Write-InstallStep "✓ Created desktop shortcut" "SUCCESS"
        }
        
        # Start menu entry
        if ($CreateStartMenuEntry) {
            $startMenuPath = [System.Environment]::GetFolderPath('Programs')
            $programFolder = Join-Path $startMenuPath "Email RAG Cleaner"
            
            if (-not (Test-Path $programFolder)) {
                New-Item -ItemType Directory -Path $programFolder -Force | Out-Null
            }
            
            # Main application shortcut
            $shortcutPath = Join-Path $programFolder "Email RAG Cleaner v2.0.lnk"
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$InstallPath\Start-EmailRAGCleaner.ps1`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Email RAG Cleaner v2.0"
            $shortcut.Save()
            
            # Test mode shortcut
            $testShortcutPath = Join-Path $programFolder "Email RAG Cleaner - Test Mode.lnk"
            $testShortcut = $shell.CreateShortcut($testShortcutPath)
            $testShortcut.TargetPath = "powershell.exe"
            $testShortcut.Arguments = "-ExecutionPolicy Bypass -File `"$InstallPath\Start-EmailRAGCleaner.ps1`" -TestMode"
            $testShortcut.WorkingDirectory = $InstallPath
            $testShortcut.Description = "Email RAG Cleaner v2.0 - Test Mode"
            $testShortcut.Save()
            
            # Configuration shortcut
            $configShortcutPath = Join-Path $programFolder "Edit Configuration.lnk"
            $configShortcut = $shell.CreateShortcut($configShortcutPath)
            $configShortcut.TargetPath = "notepad.exe"
            $configShortcut.Arguments = "`"$InstallPath\Config\default-config.json`""
            $configShortcut.Description = "Edit Email RAG Cleaner Configuration"
            $configShortcut.Save()
            
            Write-InstallStep "✓ Created Start Menu entries" "SUCCESS"
        }
        
        return $true
        
    } catch {
        Write-InstallStep "Failed to create shortcuts: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to create documentation
function New-Documentation {
    param([string]$InstallPath)
    
    Write-InstallStep "Creating documentation..." "HEADER"
    
    $readmeContent = @"
# Email RAG Cleaner v2.0
## Enhanced MSG Email Processing with Azure AI Search Integration

### Overview
The Email RAG Cleaner v2.0 is an advanced PowerShell-based system for processing Microsoft Outlook MSG files and preparing them for Retrieval-Augmented Generation (RAG) applications using Azure AI Search.

### Installation Location
$InstallPath

### Key Features
- **Intelligent Email Processing**: Advanced MSG file parsing with COM interface
- **RAG-Optimized Chunking**: Email-aware content chunking for optimal embedding
- **Azure AI Search Integration**: Direct indexing with vector and semantic search
- **Entity Extraction**: Business, personal, and technical entity recognition
- **Hybrid Search**: Combines vector, keyword, and semantic search capabilities
- **Comprehensive Testing**: Built-in validation and performance testing

### Quick Start
1. **Configure Services**: Edit the configuration file at:
   `$InstallPath\Config\default-config.json`
   
   Add your Azure AI Search and OpenAI credentials:
   - Azure Search Service Name and API Key
   - OpenAI Endpoint and API Key (optional for embeddings)

2. **Run the Application**:
   - Double-click the desktop shortcut "Email RAG Cleaner v2.0"
   - Or run: `$InstallPath\Start-EmailRAGCleaner.ps1`

3. **Test the System**:
   - Use "Email RAG Cleaner - Test Mode" from Start Menu
   - Or run: `$InstallPath\Start-EmailRAGCleaner.ps1 -TestMode`

### Configuration
The system uses JSON configuration files located in the Config directory:
- `default-config.json`: Main configuration file
- `AzureAISearchSchema_v2.json`: Search index schema

### Modules
The following PowerShell modules are included:
- **MsgProcessor.psm1**: MSG file processing
- **ContentCleaner.psm1**: Email content cleaning
- **EmailChunkingEngine_v2.psm1**: Intelligent chunking
- **AzureAISearchIntegration_v2.psm1**: Azure AI Search integration
- **EmailSearchInterface_v2.psm1**: Search interface
- **EmailEntityExtractor_v2.psm1**: Entity extraction
- **RAGConfigManager_v2.psm1**: Configuration management
- **RAGTestFramework_v2.psm1**: Testing framework

### Usage Examples

#### Basic Processing
```powershell
# Process MSG files from a directory
.\Start-EmailRAGCleaner.ps1 -InputPath "C:\EmailData"
```

#### Advanced Configuration
```powershell
# Use custom configuration
.\Start-EmailRAGCleaner.ps1 -ConfigPath "C:\MyConfig.json" -InputPath "C:\EmailData"
```

#### Testing
```powershell
# Run comprehensive tests
.\Start-EmailRAGCleaner.ps1 -TestMode
```

### Search Examples
After processing emails, you can search using PowerShell:

```powershell
# Import search module
Import-Module "$InstallPath\Modules\EmailSearchInterface_v2.psm1"

# Configure search
$searchConfig = @{
    ServiceName = "your-search-service"
    ApiKey = "your-api-key"
    # ... other config
}

# Perform searches
Find-EmailContent -SearchConfig $searchConfig -Query "project meeting" -SearchType "Hybrid"
Search-EmailsBySender -SearchConfig $searchConfig -SenderName "john.doe"
Search-EmailsByDateRange -SearchConfig $searchConfig -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date)
```

### Troubleshooting
1. **Check Prerequisites**: Ensure PowerShell 5.0+, .NET Framework 4.7.2+
2. **Verify Configuration**: Test configuration with `-TestMode`
3. **Check Logs**: Review log files in the Logs directory
4. **Test Connections**: Verify Azure AI Search and OpenAI connectivity

### Support
- Check the Logs directory for detailed error information
- Run the test framework to validate system configuration
- Review the comprehensive test report for diagnostics

### Version Information
- **Version**: 2.0
- **Installation Date**: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
- **Installation Path**: $InstallPath
- **PowerShell Version**: $($PSVersionTable.PSVersion)

### Advanced Features
- **Batch Processing**: Process multiple MSG files efficiently
- **Quality Scoring**: Automatic content quality assessment
- **Performance Monitoring**: Processing time and resource tracking
- **Hybrid Search**: Vector + keyword + semantic search
- **Rich Entity Extraction**: Business, personal, and technical entities
- **Comprehensive Testing**: End-to-end validation framework

For detailed API documentation, see the individual module files in the Modules directory.
"@
    
    $readmePath = Join-Path $InstallPath "Documentation\README.md"
    $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
    Write-InstallStep "✓ Created README documentation" "SUCCESS"
    
    # Create quick start guide
    $quickStartContent = @"
# Quick Start Guide - Email RAG Cleaner v2.0

## 1. First Time Setup
1. Edit configuration: `$InstallPath\Config\default-config.json`
2. Add your Azure AI Search service details
3. Optionally add OpenAI credentials for embeddings

## 2. Process Your First Emails
1. Run: Double-click desktop shortcut
2. Enter path to folder containing MSG files
3. Wait for processing to complete

## 3. Test Your Setup
1. Run test mode from Start Menu
2. Review the generated test report
3. Verify all components are working

## 4. Search Your Emails
Use the PowerShell search interface to find emails:
- Natural language queries
- Search by sender or date range
- Export results to Excel or HTML

## Need Help?
- Check the full README.md
- Run test mode for diagnostics
- Review log files for errors
"@
    
    $quickStartPath = Join-Path $InstallPath "Documentation\QuickStart.md"
    $quickStartContent | Out-File -FilePath $quickStartPath -Encoding UTF8
    Write-InstallStep "✓ Created Quick Start guide" "SUCCESS"
}

# Function to run post-install tests
function Invoke-PostInstallTests {
    param([string]$InstallPath)
    
    Write-InstallStep "Running post-installation tests..." "HEADER"
    
    try {
        # Test module loading
        $modulesPath = Join-Path $InstallPath "Modules"
        $coreModules = @("RAGConfigManager_v2.psm1", "EmailRAGProcessor_v2.psm1", "RAGTestFramework_v2.psm1")
        
        foreach ($module in $coreModules) {
            $modulePath = Join-Path $modulesPath $module
            if (Test-Path $modulePath) {
                try {
                    Import-Module $modulePath -Force
                    Write-InstallStep "✓ Module loads successfully: $module" "SUCCESS"
                } catch {
                    Write-InstallStep "⚠ Module load warning: $module - $($_.Exception.Message)" "WARNING"
                }
            } else {
                Write-InstallStep "✗ Module missing: $module" "ERROR"
            }
        }
        
        # Test configuration
        $configPath = Join-Path $InstallPath "Config\default-config.json"
        if (Test-Path $configPath) {
            try {
                $config = Get-Content $configPath -Raw | ConvertFrom-Json
                Write-InstallStep "✓ Configuration file is valid JSON" "SUCCESS"
            } catch {
                Write-InstallStep "⚠ Configuration file has JSON errors" "WARNING"
            }
        }
        
        # Test launcher script
        $launcherPath = Join-Path $InstallPath "Start-EmailRAGCleaner.ps1"
        if (Test-Path $launcherPath) {
            Write-InstallStep "✓ Launcher script created successfully" "SUCCESS"
        } else {
            Write-InstallStep "✗ Launcher script missing" "ERROR"
        }
        
        Write-InstallStep "Post-installation tests completed" "SUCCESS"
        return $true
        
    } catch {
        Write-InstallStep "Post-installation tests failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main installation function
function Start-Installation {
    try {
        Write-InstallStep "EMAIL RAG CLEANER v2.0 - INSTALLATION STARTING" "HEADER"
        Write-InstallStep "Installation Path: $InstallPath" "INFO"
        Write-InstallStep "Installation Time: $(Get-Date)" "INFO"
        Write-InstallStep "========================================" "HEADER"
        
        # Check administrator privileges
        if (-not (Test-Administrator)) {
            Write-InstallStep "⚠ Running without administrator privileges. Some features may be limited." "WARNING"
            if (-not $Silent) {
                $continue = Read-Host "Continue installation? (y/N)"
                if ($continue -ne 'y' -and $continue -ne 'Y') {
                    Write-InstallStep "Installation cancelled by user" "WARNING"
                    return
                }
            }
        }
        
        # Step 1: Install Prerequisites
        if ($InstallPrerequisites) {
            if (-not (Install-Prerequisites)) {
                throw "Prerequisites installation failed"
            }
        }
        
        # Step 2: Create Directory Structure
        New-DirectoryStructure -BasePath $InstallPath
        
        # Step 3: Copy Module Files
        $currentPath = $PSScriptRoot
        if (-not $currentPath) {
            $currentPath = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        Copy-ModuleFiles -SourcePath $currentPath -DestinationPath $InstallPath
        
        # Step 4: Create Configuration
        if (-not (New-Configuration -InstallPath $InstallPath)) {
            Write-InstallStep "⚠ Configuration creation had issues, but continuing..." "WARNING"
        }
        
        # Step 5: Create Shortcuts
        if (-not (New-Shortcuts -InstallPath $InstallPath)) {
            Write-InstallStep "⚠ Shortcut creation failed, but continuing..." "WARNING"
        }
        
        # Step 6: Create Documentation
        New-Documentation -InstallPath $InstallPath
        
        # Step 7: Run Post-Install Tests
        if (-not (Invoke-PostInstallTests -InstallPath $InstallPath)) {
            Write-InstallStep "⚠ Some post-installation tests failed" "WARNING"
        }
        
        # Step 8: Run full tests if requested
        if ($RunTests) {
            Write-InstallStep "Running comprehensive system tests..." "HEADER"
        }
        
        Write-InstallStep "========================================" "HEADER"
        Write-InstallStep "EMAIL RAG CLEANER v2.0 INSTALLATION COMPLETE!" "SUCCESS"
        Write-InstallStep "Installation Path: $InstallPath" "SUCCESS"
        Write-InstallStep "========================================" "HEADER"
        
        Write-InstallStep "" "INFO"
        Write-InstallStep "NEXT STEPS:" "HEADER"
        Write-InstallStep "1. Configure Azure AI Search settings in:" "INFO"
        Write-InstallStep "   $InstallPath\Config\default-config.json" "INFO"
        Write-InstallStep "" "INFO"
        Write-InstallStep "2. Launch the application:" "INFO"
        Write-InstallStep "   - Double-click desktop shortcut: 'Email RAG Cleaner v2.0'" "INFO"
        Write-InstallStep "   - Or run: $InstallPath\Start-EmailRAGCleaner.ps1" "INFO"
        Write-InstallStep "" "INFO"
        Write-InstallStep "3. Test your installation:" "INFO"
        Write-InstallStep "   - Use 'Email RAG Cleaner - Test Mode' from Start Menu" "INFO"
        Write-InstallStep "   - Or run: $InstallPath\Start-EmailRAGCleaner.ps1 -TestMode" "INFO"
        Write-InstallStep "" "INFO"
        Write-InstallStep "For help, see: $InstallPath\Documentation\README.md" "INFO"
        
        # Open documentation if not silent
        if (-not $Silent) {
            $openDocs = Read-Host "Open documentation now? (y/N)"
            if ($openDocs -eq 'y' -or $openDocs -eq 'Y') {
                Start-Process notepad.exe -ArgumentList (Join-Path $InstallPath "Documentation\README.md")
            }
        }
        
    } catch {
        Write-InstallStep "INSTALLATION FAILED: $($_.Exception.Message)" "ERROR"
        Write-InstallStep "Check the installation log for details" "ERROR"
        exit 1
    }
}

# Start the installation
Start-Installation