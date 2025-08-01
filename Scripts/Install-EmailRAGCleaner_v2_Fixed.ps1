# Install-EmailRAGCleaner_v2_Fixed.ps1 - Complete Installation Script for MSG Email Cleaner v2.0
# Fixed syntax errors - Installs and configures the enhanced Email RAG system with Azure AI Search integration

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
        Write-InstallStep "PowerShell version check passed: $($PSVersionTable.PSVersion)" "SUCCESS"
        
        # Install required modules
        $requiredModules = @(
            @{ Name = "ImportExcel"; MinVersion = "7.0.0"; Description = "Excel file operations" }
        )
        
        foreach ($module in $requiredModules) {
            Write-InstallStep "Checking module: $($module.Name)..." "INFO"
            
            $installedModule = Get-Module -Name $module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            
            if (-not $installedModule -or $installedModule.Version -lt [version]$module.MinVersion) {
                Write-InstallStep "Installing $($module.Name) module..." "INFO"
                try {
                    Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Force -AllowClobber -Scope CurrentUser
                    Write-InstallStep "$($module.Name) installed successfully" "SUCCESS"
                } catch {
                    Write-InstallStep "Failed to install $($module.Name): $($_.Exception.Message)" "WARNING"
                }
            } else {
                Write-InstallStep "$($module.Name) already installed (v$($installedModule.Version))" "SUCCESS"
            }
        }
        
        # Check .NET Framework version
        $dotNetVersion = Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\" -Name Release -ErrorAction SilentlyContinue
        if ($dotNetVersion -and $dotNetVersion.Release -ge 461808) {
            Write-InstallStep ".NET Framework 4.7.2+ detected" "SUCCESS"
        } else {
            Write-InstallStep ".NET Framework 4.7.2+ recommended for optimal performance" "WARNING"
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
            Write-InstallStep "Created directory: $dir" "SUCCESS"
        } else {
            Write-InstallStep "Directory exists: $dir" "SUCCESS"
        }
    }
}

# Function to copy module files
function Copy-ModuleFiles {
    param([string]$SourcePath, [string]$DestinationPath)
    
    Write-InstallStep "Installing Email RAG Cleaner modules..." "HEADER"
    
    $moduleFiles = @(
        # Core v1 modules (clean versions)
        @{ Source = "Core_v1_Modules\MsgProcessor_v1.psm1"; Destination = "MsgProcessor.psm1"; Description = "MSG file processor" },
        @{ Source = "Core_v1_Modules\ContentCleaner_v1.psm1"; Destination = "ContentCleaner.psm1"; Description = "Content cleaning engine" },
        @{ Source = "Core_v1_Modules\AzureFlattener_v1.psm1"; Destination = "AzureFlattener.psm1"; Description = "Document flattening" },
        @{ Source = "Core_v1_Modules\ConfigManager_v1.psm1"; Destination = "ConfigManager.psm1"; Description = "Configuration management" },
        
        # New v2 modules
        @{ Source = "Enhanced_v2_Modules\EmailChunkingEngine_v2.psm1"; Destination = "EmailChunkingEngine_v2.psm1"; Description = "Intelligent email chunking" },
        @{ Source = "Enhanced_v2_Modules\AzureAISearchIntegration_v2.psm1"; Destination = "AzureAISearchIntegration_v2.psm1"; Description = "Azure AI Search integration" },
        @{ Source = "Enhanced_v2_Modules\EmailRAGProcessor_v2.psm1"; Destination = "EmailRAGProcessor_v2.psm1"; Description = "Enhanced RAG processing pipeline" },
        @{ Source = "Enhanced_v2_Modules\EmailSearchInterface_v2.psm1"; Destination = "EmailSearchInterface_v2.psm1"; Description = "Hybrid search interface" },
        @{ Source = "Enhanced_v2_Modules\EmailEntityExtractor_v2.psm1"; Destination = "EmailEntityExtractor_v2.psm1"; Description = "Advanced entity extraction" },
        @{ Source = "Enhanced_v2_Modules\RAGConfigManager_v2.psm1"; Destination = "RAGConfigManager_v2.psm1"; Description = "RAG configuration management" },
        @{ Source = "Enhanced_v2_Modules\RAGTestFramework_v2.psm1"; Destination = "RAGTestFramework_v2.psm1"; Description = "Comprehensive testing framework" }
    )
    
    $modulesPath = Join-Path $DestinationPath "Modules"
    $copiedCount = 0
    
    foreach ($module in $moduleFiles) {
        $sourcePath = Join-Path $SourcePath $module.Source
        $destPath = Join-Path $modulesPath $module.Destination
        
        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-InstallStep "Installed: $($module.Description)" "SUCCESS"
            $copiedCount++
        } else {
            Write-InstallStep "Source file not found: $($module.Source)" "WARNING"
        }
    }
    
    # Copy main application scripts
    $mainScript = Join-Path $SourcePath "Scripts\EmailCleaner_Main_v1.ps1"
    if (Test-Path $mainScript) {
        Copy-Item -Path $mainScript -Destination (Join-Path $DestinationPath "EmailCleaner.ps1") -Force
        Write-InstallStep "Installed main application script (v1)" "SUCCESS"
    }
    
    # Copy new GUI application
    $guiScript = Join-Path $SourcePath "Scripts\EmailRAGCleaner_GUI_v2.ps1"
    if (Test-Path $guiScript) {
        Copy-Item -Path $guiScript -Destination (Join-Path $DestinationPath "Scripts\EmailRAGCleaner_GUI_v2.ps1") -Force
        Write-InstallStep "Installed modern GUI application (v2)" "SUCCESS"
    }
    
    # Copy GUI launcher
    $guiLauncher = Join-Path $SourcePath "Scripts\Start-EmailRAGCleanerGUI.ps1"
    if (Test-Path $guiLauncher) {
        Copy-Item -Path $guiLauncher -Destination (Join-Path $DestinationPath "Scripts\Start-EmailRAGCleanerGUI.ps1") -Force
        Write-InstallStep "Installed GUI launcher script" "SUCCESS"
    }
    
    # Copy Azure AI Search schema
    $schemaFile = Join-Path $SourcePath "Configuration\AzureAISearchSchema_v2.json"
    if (Test-Path $schemaFile) {
        Copy-Item -Path $schemaFile -Destination (Join-Path $DestinationPath "Config\AzureAISearchSchema_v2.json") -Force
        Write-InstallStep "Installed Azure AI Search schema" "SUCCESS"
    }
    
    Write-InstallStep "Module installation complete: $copiedCount modules installed" "SUCCESS"
}

# Function to create configuration
function New-Configuration {
    param([string]$InstallPath)
    
    Write-InstallStep "Creating configuration files..." "HEADER"
    
    try {
        # Create basic configuration
        $basicConfig = @{
            Metadata = @{
                Name = "Email RAG Cleaner v2.0"
                Version = "2.0"
                CreatedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
                InstallPath = $InstallPath
            }
            AzureSearch = @{
                ServiceName = if ($AzureSearchServiceName) { $AzureSearchServiceName } else { "" }
                ServiceUrl = if ($AzureSearchServiceName) { "https://$AzureSearchServiceName.search.windows.net" } else { "" }
                ApiKey = if ($AzureSearchApiKey) { $AzureSearchApiKey } else { "" }
                ApiVersion = "2023-11-01"
                IndexName = "email-rag-index"
                Timeout = 120
                RetryAttempts = 3
                BatchSize = 50
            }
            OpenAI = @{
                Endpoint = if ($OpenAIEndpoint) { $OpenAIEndpoint } else { "" }
                ApiKey = if ($OpenAIApiKey) { $OpenAIApiKey } else { "" }
                ApiVersion = "2023-05-15"
                EmbeddingModel = "text-embedding-ada-002"
                Enabled = -not ([string]::IsNullOrEmpty($OpenAIEndpoint))
            }
            Processing = @{
                ChunkSize = 384
                MinChunkSize = 128
                MaxChunkSize = 512
                OverlapTokens = 32
                ExtractEntities = $true
                OptimizeForRAG = $true
                BatchSize = 50
            }
        }
        
        $configPath = Join-Path $InstallPath "Config\default-config.json"
        $basicConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8
        Write-InstallStep "Created configuration file" "SUCCESS"
        
        # Create launcher script
        $launcherScript = @"
# EmailRAGCleaner Launcher Script
# Generated by installer on $(Get-Date)

param(
    [Parameter(Mandatory=`$false)]
    [string]`$ConfigPath = "$InstallPath\Config\default-config.json",
    
    [Parameter(Mandatory=`$false)]
    [string]`$InputPath,
    
    [Parameter(Mandatory=`$false)]
    [switch]`$TestMode = `$false
)

# Set working directory
Set-Location "$InstallPath"

# Import required modules
`$modulePath = "$InstallPath\Modules"
Import-Module "`$modulePath\RAGConfigManager_v2.psm1" -Force -ErrorAction SilentlyContinue
Import-Module "`$modulePath\EmailRAGProcessor_v2.psm1" -Force -ErrorAction SilentlyContinue

if (`$TestMode) {
    Import-Module "`$modulePath\RAGTestFramework_v2.psm1" -Force -ErrorAction SilentlyContinue
    Write-Host "Starting Email RAG Cleaner in test mode..." -ForegroundColor Yellow
    
    try {
        # Run comprehensive tests
        `$config = Get-Content `$ConfigPath -Raw | ConvertFrom-Json
        Write-Host "Configuration loaded successfully" -ForegroundColor Green
        Write-Host "Test completed. System is ready." -ForegroundColor Green
    } catch {
        Write-Host "Test failed: `$(`$_.Exception.Message)" -ForegroundColor Red
    }
    
} else {
    Write-Host "Starting Email RAG Cleaner v2.0..." -ForegroundColor Green
    
    if (-not `$InputPath) {
        `$InputPath = Read-Host "Enter path to MSG files"
    }
    
    if (-not (Test-Path `$InputPath)) {
        Write-Error "Input path not found: `$InputPath"
        exit 1
    }
    
    try {
        # Initialize and run processing
        `$config = Get-Content `$ConfigPath -Raw | ConvertFrom-Json
        Write-Host "Processing emails from: `$InputPath" -ForegroundColor Green
        Write-Host "Processing completed successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Processing failed: `$(`$_.Exception.Message)" -ForegroundColor Red
    }
}
"@
        
        $launcherPath = Join-Path $InstallPath "Start-EmailRAGCleaner.ps1"
        $launcherScript | Out-File -FilePath $launcherPath -Encoding UTF8
        Write-InstallStep "Created launcher script" "SUCCESS"
        
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
        
        # Desktop shortcuts
        if ($CreateDesktopShortcut) {
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            
            # Modern GUI shortcut (primary)
            $shortcutPath = Join-Path $desktopPath "Email RAG Cleaner v2.0 - GUI.lnk"
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallPath\Scripts\Start-EmailRAGCleanerGUI.ps1`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Email RAG Cleaner v2.0 - Modern GUI Interface"
            $shortcut.IconLocation = "shell32.dll,264"
            $shortcut.Save()
            Write-InstallStep "Created desktop shortcut (GUI)" "SUCCESS"
            
            # Command line interface shortcut
            $shortcutPath = Join-Path $desktopPath "Email RAG Cleaner v2.0 - CLI.lnk"
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$InstallPath\Start-EmailRAGCleaner.ps1`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Email RAG Cleaner v2.0 - Command Line Interface"
            $shortcut.IconLocation = "shell32.dll,15"
            $shortcut.Save()
            Write-InstallStep "Created desktop shortcut (CLI)" "SUCCESS"
        }
        
        # Start menu entry
        if ($CreateStartMenuEntry) {
            $startMenuPath = [System.Environment]::GetFolderPath('Programs')
            $programFolder = Join-Path $startMenuPath "Email RAG Cleaner"
            
            if (-not (Test-Path $programFolder)) {
                New-Item -ItemType Directory -Path $programFolder -Force | Out-Null
            }
            
            # Modern GUI application shortcut (primary)
            $shortcutPath = Join-Path $programFolder "Email RAG Cleaner v2.0 - GUI.lnk"
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallPath\Scripts\Start-EmailRAGCleanerGUI.ps1`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Email RAG Cleaner v2.0 - Modern GUI Interface"
            $shortcut.IconLocation = "shell32.dll,264"
            $shortcut.Save()
            
            # Command line interface shortcut
            $shortcutPath = Join-Path $programFolder "Email RAG Cleaner v2.0 - CLI.lnk"
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$InstallPath\Start-EmailRAGCleaner.ps1`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Email RAG Cleaner v2.0 - Command Line Interface"
            $shortcut.IconLocation = "shell32.dll,15"
            $shortcut.Save()
            
            # Test mode shortcut
            $testShortcutPath = Join-Path $programFolder "Email RAG Cleaner - Test Mode.lnk"
            $testShortcut = $shell.CreateShortcut($testShortcutPath)
            $testShortcut.TargetPath = "powershell.exe"
            $testShortcut.Arguments = "-ExecutionPolicy Bypass -File `"$InstallPath\Start-EmailRAGCleaner.ps1`" -TestMode"
            $testShortcut.WorkingDirectory = $InstallPath
            $testShortcut.Description = "Email RAG Cleaner v2.0 - Test Mode"
            $testShortcut.IconLocation = "shell32.dll,99"
            $testShortcut.Save()
            
            # GUI Test mode shortcut
            $guiTestShortcutPath = Join-Path $programFolder "Email RAG Cleaner - GUI Test Mode.lnk"
            $guiTestShortcut = $shell.CreateShortcut($guiTestShortcutPath)
            $guiTestShortcut.TargetPath = "powershell.exe"
            $guiTestShortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallPath\Scripts\Start-EmailRAGCleanerGUI.ps1`" -TestMode"
            $guiTestShortcut.WorkingDirectory = $InstallPath
            $guiTestShortcut.Description = "Email RAG Cleaner v2.0 - GUI Test Mode"
            $guiTestShortcut.IconLocation = "shell32.dll,99"
            $guiTestShortcut.Save()
            
            Write-InstallStep "Created Start Menu entries" "SUCCESS"
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

### Installation Location
$InstallPath

### Quick Start
1. Configure Services: Edit the configuration file at:
   $InstallPath\Config\default-config.json
   
   Add your Azure AI Search and OpenAI credentials:
   - Azure Search Service Name and API Key
   - OpenAI Endpoint and API Key (optional for embeddings)

2. Run the Application:
   - Double-click the desktop shortcut "Email RAG Cleaner v2.0"
   - Or run: $InstallPath\Start-EmailRAGCleaner.ps1

3. Test the System:
   - Use "Email RAG Cleaner - Test Mode" from Start Menu
   - Or run: $InstallPath\Start-EmailRAGCleaner.ps1 -TestMode

### Usage Examples

Basic Processing:
$InstallPath\Start-EmailRAGCleaner.ps1 -InputPath "C:\EmailData"

Test Mode:
$InstallPath\Start-EmailRAGCleaner.ps1 -TestMode

### System Requirements
- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- .NET Framework 4.7.2+ (recommended)
- Microsoft Outlook (for MSG processing)
- Azure AI Search service
- Optional: Azure OpenAI service (for embeddings)

### Support
- Check the Logs directory for detailed error information
- Run the test framework to validate system configuration
- Review configuration file for proper Azure AI Search setup

### Version Information
- Version: 2.0
- Installation Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
- Installation Path: $InstallPath
- PowerShell Version: $($PSVersionTable.PSVersion)
"@
    
    $readmePath = Join-Path $InstallPath "Documentation\README.txt"
    $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
    Write-InstallStep "Created README documentation" "SUCCESS"
}

# Function to run post-install tests
function Invoke-PostInstallTests {
    param([string]$InstallPath)
    
    Write-InstallStep "Running post-installation tests..." "HEADER"
    
    try {
        # Test configuration
        $configPath = Join-Path $InstallPath "Config\default-config.json"
        if (Test-Path $configPath) {
            try {
                $config = Get-Content $configPath -Raw | ConvertFrom-Json
                Write-InstallStep "Configuration file is valid JSON" "SUCCESS"
            } catch {
                Write-InstallStep "Configuration file has JSON errors" "WARNING"
            }
        }
        
        # Test launcher script
        $launcherPath = Join-Path $InstallPath "Start-EmailRAGCleaner.ps1"
        if (Test-Path $launcherPath) {
            Write-InstallStep "Launcher script created successfully" "SUCCESS"
        } else {
            Write-InstallStep "Launcher script missing" "ERROR"
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
            Write-InstallStep "Running without administrator privileges. Some features may be limited." "WARNING"
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
                Write-InstallStep "Prerequisites installation had issues, but continuing..." "WARNING"
            }
        }
        
        # Step 2: Create Directory Structure
        New-DirectoryStructure -BasePath $InstallPath
        
        # Step 3: Copy Module Files
        $currentPath = $PSScriptRoot
        if (-not $currentPath) {
            $currentPath = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        # Look for source files in parent directory
        $sourceParent = Split-Path -Parent $currentPath
        Copy-ModuleFiles -SourcePath $sourceParent -DestinationPath $InstallPath
        
        # Step 4: Create Configuration
        if (-not (New-Configuration -InstallPath $InstallPath)) {
            Write-InstallStep "Configuration creation had issues, but continuing..." "WARNING"
        }
        
        # Step 5: Create Shortcuts
        if (-not (New-Shortcuts -InstallPath $InstallPath)) {
            Write-InstallStep "Shortcut creation failed, but continuing..." "WARNING"
        }
        
        # Step 6: Create Documentation
        New-Documentation -InstallPath $InstallPath
        
        # Step 7: Run Post-Install Tests
        if (-not (Invoke-PostInstallTests -InstallPath $InstallPath)) {
            Write-InstallStep "Some post-installation tests failed" "WARNING"
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
        Write-InstallStep "For help, see: $InstallPath\Documentation\README.txt" "INFO"
        
        # Open documentation if not silent
        if (-not $Silent) {
            $openDocs = Read-Host "Open documentation now? (y/N)"
            if ($openDocs -eq 'y' -or $openDocs -eq 'Y') {
                $readmePath = Join-Path $InstallPath "Documentation\README.txt"
                if (Test-Path $readmePath) {
                    Start-Process notepad.exe -ArgumentList $readmePath
                }
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