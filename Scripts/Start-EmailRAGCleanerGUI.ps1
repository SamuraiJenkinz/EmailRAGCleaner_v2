# Email RAG Cleaner v2.0 - GUI Launcher Script
# Launches the modern WPF interface with proper error handling

param(
    [Parameter(Mandatory=$false)]
    [switch]$TestMode = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Debug = $false
)

# Set error handling
$ErrorActionPreference = "Stop"

# Define paths
$installPath = "C:\EmailRAGCleaner"
$scriptsPath = Join-Path $installPath "Scripts"
$guiScript = Join-Path $scriptsPath "EmailRAGCleaner_GUI_v2.ps1"
$logPath = Join-Path $installPath "Logs"

# Ensure log directory exists
if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath -Force | Out-Null
}

# Function to write startup log
function Write-StartupLog {
    param([string]$Message, [string]$Level = "INFO")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] GUI Launcher: $Message"
    
    $logFile = Join-Path $logPath "GUI_Launcher_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logMessage
    
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

try {
    Write-StartupLog "Starting Email RAG Cleaner v2.0 GUI..." "INFO"
    
    # Check if installation exists
    if (-not (Test-Path $installPath)) {
        throw "Email RAG Cleaner v2.0 installation not found at: $installPath"
    }
    
    # Check if GUI script exists
    if (-not (Test-Path $guiScript)) {
        throw "GUI script not found at: $guiScript"
    }
    
    # Verify PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5.0 or higher is required. Current version: $($PSVersionTable.PSVersion)"
    }
    
    Write-StartupLog "System checks passed" "SUCCESS"
    
    # Set working directory
    Set-Location $installPath
    Write-StartupLog "Working directory set to: $installPath" "INFO"
    
    # Check if we're in STA mode
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Write-StartupLog "PowerShell not in STA mode, restarting..." "WARN"
        
        $args = @("-sta", "-ExecutionPolicy", "Bypass", "-File", $MyInvocation.MyCommand.Path)
        if ($TestMode) { $args += "-TestMode" }
        if ($Debug) { $args += "-Debug" }
        
        Start-Process powershell.exe -ArgumentList $args -NoNewWindow
        exit
    }
    
    Write-StartupLog "Running in STA mode" "INFO"
    
    if ($TestMode) {
        Write-StartupLog "Starting in test mode..." "INFO"
        # Use the fixed GUI script
        & "$scriptsPath\EmailRAGCleaner_GUI_v2_Fixed.ps1" -TestMode
    } elseif ($Debug) {
        Write-StartupLog "Starting in debug mode..." "INFO"
        # Use the fixed GUI script
        & "$scriptsPath\EmailRAGCleaner_GUI_v2_Fixed.ps1" -Debug
    } else {
        Write-StartupLog "Starting normal GUI mode..." "INFO"
        # Use the fixed GUI script
        & "$scriptsPath\EmailRAGCleaner_GUI_v2_Fixed.ps1"
    }
    
    Write-StartupLog "GUI session completed" "SUCCESS"
    
} catch {
    Write-StartupLog "Failed to start GUI: $($_.Exception.Message)" "ERROR"
    
    # Show error dialog
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to start Email RAG Cleaner v2.0 GUI:`n`n$($_.Exception.Message)`n`nPlease check the installation and try again.",
        "Startup Error",
        "OK",
        "Error"
    )
    
    exit 1
}