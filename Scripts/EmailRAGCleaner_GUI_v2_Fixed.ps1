# Email RAG Cleaner v2.0 - Modern WPF GUI Interface (Fixed)
# Professional Azure AI Search RAG processing with beautiful interface

# Set STA mode for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Ensure we're in STA mode
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Host "ERROR: PowerShell must be running in STA mode for WPF. Use -sta flag when launching PowerShell." -ForegroundColor Red
    Write-Host "Attempting to restart in STA mode..." -ForegroundColor Yellow
    
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -ArgumentList "-sta", "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"" -NoNewWindow
    exit
}

# Error handling wrapper
try {
    # Import required modules
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $installPath = "C:\EmailRAGCleaner"
    $modulesPath = Join-Path $installPath "Modules"

    # Test if installation exists
    if (-not (Test-Path $installPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Email RAG Cleaner installation not found at: $installPath`n`nPlease run the installer first.",
            "Installation Not Found",
            "OK",
            "Error"
        )
        exit 1
    }

    Write-Host "Loading Email RAG Cleaner v2.0 GUI..." -ForegroundColor Green
    Write-Host "Installation path: $installPath" -ForegroundColor Gray

    # Try to load modules with error handling
    $modulesToLoad = @(
        "RAGConfigManager_v2.psm1",
        "EmailRAGProcessor_v2.psm1", 
        "EmailSearchInterface_v2.psm1",
        "RAGTestFramework_v2.psm1"
    )

    $moduleLoadErrors = @()
    foreach ($module in $modulesToLoad) {
        $modulePath = Join-Path $modulesPath $module
        if (Test-Path $modulePath) {
            try {
                Import-Module $modulePath -Force -ErrorAction Stop
                Write-Host "✅ Loaded module: $module" -ForegroundColor Green
            } catch {
                $moduleLoadErrors += "Failed to load $module`: $_"
                Write-Host "⚠️ Module load warning: $module - $_" -ForegroundColor Yellow
            }
        } else {
            $moduleLoadErrors += "Module not found: $module"
            Write-Host "⚠️ Module not found: $module" -ForegroundColor Yellow
        }
    }

    # Global variables
    $global:CurrentConfig = $null
    $global:ProcessingStats = @{
        TotalFiles = 0
        ProcessedFiles = 0
        SuccessfulFiles = 0
        FailedFiles = 0
        TotalChunks = 0
        StartTime = $null
        EndTime = $null
        IsProcessing = $false
    }

    # Create simple test window first
    Write-Host "Creating GUI window..." -ForegroundColor Green

    # XAML for modern WPF interface (simplified for testing)
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email RAG Cleaner v2.0 - Azure AI Search Integration" 
        Height="600" Width="900" 
        MinHeight="400" MinWidth="600"
        WindowStartupLocation="CenterScreen">
    
    <Grid Background="#FF2D2D30">
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#FF1E1E1E" BorderBrush="#FF3F3F46" BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0">
                <TextBlock Text="📧 " FontSize="24" Margin="0,0,10,0" Foreground="White"/>
                <StackPanel>
                    <TextBlock Text="Email RAG Cleaner v2.0" FontSize="18" FontWeight="Bold" Foreground="#FF007ACC"/>
                    <TextBlock Text="Azure AI Search Integration" FontSize="12" Foreground="#FFCCCCCC"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- Main Content -->
        <TabControl Grid.Row="1" Background="#FF2D2D30" BorderThickness="0" Margin="10">
            <!-- Processing Tab -->
            <TabItem Header="📁 Processing">
                <ScrollViewer>
                    <StackPanel Margin="20">
                        <!-- File Selection -->
                        <GroupBox Header="📂 File Selection" Margin="0,0,0,20">
                            <Grid Margin="10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Label Grid.Column="0" Content="MSG Files Path:" Foreground="White"/>
                                <TextBox Name="FilePathTextBox" Grid.Column="1" Margin="5,0" Background="#FF3F3F46" Foreground="White" Padding="5"/>
                                <Button Name="BrowseButton" Grid.Column="2" Content="📁 Browse" Padding="10,5" Background="#FF007ACC" Foreground="White"/>
                            </Grid>
                            <Label Name="FileCountLabel" Content="No files selected" Foreground="#FFCCCCCC" Margin="10,5,10,0"/>
                        </GroupBox>

                        <!-- Processing Options -->
                        <GroupBox Header="⚙️ Processing Options" Margin="0,0,0,20">
                            <WrapPanel Margin="10">
                                <CheckBox Name="CleanContentCheck" Content="🧹 Clean Content" Foreground="White" IsChecked="True" Margin="5"/>
                                <CheckBox Name="ExtractEntitiesCheck" Content="🏷️ Extract Entities" Foreground="White" IsChecked="True" Margin="5"/>
                                <CheckBox Name="CreateRAGCheck" Content="🤖 Create RAG Chunks" Foreground="White" IsChecked="True" Margin="5"/>
                                <CheckBox Name="UploadAzureCheck" Content="☁️ Upload to Azure" Foreground="White" IsChecked="True" Margin="5"/>
                            </WrapPanel>
                        </GroupBox>

                        <!-- Processing Control -->
                        <GroupBox Header="🚀 Processing Control" Margin="0,0,0,20">
                            <StackPanel Margin="10">
                                <ProgressBar Name="ProcessingProgressBar" Height="25" Margin="0,0,0,10"/>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Button Name="StartProcessingButton" Content="▶️ Start Processing" Padding="15,8" Margin="5" Background="#FF28A745" Foreground="White"/>
                                    <Button Name="StopProcessingButton" Content="⏹️ Stop" Padding="15,8" Margin="5" Background="#FFDC3545" Foreground="White" IsEnabled="False"/>
                                </StackPanel>
                                <Label Name="ProcessingStatusLabel" Content="Ready to process emails" Foreground="White" HorizontalAlignment="Center" Margin="0,10,0,0"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Configuration Tab -->
            <TabItem Header="⚙️ Configuration">
                <ScrollViewer>
                    <StackPanel Margin="20">
                        <GroupBox Header="☁️ Azure AI Search Configuration" Margin="0,0,0,20">
                            <Grid Margin="10">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>

                                <Label Grid.Row="0" Grid.Column="0" Content="Service Name:" Foreground="White"/>
                                <TextBox Name="AzureServiceNameTextBox" Grid.Row="0" Grid.Column="1" Background="#FF3F3F46" Foreground="White" Padding="5"/>

                                <Label Grid.Row="1" Grid.Column="0" Content="API Key:" Foreground="White"/>
                                <PasswordBox Name="AzureApiKeyBox" Grid.Row="1" Grid.Column="1" Background="#FF3F3F46" Foreground="White" Padding="5"/>

                                <Button Name="TestConnectionButton" Grid.Row="2" Grid.Column="1" Content="🧪 Test Connection" 
                                        HorizontalAlignment="Left" Padding="10,5" Margin="0,10,0,0" 
                                        Background="#FF007ACC" Foreground="White"/>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Logs Tab -->
            <TabItem Header="📋 Logs">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Button Name="ClearLogsButton" Grid.Row="0" Content="🗑️ Clear Logs" 
                            HorizontalAlignment="Left" Padding="10,5" Margin="0,0,0,10"
                            Background="#FF007ACC" Foreground="White"/>
                    <TextBox Name="LogsTextBox" Grid.Row="1" 
                             Background="#FF1E1E1E" Foreground="#FFCCCCCC" 
                             FontFamily="Consolas" IsReadOnly="True" 
                             TextWrapping="Wrap" AcceptsReturn="True" 
                             VerticalScrollBarVisibility="Auto"/>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Status Bar -->
        <Border Grid.Row="2" Background="#FF1E1E1E" BorderBrush="#FF3F3F46" BorderThickness="0,1,0,0">
            <Label Name="StatusBarLabel" Content="Ready - Email RAG Cleaner v2.0" 
                   Foreground="#FFCCCCCC" VerticalAlignment="Center" Margin="10,0"/>
        </Border>
    </Grid>
</Window>
"@

    # Create WPF window
    Write-Host "Parsing XAML..." -ForegroundColor Gray
    
    [xml]$xamlXml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xamlXml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    Write-Host "✅ Window created successfully" -ForegroundColor Green

    # Get control references
    $controls = @{}
    
    # Find all named elements
    $xamlXml.SelectNodes("//*[@Name]") | ForEach-Object {
        $name = $_.Name
        $control = $window.FindName($name)
        if ($control) {
            $controls[$name] = $control
            Write-Host "  Found control: $name" -ForegroundColor Gray
        }
    }

    # Logging function
    function Write-GUILog {
        param(
            [string]$Message,
            [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
            [string]$Level = "INFO"
        )
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        
        # Update GUI log if available
        if ($controls["LogsTextBox"]) {
            $controls["LogsTextBox"].Dispatcher.Invoke([action]{
                $controls["LogsTextBox"].Text += "$logMessage`r`n"
                $controls["LogsTextBox"].ScrollToEnd()
            })
        }
        
        # Console output
        $color = switch ($Level) {
            "INFO" { "White" }
            "WARN" { "Yellow" }
            "ERROR" { "Red" }
            "SUCCESS" { "Green" }
        }
        Write-Host $logMessage -ForegroundColor $color
    }

    # Event Handlers
    Write-Host "Setting up event handlers..." -ForegroundColor Gray

    if ($controls["BrowseButton"]) {
        $controls["BrowseButton"].Add_Click({
            Write-GUILog "Opening folder browser..." "INFO"
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select folder containing MSG files"
            
            if ($folderDialog.ShowDialog() -eq "OK") {
                $controls["FilePathTextBox"].Text = $folderDialog.SelectedPath
                Write-GUILog "Selected folder: $($folderDialog.SelectedPath)" "INFO"
                
                # Update file count
                try {
                    $msgFiles = Get-ChildItem -Path $folderDialog.SelectedPath -Filter "*.msg" -Recurse -ErrorAction Stop
                    $count = $msgFiles.Count
                    $controls["FileCountLabel"].Content = "📄 $count MSG files found"
                    $controls["FileCountLabel"].Foreground = if ($count -gt 0) { "#FF28A745" } else { "#FFDC3545" }
                    Write-GUILog "Found $count MSG files" "INFO"
                } catch {
                    Write-GUILog "Error counting files: $_" "ERROR"
                    $controls["FileCountLabel"].Content = "❌ Error accessing folder"
                    $controls["FileCountLabel"].Foreground = "#FFDC3545"
                }
            }
        })
    }

    if ($controls["StartProcessingButton"]) {
        $controls["StartProcessingButton"].Add_Click({
            Write-GUILog "Starting processing..." "INFO"
            [System.Windows.MessageBox]::Show(
                "Processing functionality will be implemented here.`n`nThis is a test of the GUI interface.",
                "Processing",
                "OK",
                "Information"
            )
        })
    }

    if ($controls["TestConnectionButton"]) {
        $controls["TestConnectionButton"].Add_Click({
            Write-GUILog "Testing Azure connection..." "INFO"
            [System.Windows.MessageBox]::Show(
                "Azure connection test will be implemented here.",
                "Connection Test",
                "OK",
                "Information"
            )
        })
    }

    if ($controls["ClearLogsButton"]) {
        $controls["ClearLogsButton"].Add_Click({
            $controls["LogsTextBox"].Clear()
            Write-GUILog "Logs cleared" "INFO"
        })
    }

    # Initialize
    Write-GUILog "Email RAG Cleaner v2.0 GUI started successfully" "SUCCESS"
    Write-GUILog "Installation path: $installPath" "INFO"
    
    if ($moduleLoadErrors.Count -gt 0) {
        Write-GUILog "Some modules failed to load. GUI running in limited mode." "WARN"
        foreach ($error in $moduleLoadErrors) {
            Write-GUILog "  - $error" "WARN"
        }
    }

    # Handle window closing
    $window.Add_Closing({
        Write-Host "GUI window closing..." -ForegroundColor Yellow
    })

    # Show the window
    Write-Host "Showing window..." -ForegroundColor Green
    
    # Use ShowDialog to keep the window open
    $null = $window.Dispatcher.InvokeAsync({
        $window.ShowDialog()
    }).Wait()

} catch {
    Write-Host "FATAL ERROR: $_" -ForegroundColor Red
    Write-Host "Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to start Email RAG Cleaner GUI:`n`n$_`n`nCheck the console for details.",
        "Startup Error",
        "OK",
        "Error"
    )
    
    # Keep console open to see error
    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}