# Email RAG Cleaner v2.0 - Fully Functional WPF GUI Interface
# Now with REAL email processing functionality powered by PowerShell 7

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
    # Import required modules (use the FIXED versions!)
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $installPath = "C:\EmailRAGCleaner"
    $modulesPath = Join-Path $installPath "Modules"
    $enhancedModulesPath = "C:\users\taylo\Downloads\EmailRAGCleaner_v2\Enhanced_v2_Modules"

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

    Write-Host "🚀 Loading Email RAG Cleaner v2.0 GUI with REAL functionality..." -ForegroundColor Green
    Write-Host "Installation path: $installPath" -ForegroundColor Gray

    # Load the FIXED modules with real functionality
    $modulesToLoad = @(
        @{ Name = "RAGConfigManager_v2_Fixed.psm1"; Required = $true },
        @{ Name = "EmailRAGProcessor_v2_Fixed.psm1"; Required = $true },
        @{ Name = "RAGTestFramework_v2_Fixed.psm1"; Required = $false },
        @{ Name = "EmailSearchInterface_v2.psm1"; Required = $false }
    )

    $moduleLoadErrors = @()
    $loadedModules = @()
    
    foreach ($moduleInfo in $modulesToLoad) {
        $modulePath = Join-Path $enhancedModulesPath $moduleInfo.Name
        if (Test-Path $modulePath) {
            try {
                Import-Module $modulePath -Force -ErrorAction Stop
                Write-Host "✅ Loaded module: $($moduleInfo.Name)" -ForegroundColor Green
                $loadedModules += $moduleInfo.Name
            } catch {
                $error = "Failed to load $($moduleInfo.Name): $_"
                $moduleLoadErrors += $error
                if ($moduleInfo.Required) {
                    Write-Host "❌ CRITICAL: $error" -ForegroundColor Red
                } else {
                    Write-Host "⚠️ Optional module failed: $($moduleInfo.Name)" -ForegroundColor Yellow
                }
            }
        } else {
            $error = "Module not found: $($moduleInfo.Name)"
            $moduleLoadErrors += $error
            if ($moduleInfo.Required) {
                Write-Host "❌ CRITICAL: $error" -ForegroundColor Red
            }
        }
    }

    # Check if critical modules loaded
    $criticalModulesLoaded = $loadedModules -contains "RAGConfigManager_v2_Fixed.psm1" -and 
                           $loadedModules -contains "EmailRAGProcessor_v2_Fixed.psm1"

    if (-not $criticalModulesLoaded) {
        $errorMsg = "Critical modules failed to load. Cannot start application.`n`nErrors:`n" + ($moduleLoadErrors -join "`n")
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Critical Error", "OK", "Error")
        exit 1
    }

    # Global variables for application state
    $global:CurrentConfig = $null
    $global:ProcessingJob = $null
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

    # Create the main GUI window
    Write-Host "🎨 Creating functional GUI window..." -ForegroundColor Green

    # XAML for modern WPF interface with enhanced functionality
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email RAG Cleaner v2.0 - Azure AI Search Integration (FUNCTIONAL)" 
        Height="700" Width="1000" 
        MinHeight="500" MinWidth="800"
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
                <TextBlock Text="🚀 " FontSize="24" Margin="0,0,10,0" Foreground="White"/>
                <StackPanel>
                    <TextBlock Text="Email RAG Cleaner v2.0 (FUNCTIONAL)" FontSize="18" FontWeight="Bold" Foreground="#FF007ACC"/>
                    <TextBlock Text="Real Azure AI Search Processing • PowerShell 7 Enhanced" FontSize="12" Foreground="#FFCCCCCC"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- Main Content -->
        <TabControl Grid.Row="1" Background="#FF2D2D30" BorderThickness="0" Margin="10">
            <!-- Processing Tab -->
            <TabItem Header="🎯 Smart Processing">
                <ScrollViewer>
                    <StackPanel Margin="20">
                        <!-- File Selection -->
                        <GroupBox Header="📂 File Selection &amp; Analysis" Margin="0,0,0,20">
                            <StackPanel>
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
                                <Label Name="FileAnalysisLabel" Content="" Foreground="#FF90EE90" Margin="10,0,10,0"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- Processing Configuration -->
                        <GroupBox Header="⚙️ Processing Configuration" Margin="0,0,0,20">
                            <StackPanel Margin="10">
                                <WrapPanel Margin="0,0,0,10">
                                    <CheckBox Name="CleanContentCheck" Content="🧹 Clean Content" Foreground="White" IsChecked="True" Margin="5"/>
                                    <CheckBox Name="ExtractEntitiesCheck" Content="🏷️ Extract Entities" Foreground="White" IsChecked="True" Margin="5"/>
                                    <CheckBox Name="CreateRAGCheck" Content="🤖 Create RAG Chunks" Foreground="White" IsChecked="True" Margin="5"/>
                                    <CheckBox Name="UploadAzureCheck" Content="☁️ Upload to Azure" Foreground="White" IsChecked="True" Margin="5"/>
                                    <CheckBox Name="ParallelProcessingCheck" Content="⚡ Parallel Processing" Foreground="White" IsChecked="True" Margin="5"/>
                                </WrapPanel>
                                <StackPanel Orientation="Horizontal" Margin="0,5">
                                    <Label Content="Max Concurrency:" Foreground="White" Margin="0,0,5,0"/>
                                    <Slider Name="ConcurrencySlider" Minimum="1" Maximum="10" Value="5" Width="100" Margin="0,0,10,0"/>
                                    <Label Name="ConcurrencyLabel" Content="5" Foreground="#FF007ACC" Width="20"/>
                                </StackPanel>
                            </StackPanel>
                        </GroupBox>

                        <!-- Processing Control -->
                        <GroupBox Header="🚀 Real-Time Processing Control" Margin="0,0,0,20">
                            <StackPanel Margin="10">
                                <ProgressBar Name="ProcessingProgressBar" Height="25" Margin="0,0,0,10" Background="#FF3F3F46" Foreground="#FF007ACC"/>
                                <Label Name="ProgressLabel" Content="Ready to process emails" Foreground="White" HorizontalAlignment="Center" FontWeight="Bold"/>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
                                    <Button Name="StartProcessingButton" Content="▶️ Start REAL Processing" Padding="15,8" Margin="5" Background="#FF28A745" Foreground="White" FontWeight="Bold"/>
                                    <Button Name="StopProcessingButton" Content="⏹️ Stop" Padding="15,8" Margin="5" Background="#FFDC3545" Foreground="White" IsEnabled="False"/>
                                    <Button Name="TestConfigButton" Content="🧪 Test Config" Padding="15,8" Margin="5" Background="#FF6F42C1" Foreground="White"/>
                                </StackPanel>
                            </StackPanel>
                        </GroupBox>

                        <!-- Real-Time Statistics -->
                        <GroupBox Header="📊 Real-Time Processing Statistics" Margin="0,0,0,20">
                            <Grid Margin="10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" Margin="5">
                                    <Label Name="TotalFilesLabel" Content="Total: 0" Foreground="#FF007ACC" FontWeight="Bold" HorizontalAlignment="Center"/>
                                    <Label Content="Files" Foreground="White" HorizontalAlignment="Center"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="5">
                                    <Label Name="ProcessedFilesLabel" Content="Processed: 0" Foreground="#FF28A745" FontWeight="Bold" HorizontalAlignment="Center"/>
                                    <Label Content="Success" Foreground="White" HorizontalAlignment="Center"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Margin="5">
                                    <Label Name="ChunksLabel" Content="Chunks: 0" Foreground="#FFFFC107" FontWeight="Bold" HorizontalAlignment="Center"/>
                                    <Label Content="Generated" Foreground="White" HorizontalAlignment="Center"/>
                                </StackPanel>
                                <StackPanel Grid.Column="3" Margin="5">
                                    <Label Name="IndexedLabel" Content="Indexed: 0" Foreground="#FF17A2B8" FontWeight="Bold" HorizontalAlignment="Center"/>
                                    <Label Content="Documents" Foreground="White" HorizontalAlignment="Center"/>
                                </StackPanel>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Configuration Tab -->
            <TabItem Header="⚙️ Azure Configuration">
                <ScrollViewer>
                    <StackPanel Margin="20">
                        <GroupBox Header="☁️ Azure AI Search Configuration" Margin="0,0,0,20">
                            <Grid Margin="10">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>

                                <Label Grid.Row="0" Grid.Column="0" Content="Service Name:" Foreground="White"/>
                                <TextBox Name="AzureServiceNameTextBox" Grid.Row="0" Grid.Column="1" Background="#FF3F3F46" Foreground="White" Padding="5" Text="your-search-service"/>

                                <Label Grid.Row="1" Grid.Column="0" Content="API Key:" Foreground="White"/>
                                <PasswordBox Name="AzureApiKeyBox" Grid.Row="1" Grid.Column="1" Background="#FF3F3F46" Foreground="White" Padding="5"/>

                                <Label Grid.Row="2" Grid.Column="0" Content="Index Name:" Foreground="White"/>
                                <TextBox Name="IndexNameTextBox" Grid.Row="2" Grid.Column="1" Background="#FF3F3F46" Foreground="White" Padding="5" Text="email-rag-index"/>

                                <StackPanel Grid.Row="3" Grid.Column="1" Orientation="Horizontal" Margin="0,10,0,0">
                                    <Button Name="TestConnectionButton" Content="🧪 Test Connection" Padding="10,5" Margin="0,0,10,0" Background="#FF007ACC" Foreground="White"/>
                                    <Button Name="SaveConfigButton" Content="💾 Save Config" Padding="10,5" Margin="0,0,10,0" Background="#FF28A745" Foreground="White"/>
                                    <Button Name="LoadConfigButton" Content="📂 Load Config" Padding="10,5" Background="#FF6C757D" Foreground="White"/>
                                </StackPanel>
                            </Grid>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Real-Time Logs Tab -->
            <TabItem Header="📋 Real-Time Logs">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button Name="ClearLogsButton" Content="🗑️ Clear Logs" Padding="10,5" Margin="0,0,10,0" Background="#FF007ACC" Foreground="White"/>
                        <Button Name="ExportLogsButton" Content="📤 Export Logs" Padding="10,5" Margin="0,0,10,0" Background="#FF28A745" Foreground="White"/>
                        <CheckBox Name="AutoScrollCheck" Content="Auto-scroll" Foreground="White" IsChecked="True" VerticalAlignment="Center" Margin="10,0,0,0"/>
                    </StackPanel>
                    <TextBox Name="LogsTextBox" Grid.Row="1" 
                             Background="#FF1E1E1E" Foreground="#FFCCCCCC" 
                             FontFamily="Consolas" IsReadOnly="True" 
                             TextWrapping="Wrap" AcceptsReturn="True" 
                             VerticalScrollBarVisibility="Auto"/>
                </Grid>
            </TabItem>

            <!-- Results & Analytics Tab -->
            <TabItem Header="📈 Results &amp; Analytics">
                <ScrollViewer>
                    <StackPanel Margin="20">
                        <GroupBox Header="📊 Processing Results Summary" Margin="0,0,0,20">
                            <StackPanel Margin="10">
                                <TextBlock Name="ResultsSummaryText" Text="No processing completed yet. Run processing to see results." 
                                          Foreground="White" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                <Button Name="GenerateReportButton" Content="📄 Generate Detailed Report" 
                                       Padding="10,5" Background="#FF6F42C1" Foreground="White" 
                                       HorizontalAlignment="Left" IsEnabled="False"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Enhanced Status Bar -->
        <Border Grid.Row="2" Background="#FF1E1E1E" BorderBrush="#FF3F3F46" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Name="StatusBarLabel" Grid.Column="0" Content="Ready - Email RAG Cleaner v2.0 with REAL processing capability" 
                       Foreground="#FFCCCCCC" VerticalAlignment="Center" Margin="10,0"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="10,0">
                    <Label Name="ModuleStatusLabel" Content="✅ Modules Loaded" Foreground="#FF28A745" VerticalAlignment="Center"/>
                    <Label Name="ProcessingTimeLabel" Content="" Foreground="#FF007ACC" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    # Create WPF window
    Write-Host "🎨 Parsing XAML..." -ForegroundColor Gray
    
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

    # Enhanced logging function with real-time updates
    function Write-GUILog {
        param(
            [string]$Message,
            [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "PROGRESS")]
            [string]$Level = "INFO"
        )
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $emoji = switch ($Level) {
            "INFO" { "ℹ️" }
            "WARN" { "⚠️" }
            "ERROR" { "❌" }
            "SUCCESS" { "✅" }
            "PROGRESS" { "🔄" }
        }
        
        $logMessage = "[$timestamp] $emoji [$Level] $Message"
        
        # Update GUI log if available
        if ($controls["LogsTextBox"]) {
            $controls["LogsTextBox"].Dispatcher.Invoke([action]{
                $controls["LogsTextBox"].Text += "$logMessage`r`n"
                if ($controls["AutoScrollCheck"].IsChecked) {
                    $controls["LogsTextBox"].ScrollToEnd()
                }
            })
        }
        
        # Console output with colors
        $color = switch ($Level) {
            "INFO" { "White" }
            "WARN" { "Yellow" }
            "ERROR" { "Red" }
            "SUCCESS" { "Green" }
            "PROGRESS" { "Cyan" }
        }
        Write-Host $logMessage -ForegroundColor $color
    }

    # Function to update processing statistics in real-time
    function Update-ProcessingStats {
        param(
            [hashtable]$Stats
        )
        
        if ($controls["TotalFilesLabel"]) {
            $controls["TotalFilesLabel"].Dispatcher.Invoke([action]{
                $controls["TotalFilesLabel"].Content = "Total: $($Stats.TotalFiles)"
            })
        }
        
        if ($controls["ProcessedFilesLabel"]) {
            $controls["ProcessedFilesLabel"].Dispatcher.Invoke([action]{
                $controls["ProcessedFilesLabel"].Content = "Processed: $($Stats.ProcessedFiles)"
            })
        }
        
        if ($controls["ChunksLabel"]) {
            $controls["ChunksLabel"].Dispatcher.Invoke([action]{
                $controls["ChunksLabel"].Content = "Chunks: $($Stats.TotalChunks)"
            })
        }
        
        if ($controls["IndexedLabel"]) {
            $controls["IndexedLabel"].Dispatcher.Invoke([action]{
                $controls["IndexedLabel"].Content = "Indexed: $($Stats.IndexedDocuments ?? 0)"
            })
        }
    }

    # Function to create configuration from GUI inputs
    function Get-ConfigurationFromGUI {
        $azureServiceName = $controls["AzureServiceNameTextBox"].Text
        $azureApiKey = $controls["AzureApiKeyBox"].Password
        $indexName = $controls["IndexNameTextBox"].Text
        
        if ([string]::IsNullOrWhiteSpace($azureServiceName) -or [string]::IsNullOrWhiteSpace($azureApiKey)) {
            Write-GUILog "Azure configuration incomplete. Using default/test configuration." "WARN"
            $useAzure = $false
        } else {
            $useAzure = $true
        }
        
        return New-RAGConfiguration -ConfigurationName "GUI Generated Config" -AzureSearchServiceName ($azureServiceName ? $azureServiceName : "test-service") -AzureSearchApiKey ($azureApiKey ? $azureApiKey : "test-key") -IndexName ($indexName ? $indexName : "email-rag-index") -ProcessingSettings @{
            Chunking = @{
                Enabled = $controls["CreateRAGCheck"].IsChecked
                TargetTokens = 384
                MaxTokens = 512
                OverlapTokens = 32
            }
            ContentCleaning = @{
                Enabled = $controls["CleanContentCheck"].IsChecked
                RemoveSignatures = $true
                NormalizeWhitespace = $true
            }
            EntityExtraction = @{
                Enabled = $controls["ExtractEntitiesCheck"].IsChecked
            }
            Parallel = @{
                Enabled = $controls["ParallelProcessingCheck"].IsChecked
                MaxConcurrency = [int]$controls["ConcurrencySlider"].Value
            }
        }
    }

    # Event Handlers with REAL functionality
    Write-Host "🔧 Setting up REAL event handlers..." -ForegroundColor Gray

    # Browse button - File selection with analysis
    if ($controls["BrowseButton"]) {
        $controls["BrowseButton"].Add_Click({
            Write-GUILog "Opening folder browser..." "INFO"
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select folder containing MSG files for RAG processing"
            
            if ($folderDialog.ShowDialog() -eq "OK") {
                $selectedPath = $folderDialog.SelectedPath
                $controls["FilePathTextBox"].Text = $selectedPath
                Write-GUILog "Selected folder: $selectedPath" "INFO"
                
                # Analyze files in background
                try {
                    $msgFiles = Get-ChildItem -Path $selectedPath -Filter "*.msg" -Recurse -ErrorAction Stop
                    $count = $msgFiles.Count
                    $totalSize = ($msgFiles | Measure-Object -Property Length -Sum).Sum
                    
                    $controls["FileCountLabel"].Content = "📄 $count MSG files found"
                    $controls["FileCountLabel"].Foreground = $count -gt 0 ? "#FF28A745" : "#FFDC3545"
                    
                    if ($count -gt 0) {
                        $sizeMB = [Math]::Round($totalSize / 1MB, 2)
                        $estimatedChunks = [Math]::Ceiling($count * 2.5) # Rough estimate
                        $controls["FileAnalysisLabel"].Content = "📊 Analysis: $sizeMB MB total, ~$estimatedChunks chunks estimated"
                        $global:ProcessingStats.TotalFiles = $count
                        Update-ProcessingStats -Stats $global:ProcessingStats
                        Write-GUILog "File analysis: $count files, $sizeMB MB, estimated $estimatedChunks chunks" "SUCCESS"
                    } else {
                        $controls["FileAnalysisLabel"].Content = "⚠️ No MSG files found in selected folder"
                    }
                } catch {
                    Write-GUILog "Error analyzing files: $_" "ERROR"
                    $controls["FileCountLabel"].Content = "❌ Error accessing folder"
                    $controls["FileCountLabel"].Foreground = "#FFDC3545"
                    $controls["FileAnalysisLabel"].Content = "❌ Unable to analyze folder contents"
                }
            }
        })
    }

    # Concurrency slider update
    if ($controls["ConcurrencySlider"]) {
        $controls["ConcurrencySlider"].Add_ValueChanged({
            $controls["ConcurrencyLabel"].Content = [int]$controls["ConcurrencySlider"].Value
        })
    }

    # START PROCESSING - The main event with REAL functionality!
    if ($controls["StartProcessingButton"]) {
        $controls["StartProcessingButton"].Add_Click({
            try {
                $inputPath = $controls["FilePathTextBox"].Text
                
                if ([string]::IsNullOrWhiteSpace($inputPath) -or -not (Test-Path $inputPath)) {
                    [System.Windows.MessageBox]::Show(
                        "Please select a valid folder containing MSG files.",
                        "Invalid Path",
                        "OK",
                        "Warning"
                    )
                    return
                }
                
                Write-GUILog "🚀 STARTING REAL EMAIL PROCESSING..." "SUCCESS"
                
                # Create configuration from GUI
                $config = Get-ConfigurationFromGUI
                $global:CurrentConfig = $config
                
                # Reset processing stats
                $global:ProcessingStats = @{
                    TotalFiles = $global:ProcessingStats.TotalFiles
                    ProcessedFiles = 0
                    SuccessfulFiles = 0
                    FailedFiles = 0
                    TotalChunks = 0
                    IndexedDocuments = 0
                    StartTime = Get-Date
                    EndTime = $null
                    IsProcessing = $true
                }
                
                # Update UI state
                $controls["StartProcessingButton"].IsEnabled = $false
                $controls["StopProcessingButton"].IsEnabled = $true
                $controls["ProcessingProgressBar"].Value = 0
                $controls["ProgressLabel"].Content = "Initializing processing pipeline..."
                $controls["ProcessingTimeLabel"].Content = "Processing..."
                
                Write-GUILog "Configuration created: $($config.Metadata.Name)" "INFO"
                Write-GUILog "Processing settings: Clean=$($config.Processing.ContentCleaning.Enabled), RAG=$($config.Processing.Chunking.Enabled), Parallel=$($config.Processing.Parallel.Enabled)" "INFO"
                
                # Create progress callback
                $progressCallback = {
                    param($ProgressInfo)
                    
                    $controls["ProcessingProgressBar"].Dispatcher.Invoke([action]{
                        $progress = ($ProgressInfo.Current / [Math]::Max($ProgressInfo.Total, 1)) * 100
                        $controls["ProcessingProgressBar"].Value = $progress
                        
                        $eta = ""
                        if ($progress -gt 0 -and $global:ProcessingStats.StartTime) {
                            $elapsed = ((Get-Date) - $global:ProcessingStats.StartTime).TotalSeconds
                            $estimatedTotal = $elapsed * (100 / $progress)
                            $remaining = $estimatedTotal - $elapsed
                            $eta = " (ETA: $([Math]::Round($remaining, 0))s)"
                        }
                        
                        $controls["ProgressLabel"].Content = "$($ProgressInfo.Status) - $([Math]::Round($progress, 1))%$eta"
                    })
                    
                    Write-GUILog "Progress: $($ProgressInfo.Current)/$($ProgressInfo.Total) - $($ProgressInfo.Status)" "PROGRESS"
                }
                
                # Create status callback
                $statusCallback = {
                    param($StatusMessage)
                    Write-GUILog $StatusMessage "INFO"
                    
                    $controls["StatusBarLabel"].Dispatcher.Invoke([action]{
                        $controls["StatusBarLabel"].Content = $StatusMessage
                    })
                }
                
                # Start REAL processing in background
                Write-GUILog "Invoking EmailRAGProcessor with REAL MSG processing..." "INFO"
                
                $processingJob = Start-Job -ScriptBlock {
                    param($InputPath, $Configuration, $ModulePath)
                    
                    # Import modules in job context
                    Import-Module (Join-Path $ModulePath "RAGConfigManager_v2_Fixed.psm1") -Force
                    Import-Module (Join-Path $ModulePath "EmailRAGProcessor_v2_Fixed.psm1") -Force
                    
                    # Run the actual processing
                    return Invoke-EmailRAGProcessing -InputPath $InputPath -Configuration $Configuration -Parallel $Configuration.Processing.Parallel.Enabled -MaxConcurrency $Configuration.Processing.Parallel.MaxConcurrency
                    
                } -ArgumentList $inputPath, $config, $enhancedModulesPath
                
                $global:ProcessingJob = $processingJob
                
                # Monitor job progress
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromSeconds(1)
                
                $timer.Add_Tick({
                    if ($global:ProcessingJob) {
                        if ($global:ProcessingJob.State -eq "Completed") {
                            # Job completed - get results
                            $result = Receive-Job -Job $global:ProcessingJob -ErrorAction SilentlyContinue
                            Remove-Job -Job $global:ProcessingJob
                            $global:ProcessingJob = $null
                            
                            # Update final statistics
                            $global:ProcessingStats.EndTime = Get-Date
                            $global:ProcessingStats.IsProcessing = $false
                            
                            if ($result -and $result.Status -eq "Success") {
                                $stats = $result.Statistics
                                $global:ProcessingStats.ProcessedFiles = $stats.ProcessedFiles
                                $global:ProcessingStats.SuccessfulFiles = $stats.SuccessfulFiles
                                $global:ProcessingStats.FailedFiles = $stats.FailedFiles
                                $global:ProcessingStats.TotalChunks = $stats.TotalChunks
                                $global:ProcessingStats.IndexedDocuments = $stats.IndexedDocuments ?? 0
                                
                                Update-ProcessingStats -Stats $global:ProcessingStats
                                
                                $duration = ($global:ProcessingStats.EndTime - $global:ProcessingStats.StartTime).TotalSeconds
                                
                                Write-GUILog "🎉 PROCESSING COMPLETED SUCCESSFULLY!" "SUCCESS"
                                Write-GUILog "📊 Final Results: $($stats.SuccessfulFiles)/$($stats.TotalFiles) files processed, $($stats.TotalChunks) chunks, $($stats.IndexedDocuments) indexed" "SUCCESS"
                                Write-GUILog "⏱️ Total time: $([Math]::Round($duration, 2)) seconds" "SUCCESS"
                                
                                # Update results summary
                                $summaryText = "✅ Processing completed successfully!`n" +
                                               "📊 Files: $($stats.SuccessfulFiles)/$($stats.TotalFiles) processed`n" +
                                               "🎯 Chunks: $($stats.TotalChunks) generated`n" +
                                               "☁️ Indexed: $($stats.IndexedDocuments) documents`n" +
                                               "⏱️ Duration: $([Math]::Round($duration, 2)) seconds`n" +
                                               "🚀 Rate: $([Math]::Round($stats.SuccessfulFiles / [Math]::Max($duration, 1), 2)) files/sec"
                                
                                $controls["ResultsSummaryText"].Text = $summaryText
                                $controls["GenerateReportButton"].IsEnabled = $true
                                
                                $controls["ProgressLabel"].Content = "✅ Processing completed successfully!"
                                $controls["ProcessingProgressBar"].Value = 100
                                
                            } else {
                                Write-GUILog "❌ Processing failed: $($result.Error ?? 'Unknown error')" "ERROR"
                                $controls["ProgressLabel"].Content = "❌ Processing failed"
                                $controls["ResultsSummaryText"].Text = "❌ Processing failed: $($result.Error ?? 'Unknown error')"
                            }
                            
                            # Reset UI state
                            $controls["StartProcessingButton"].IsEnabled = $true
                            $controls["StopProcessingButton"].IsEnabled = $false
                            $controls["ProcessingTimeLabel"].Content = ""
                            
                            $this.Stop() # Stop the timer
                            
                        } elseif ($global:ProcessingJob.State -eq "Failed") {
                            # Job failed
                            $errorInfo = $global:ProcessingJob.ChildJobs[0].Error
                            Write-GUILog "❌ Processing job failed: $errorInfo" "ERROR"
                            
                            $controls["StartProcessingButton"].IsEnabled = $true
                            $controls["StopProcessingButton"].IsEnabled = $false
                            $controls["ProgressLabel"].Content = "❌ Processing failed"
                            $controls["ProcessingTimeLabel"].Content = ""
                            
                            Remove-Job -Job $global:ProcessingJob -Force
                            $global:ProcessingJob = $null
                            $this.Stop()
                        } else {
                            # Job still running - update elapsed time
                            if ($global:ProcessingStats.StartTime) {
                                $elapsed = ((Get-Date) - $global:ProcessingStats.StartTime).TotalSeconds
                                $controls["ProcessingTimeLabel"].Content = "⏱️ $([Math]::Round($elapsed, 0))s"
                            }
                        }
                    }
                })
                
                $timer.Start()
                
            } catch {
                Write-GUILog "❌ Failed to start processing: $_" "ERROR"
                [System.Windows.MessageBox]::Show(
                    "Failed to start processing:`n`n$_",
                    "Processing Error",
                    "OK",
                    "Error"
                )
                
                # Reset UI state
                $controls["StartProcessingButton"].IsEnabled = $true
                $controls["StopProcessingButton"].IsEnabled = $false
            }
        })
    }

    # Stop processing button
    if ($controls["StopProcessingButton"]) {
        $controls["StopProcessingButton"].Add_Click({
            if ($global:ProcessingJob) {
                Write-GUILog "Stopping processing..." "WARN"
                Stop-Job -Job $global:ProcessingJob
                Remove-Job -Job $global:ProcessingJob -Force
                $global:ProcessingJob = $null
                
                $controls["StartProcessingButton"].IsEnabled = $true
                $controls["StopProcessingButton"].IsEnabled = $false
                $controls["ProgressLabel"].Content = "❌ Processing stopped by user"
                $controls["ProcessingTimeLabel"].Content = ""
                
                Write-GUILog "Processing stopped by user" "WARN"
            }
        })
    }

    # Test configuration button  
    if ($controls["TestConfigButton"]) {
        $controls["TestConfigButton"].Add_Click({
            Write-GUILog "Testing configuration..." "INFO"
            
            try {
                $config = Get-ConfigurationFromGUI
                $testResult = Test-RAGConfiguration -Configuration $config -TestConnections $true -ValidateSettings $true
                
                $statusMsg = "Configuration test: $($testResult.OverallStatus) ($($testResult.TestsPassed)/$($testResult.TestsRun) passed)"
                Write-GUILog $statusMsg ($testResult.OverallStatus -eq "Passed" ? "SUCCESS" : "WARN")
                
                [System.Windows.MessageBox]::Show(
                    "Configuration Test Results:`n`n" +
                    "Status: $($testResult.OverallStatus)`n" +
                    "Tests Passed: $($testResult.TestsPassed)/$($testResult.TestsRun)`n" +
                    "Tests Failed: $($testResult.TestsFailed)",
                    "Configuration Test",
                    "OK",
                    ($testResult.OverallStatus -eq "Passed" ? "Information" : "Warning")
                )
                
            } catch {
                Write-GUILog "Configuration test failed: $_" "ERROR"
                [System.Windows.MessageBox]::Show(
                    "Configuration test failed:`n`n$_",
                    "Test Error",
                    "OK",
                    "Error"
                )
            }
        })
    }

    # Test Azure connection button
    if ($controls["TestConnectionButton"]) {
        $controls["TestConnectionButton"].Add_Click({
            Write-GUILog "Testing Azure Search connection..." "INFO"
            
            $serviceName = $controls["AzureServiceNameTextBox"].Text
            $apiKey = $controls["AzureApiKeyBox"].Password
            
            if ([string]::IsNullOrWhiteSpace($serviceName) -or [string]::IsNullOrWhiteSpace($apiKey)) {
                [System.Windows.MessageBox]::Show(
                    "Please enter both Service Name and API Key before testing connection.",
                    "Missing Configuration",
                    "OK",
                    "Warning"
                )
                return
            }
            
            try {
                $serviceUrl = "https://$serviceName.search.windows.net"
                $testUrl = "$serviceUrl/servicestats?api-version=2023-11-01"
                $headers = @{
                    'api-key' = $apiKey
                    'Content-Type' = 'application/json'
                }
                
                $response = Invoke-RestMethod -Uri $testUrl -Headers $headers -Method GET -TimeoutSec 15
                
                Write-GUILog "✅ Azure Search connection successful!" "SUCCESS"
                [System.Windows.MessageBox]::Show(
                    "Azure Search connection successful!`n`n" +
                    "Service: $serviceName`n" +
                    "Status: Connected",
                    "Connection Test",
                    "OK",
                    "Information"
                )
                
            } catch {
                Write-GUILog "❌ Azure Search connection failed: $_" "ERROR"
                [System.Windows.MessageBox]::Show(
                    "Azure Search connection failed:`n`n$_",
                    "Connection Test",
                    "OK",
                    "Error"
                )
            }
        })
    }

    # Clear logs button
    if ($controls["ClearLogsButton"]) {
        $controls["ClearLogsButton"].Add_Click({
            $controls["LogsTextBox"].Clear()
            Write-GUILog "Logs cleared" "INFO"
        })
    }

    # Export logs button
    if ($controls["ExportLogsButton"]) {
        $controls["ExportLogsButton"].Add_Click({
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            $saveDialog.FileName = "email-rag-cleaner-logs-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
            
            if ($saveDialog.ShowDialog() -eq "OK") {
                $controls["LogsTextBox"].Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                Write-GUILog "Logs exported to: $($saveDialog.FileName)" "SUCCESS"
            }
        })
    }

    # Generate report button
    if ($controls["GenerateReportButton"]) {
        $controls["GenerateReportButton"].Add_Click({
            if ($global:ProcessingStats -and $global:ProcessingStats.EndTime) {
                Write-GUILog "Generating detailed processing report..." "INFO"
                
                try {
                    # Import test framework module
                    Import-Module (Join-Path $enhancedModulesPath "RAGTestFramework_v2_Fixed.psm1") -Force
                    
                    # Create test results structure from processing stats
                    $testResults = @{
                        TestName = "Email Processing Report"
                        Configuration = $global:CurrentConfig.Metadata.Name
                        StartTime = $global:ProcessingStats.StartTime
                        EndTime = $global:ProcessingStats.EndTime
                        TotalDuration = ($global:ProcessingStats.EndTime - $global:ProcessingStats.StartTime).TotalSeconds
                        Results = @{
                            Processing = @{
                                Status = "Success"
                                FilesProcessed = $global:ProcessingStats.TotalFiles
                                SuccessfulFiles = $global:ProcessingStats.SuccessfulFiles
                                FailedFiles = $global:ProcessingStats.FailedFiles
                                TotalChunks = $global:ProcessingStats.TotalChunks
                                IndexedDocuments = $global:ProcessingStats.IndexedDocuments
                                ProcessingTime = ($global:ProcessingStats.EndTime - $global:ProcessingStats.StartTime).TotalSeconds
                            }
                        }
                        Summary = @{
                            OverallStatus = "Passed"
                            TotalTests = 1
                            TestsPassed = 1
                            TestsFailed = 0
                            SuccessRate = 100.0
                        }
                        Recommendations = @()
                    }
                    
                    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                    $saveDialog.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
                    $saveDialog.FileName = "email-rag-processing-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
                    
                    if ($saveDialog.ShowDialog() -eq "OK") {
                        $reportResult = Generate-TestReport -TestResults $testResults -OutputPath $saveDialog.FileName
                        Write-GUILog "Report generated: $($saveDialog.FileName)" "SUCCESS"
                        
                        # Ask if user wants to open the report
                        $result = [System.Windows.MessageBox]::Show(
                            "Report generated successfully!`n`nWould you like to open it now?",
                            "Report Generated",
                            "YesNo",
                            "Information"
                        )
                        
                        if ($result -eq "Yes") {
                            Start-Process $saveDialog.FileName
                        }
                    }
                    
                } catch {
                    Write-GUILog "Failed to generate report: $_" "ERROR"
                    [System.Windows.MessageBox]::Show(
                        "Failed to generate report:`n`n$_",
                        "Report Error",
                        "OK",
                        "Error"
                    )
                }
            }
        })
    }

    # Initialize UI
    Write-GUILog "🚀 Email RAG Cleaner v2.0 FUNCTIONAL GUI started successfully!" "SUCCESS"
    Write-GUILog "Installation path: $installPath" "INFO"
    Write-GUILog "🔥 Enhanced modules loaded: $($loadedModules -join ', ')" "SUCCESS"
    
    if ($moduleLoadErrors.Count -gt 0) {
        Write-GUILog "⚠️ Some optional modules failed to load:" "WARN"
        foreach ($moduleError in $moduleLoadErrors) {
            Write-GUILog "  - $moduleError" "WARN"
        }
    } else {
        Write-GUILog "✅ All modules loaded successfully - FULL functionality available!" "SUCCESS"
    }

    # Set module status indicator
    $controls["ModuleStatusLabel"].Content = "✅ $($loadedModules.Count) Modules Loaded"
    $controls["ModuleStatusLabel"].Foreground = "#FF28A745"

    # Handle window closing
    $window.Add_Closing({
        Write-Host "🚪 GUI window closing..." -ForegroundColor Yellow
        
        # Clean up any running jobs
        if ($global:ProcessingJob) {
            Write-Host "Stopping background processing job..." -ForegroundColor Yellow
            Stop-Job -Job $global:ProcessingJob -ErrorAction SilentlyContinue
            Remove-Job -Job $global:ProcessingJob -Force -ErrorAction SilentlyContinue
        }
    })

    # Show the window
    Write-Host "🎉 Showing FUNCTIONAL window with REAL processing capabilities..." -ForegroundColor Green
    
    # Use ShowDialog to keep the window open
    $null = $window.ShowDialog()

} catch {
    Write-Host "❌ FATAL ERROR: $_" -ForegroundColor Red
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