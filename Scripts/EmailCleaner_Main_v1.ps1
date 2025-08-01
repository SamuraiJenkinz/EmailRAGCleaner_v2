# MSG Email Cleaner v1.0 - Production GUI Application
# Professional email processing system with Azure integration

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import required modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulesPath = Join-Path $scriptPath "Modules"

try {
    Import-Module (Join-Path $modulesPath "ConfigManager.psm1") -Force
    Import-Module (Join-Path $modulesPath "MsgProcessor.psm1") -Force
    Import-Module (Join-Path $modulesPath "ContentCleaner.psm1") -Force
    Import-Module (Join-Path $modulesPath "AzureUploader.psm1") -Force
    Import-Module (Join-Path $modulesPath "AzureFlattener.psm1") -Force
    Write-Host "✅ All modules loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to load modules: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Global variables
$global:ProcessingStats = @{
    TotalFiles = 0
    ProcessedFiles = 0
    SuccessfulFiles = 0
    FailedFiles = 0
    TotalChunks = 0
    StartTime = $null
    EndTime = $null
}

$global:ConfigPath = Join-Path $scriptPath "Config\settings.json"
$global:LogPath = Join-Path $scriptPath "Logs"
$global:OutputPath = Join-Path $scriptPath "Output"

# Ensure required directories exist
@($global:LogPath, $global:OutputPath) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# Initialize logging
function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
    }
    Write-Host $logMessage -ForegroundColor $color
    
    # File output
    $logFile = Join-Path $global:LogPath "EmailCleaner_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logMessage
    
    # Update GUI log if available
    if ($global:LogTextBox) {
        $global:LogTextBox.AppendText("$logMessage`r`n")
        $global:LogTextBox.SelectionStart = $global:LogTextBox.Text.Length
        $global:LogTextBox.ScrollToCaret()
        $global:LogTextBox.Refresh()
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSG Email Cleaner v1.0 - Production System"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$form.Controls.Add($tabControl)

# Processing Tab
$processingTab = New-Object System.Windows.Forms.TabPage
$processingTab.Text = "Email Processing"
$processingTab.UseVisualStyleBackColor = $true
$tabControl.TabPages.Add($processingTab)

# File selection group
$fileGroup = New-Object System.Windows.Forms.GroupBox
$fileGroup.Text = "File Selection"
$fileGroup.Location = New-Object System.Drawing.Point(10, 10)
$fileGroup.Size = New-Object System.Drawing.Size(950, 100)
$processingTab.Controls.Add($fileGroup)

# File path textbox
$filePathLabel = New-Object System.Windows.Forms.Label
$filePathLabel.Text = "MSG Files Path:"
$filePathLabel.Location = New-Object System.Drawing.Point(10, 25)
$filePathLabel.Size = New-Object System.Drawing.Size(100, 20)
$fileGroup.Controls.Add($filePathLabel)

$filePathTextBox = New-Object System.Windows.Forms.TextBox
$filePathTextBox.Location = New-Object System.Drawing.Point(120, 22)
$filePathTextBox.Size = New-Object System.Drawing.Size(600, 20)
$fileGroup.Controls.Add($filePathTextBox)

# Browse button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Location = New-Object System.Drawing.Point(730, 20)
$browseButton.Size = New-Object System.Drawing.Size(80, 25)
$fileGroup.Controls.Add($browseButton)

# File count label
$fileCountLabel = New-Object System.Windows.Forms.Label
$fileCountLabel.Text = "No files selected"
$fileCountLabel.Location = New-Object System.Drawing.Point(120, 50)
$fileCountLabel.Size = New-Object System.Drawing.Size(600, 20)
$fileCountLabel.ForeColor = "Gray"
$fileGroup.Controls.Add($fileCountLabel)

# Processing options group
$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Text = "Processing Options"
$optionsGroup.Location = New-Object System.Drawing.Point(10, 120)
$optionsGroup.Size = New-Object System.Drawing.Size(950, 120)
$processingTab.Controls.Add($optionsGroup)

# Clean content checkbox
$cleanContentCheckBox = New-Object System.Windows.Forms.CheckBox
$cleanContentCheckBox.Text = "Clean and normalize content"
$cleanContentCheckBox.Location = New-Object System.Drawing.Point(15, 25)
$cleanContentCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$cleanContentCheckBox.Checked = $true
$optionsGroup.Controls.Add($cleanContentCheckBox)

# Extract entities checkbox
$extractEntitiesCheckBox = New-Object System.Windows.Forms.CheckBox
$extractEntitiesCheckBox.Text = "Extract entities (emails, URLs, phones)"
$extractEntitiesCheckBox.Location = New-Object System.Drawing.Point(230, 25)
$extractEntitiesCheckBox.Size = New-Object System.Drawing.Size(250, 20)
$extractEntitiesCheckBox.Checked = $true
$optionsGroup.Controls.Add($extractEntitiesCheckBox)

# Create RAG chunks checkbox
$createRAGCheckBox = New-Object System.Windows.Forms.CheckBox
$createRAGCheckBox.Text = "Create RAG chunks for AI/ML"
$createRAGCheckBox.Location = New-Object System.Drawing.Point(500, 25)
$createRAGCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$createRAGCheckBox.Checked = $true
$optionsGroup.Controls.Add($createRAGCheckBox)

# Upload to Azure checkbox
$uploadAzureCheckBox = New-Object System.Windows.Forms.CheckBox
$uploadAzureCheckBox.Text = "Upload to Azure Blob Storage"
$uploadAzureCheckBox.Location = New-Object System.Drawing.Point(15, 55)
$uploadAzureCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$uploadAzureCheckBox.Checked = $false
$optionsGroup.Controls.Add($uploadAzureCheckBox)

# Generate flattened JSON checkbox
$flattenJSONCheckBox = New-Object System.Windows.Forms.CheckBox
$flattenJSONCheckBox.Text = "Generate flattened JSON for vector databases"
$flattenJSONCheckBox.Location = New-Object System.Drawing.Point(230, 55)
$flattenJSONCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$flattenJSONCheckBox.Checked = $true
$optionsGroup.Controls.Add($flattenJSONCheckBox)

# Chunk size controls
$chunkSizeLabel = New-Object System.Windows.Forms.Label
$chunkSizeLabel.Text = "Chunk Size:"
$chunkSizeLabel.Location = New-Object System.Drawing.Point(15, 85)
$chunkSizeLabel.Size = New-Object System.Drawing.Size(80, 20)
$optionsGroup.Controls.Add($chunkSizeLabel)

$chunkSizeTextBox = New-Object System.Windows.Forms.TextBox
$chunkSizeTextBox.Text = "512"
$chunkSizeTextBox.Location = New-Object System.Drawing.Point(100, 82)
$chunkSizeTextBox.Size = New-Object System.Drawing.Size(60, 20)
$optionsGroup.Controls.Add($chunkSizeTextBox)

$chunkOverlapLabel = New-Object System.Windows.Forms.Label
$chunkOverlapLabel.Text = "Overlap:"
$chunkOverlapLabel.Location = New-Object System.Drawing.Point(180, 85)
$chunkOverlapLabel.Size = New-Object System.Drawing.Size(60, 20)
$optionsGroup.Controls.Add($chunkOverlapLabel)

$chunkOverlapTextBox = New-Object System.Windows.Forms.TextBox
$chunkOverlapTextBox.Text = "50"
$chunkOverlapTextBox.Location = New-Object System.Drawing.Point(245, 82)
$chunkOverlapTextBox.Size = New-Object System.Drawing.Size(60, 20)
$optionsGroup.Controls.Add($chunkOverlapTextBox)

# Progress group
$progressGroup = New-Object System.Windows.Forms.GroupBox
$progressGroup.Text = "Processing Progress"
$progressGroup.Location = New-Object System.Drawing.Point(10, 250)
$progressGroup.Size = New-Object System.Drawing.Size(950, 100)
$processingTab.Controls.Add($progressGroup)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 25)
$progressBar.Size = New-Object System.Drawing.Size(800, 20)
$progressBar.Style = "Continuous"
$progressGroup.Controls.Add($progressBar)

# Progress label
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = "Ready to process"
$progressLabel.Location = New-Object System.Drawing.Point(15, 55)
$progressLabel.Size = New-Object System.Drawing.Size(800, 20)
$progressGroup.Controls.Add($progressLabel)

# Stats labels
$statsLabel = New-Object System.Windows.Forms.Label
$statsLabel.Text = "Files: 0 processed, 0 successful, 0 failed | Chunks: 0 | Time: 0s"
$statsLabel.Location = New-Object System.Drawing.Point(15, 75)
$statsLabel.Size = New-Object System.Drawing.Size(800, 20)
$statsLabel.ForeColor = "DarkBlue"
$progressGroup.Controls.Add($statsLabel)

# Control buttons
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start Processing"
$startButton.Location = New-Object System.Drawing.Point(830, 30)
$startButton.Size = New-Object System.Drawing.Size(100, 35)
$startButton.BackColor = "LightGreen"
$progressGroup.Controls.Add($startButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop"
$stopButton.Location = New-Object System.Drawing.Point(830, 70)
$stopButton.Size = New-Object System.Drawing.Size(100, 25)
$stopButton.Enabled = $false
$progressGroup.Controls.Add($stopButton)

# Results group
$resultsGroup = New-Object System.Windows.Forms.GroupBox
$resultsGroup.Text = "Processing Results"
$resultsGroup.Location = New-Object System.Drawing.Point(10, 360)
$resultsGroup.Size = New-Object System.Drawing.Size(950, 250)
$processingTab.Controls.Add($resultsGroup)

# Results list view
$resultsListView = New-Object System.Windows.Forms.ListView
$resultsListView.Location = New-Object System.Drawing.Point(15, 25)
$resultsListView.Size = New-Object System.Drawing.Size(920, 180)
$resultsListView.View = "Details"
$resultsListView.FullRowSelect = $true
$resultsListView.GridLines = $true

$resultsListView.Columns.Add("File Name", 200) | Out-Null
$resultsListView.Columns.Add("Status", 80) | Out-Null
$resultsListView.Columns.Add("Chunks", 60) | Out-Null
$resultsListView.Columns.Add("Size (KB)", 80) | Out-Null
$resultsListView.Columns.Add("Processing Time", 100) | Out-Null
$resultsListView.Columns.Add("Details", 300) | Out-Null

$resultsGroup.Controls.Add($resultsListView)

# Export results button
$exportResultsButton = New-Object System.Windows.Forms.Button
$exportResultsButton.Text = "Export Results"
$exportResultsButton.Location = New-Object System.Drawing.Point(15, 215)
$exportResultsButton.Size = New-Object System.Drawing.Size(120, 25)
$resultsGroup.Controls.Add($exportResultsButton)

# Configuration Tab
$configTab = New-Object System.Windows.Forms.TabPage
$configTab.Text = "Configuration"
$configTab.UseVisualStyleBackColor = $true
$tabControl.TabPages.Add($configTab)

# Azure configuration group
$azureConfigGroup = New-Object System.Windows.Forms.GroupBox
$azureConfigGroup.Text = "Azure Configuration"
$azureConfigGroup.Location = New-Object System.Drawing.Point(10, 10)
$azureConfigGroup.Size = New-Object System.Drawing.Size(950, 200)
$configTab.Controls.Add($azureConfigGroup)

# Connection string
$connectionStringLabel = New-Object System.Windows.Forms.Label
$connectionStringLabel.Text = "Connection String:"
$connectionStringLabel.Location = New-Object System.Drawing.Point(15, 30)
$connectionStringLabel.Size = New-Object System.Drawing.Size(120, 20)
$azureConfigGroup.Controls.Add($connectionStringLabel)

$connectionStringTextBox = New-Object System.Windows.Forms.TextBox
$connectionStringTextBox.Location = New-Object System.Drawing.Point(140, 27)
$connectionStringTextBox.Size = New-Object System.Drawing.Size(600, 20)
$connectionStringTextBox.PasswordChar = '*'
$azureConfigGroup.Controls.Add($connectionStringTextBox)

# Test connection button
$testConnectionButton = New-Object System.Windows.Forms.Button
$testConnectionButton.Text = "Test"
$testConnectionButton.Location = New-Object System.Drawing.Point(750, 25)
$testConnectionButton.Size = New-Object System.Drawing.Size(60, 25)
$azureConfigGroup.Controls.Add($testConnectionButton)

# Container name
$containerNameLabel = New-Object System.Windows.Forms.Label
$containerNameLabel.Text = "Container Name:"
$containerNameLabel.Location = New-Object System.Drawing.Point(15, 60)
$containerNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$azureConfigGroup.Controls.Add($containerNameLabel)

$containerNameTextBox = New-Object System.Windows.Forms.TextBox
$containerNameTextBox.Text = "email-data"
$containerNameTextBox.Location = New-Object System.Drawing.Point(140, 57)
$containerNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
$azureConfigGroup.Controls.Add($containerNameTextBox)

# Save and load config buttons
$saveConfigButton = New-Object System.Windows.Forms.Button
$saveConfigButton.Text = "Save Configuration"
$saveConfigButton.Location = New-Object System.Drawing.Point(15, 160)
$saveConfigButton.Size = New-Object System.Drawing.Size(150, 30)
$saveConfigButton.BackColor = "LightBlue"
$azureConfigGroup.Controls.Add($saveConfigButton)

$loadConfigButton = New-Object System.Windows.Forms.Button
$loadConfigButton.Text = "Load Configuration"
$loadConfigButton.Location = New-Object System.Drawing.Point(175, 160)
$loadConfigButton.Size = New-Object System.Drawing.Size(150, 30)
$azureConfigGroup.Controls.Add($loadConfigButton)

# Log Tab
$logTab = New-Object System.Windows.Forms.TabPage
$logTab.Text = "Logs"
$logTab.UseVisualStyleBackColor = $true
$tabControl.TabPages.Add($logTab)

# Log text box
$global:LogTextBox = New-Object System.Windows.Forms.TextBox
$global:LogTextBox.Multiline = $true
$global:LogTextBox.ScrollBars = "Vertical"
$global:LogTextBox.Location = New-Object System.Drawing.Point(10, 40)
$global:LogTextBox.Size = New-Object System.Drawing.Size(950, 520)
$global:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$global:LogTextBox.ReadOnly = $true
$logTab.Controls.Add($global:LogTextBox)

# Log controls
$clearLogButton = New-Object System.Windows.Forms.Button
$clearLogButton.Text = "Clear Log"
$clearLogButton.Location = New-Object System.Drawing.Point(10, 10)
$clearLogButton.Size = New-Object System.Drawing.Size(80, 25)
$logTab.Controls.Add($clearLogButton)

$exportLogButton = New-Object System.Windows.Forms.Button
$exportLogButton.Text = "Export Log"
$exportLogButton.Location = New-Object System.Drawing.Point(100, 10)
$exportLogButton.Size = New-Object System.Drawing.Size(80, 25)
$logTab.Controls.Add($exportLogButton)

# Event handlers
$browseButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder containing MSG files"
    
    if ($folderDialog.ShowDialog() -eq "OK") {
        $filePathTextBox.Text = $folderDialog.SelectedPath
        Update-FileCount
    }
})

function Update-FileCount {
    if ($filePathTextBox.Text -and (Test-Path $filePathTextBox.Text)) {
        $msgFiles = Get-ChildItem -Path $filePathTextBox.Text -Filter "*.msg" -Recurse
        $count = $msgFiles.Count
        $fileCountLabel.Text = "$count MSG files found"
        $fileCountLabel.ForeColor = if ($count -gt 0) { "Green" } else { "Red" }
        $global:ProcessingStats.TotalFiles = $count
    } else {
        $fileCountLabel.Text = "Invalid path or no files found"
        $fileCountLabel.ForeColor = "Red"
        $global:ProcessingStats.TotalFiles = 0
    }
}

$filePathTextBox.Add_TextChanged({ Update-FileCount })

$startButton.Add_Click({
    Start-Processing
})

$stopButton.Add_Click({
    Write-LogMessage "Processing stopped by user" "WARN"
    $global:StopProcessing = $true
})

$testConnectionButton.Add_Click({
    Test-AzureConnection
})

$saveConfigButton.Add_Click({
    Save-Configuration
})

$loadConfigButton.Add_Click({
    Load-Configuration
})

$exportResultsButton.Add_Click({
    Export-ProcessingResults
})

$clearLogButton.Add_Click({
    $global:LogTextBox.Clear()
})

$exportLogButton.Add_Click({
    Export-LogFile
})

# Processing functions
function Start-Processing {
    if (-not $filePathTextBox.Text -or -not (Test-Path $filePathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid folder containing MSG files.", "Invalid Path", "OK", "Warning")
        return
    }
    
    if ($global:ProcessingStats.TotalFiles -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No MSG files found in the selected folder.", "No Files", "OK", "Warning")
        return
    }
    
    # Reset stats
    $global:ProcessingStats.ProcessedFiles = 0
    $global:ProcessingStats.SuccessfulFiles = 0
    $global:ProcessingStats.FailedFiles = 0
    $global:ProcessingStats.TotalChunks = 0
    $global:ProcessingStats.StartTime = Get-Date
    $global:StopProcessing = $false
    
    # Update UI state
    $startButton.Enabled = $false
    $stopButton.Enabled = $true
    $progressBar.Value = 0
    $resultsListView.Items.Clear()
    
    Write-LogMessage "Starting processing of $($global:ProcessingStats.TotalFiles) MSG files" "INFO"
    
    # Get processing options
    $options = @{
        CleanContent = $cleanContentCheckBox.Checked
        ExtractEntities = $extractEntitiesCheckBox.Checked
        CreateRAGChunks = $createRAGCheckBox.Checked
        UploadToAzure = $uploadAzureCheckBox.Checked
        GenerateFlattenedJSON = $flattenJSONCheckBox.Checked
        ChunkSize = [int]$chunkSizeTextBox.Text
        ChunkOverlap = [int]$chunkOverlapTextBox.Text
    }
    
    # Process files
    $msgFiles = Get-ChildItem -Path $filePathTextBox.Text -Filter "*.msg" -Recurse
    
    foreach ($msgFile in $msgFiles) {
        if ($global:StopProcessing) {
            Write-LogMessage "Processing stopped" "WARN"
            break
        }
        
        Process-SingleFile -FilePath $msgFile.FullName -Options $options
        
        # Update progress
        $global:ProcessingStats.ProcessedFiles++
        $progressPercent = [math]::Round(($global:ProcessingStats.ProcessedFiles / $global:ProcessingStats.TotalFiles) * 100)
        $progressBar.Value = $progressPercent
        $progressLabel.Text = "Processing: $($msgFile.Name) ($($global:ProcessingStats.ProcessedFiles)/$($global:ProcessingStats.TotalFiles))"
        
        Update-StatsDisplay
        $form.Refresh()
    }
    
    # Finish processing
    $global:ProcessingStats.EndTime = Get-Date
    $duration = ($global:ProcessingStats.EndTime - $global:ProcessingStats.StartTime).TotalSeconds
    
    Write-LogMessage "Processing completed in $([math]::Round($duration, 2)) seconds" "SUCCESS"
    Write-LogMessage "Results: $($global:ProcessingStats.SuccessfulFiles) successful, $($global:ProcessingStats.FailedFiles) failed" "INFO"
    
    $progressLabel.Text = "Processing completed"
    $startButton.Enabled = $true
    $stopButton.Enabled = $false
}

function Process-SingleFile {
    param(
        [string]$FilePath,
        [hashtable]$Options
    )
    
    $fileName = Split-Path -Leaf $FilePath
    $startTime = Get-Date
    
    try {
        Write-LogMessage "Processing file: $fileName" "INFO"
        
        # Process MSG file
        $msgData = Read-MsgFile -FilePath $FilePath
        if (-not $msgData) {
            throw "Failed to read MSG file"
        }
        
        $fileSize = [math]::Round((Get-Item $FilePath).Length / 1KB, 1)
        $chunkCount = 0
        
        # Clean content if requested
        if ($Options.CleanContent) {
            $msgData = Clean-EmailContent -EmailData $msgData -ExtractEntities:$Options.ExtractEntities
        }
        
        # Create RAG chunks if requested
        if ($Options.CreateRAGChunks) {
            $msgData = Add-RAGChunks -EmailData $msgData -ChunkSize $Options.ChunkSize -Overlap $Options.ChunkOverlap
            $chunkCount = if ($msgData.Content.Chunks) { $msgData.Content.Chunks.Count } else { 0 }
            $global:ProcessingStats.TotalChunks += $chunkCount
        }
        
        # Generate flattened JSON if requested
        if ($Options.GenerateFlattenedJSON) {
            $flattenedData = ConvertTo-AzureSearchFormat -EmailData $msgData
            $outputDir = Join-Path $global:OutputPath "Azure"
            if (-not (Test-Path $outputDir)) {
                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
            }
            
            $jsonFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + "_flattened.json"
            $jsonPath = Join-Path $outputDir $jsonFileName
            $flattenedData | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding UTF8
        }
        
        # Upload to Azure if requested
        if ($Options.UploadToAzure -and $connectionStringTextBox.Text) {
            $blobName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".json"
            $jsonData = $msgData | ConvertTo-Json -Depth 10
            $tempFile = [System.IO.Path]::GetTempFileName()
            $jsonData | Set-Content -Path $tempFile -Encoding UTF8
            
            try {
                Upload-ToAzureBlob -FilePath $tempFile -BlobName $blobName -ConnectionString $connectionStringTextBox.Text -ContainerName $containerNameTextBox.Text
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            } catch {
                Write-LogMessage "Azure upload failed for $fileName`: $($_.Exception.Message)" "ERROR"
            }
        }
        
        # Add to results
        $processingTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
        $listItem = New-Object System.Windows.Forms.ListViewItem($fileName)
        $listItem.SubItems.Add("Success") | Out-Null
        $listItem.SubItems.Add($chunkCount.ToString()) | Out-Null
        $listItem.SubItems.Add($fileSize.ToString()) | Out-Null
        $listItem.SubItems.Add("$processingTime`s") | Out-Null
        $listItem.SubItems.Add("Processed successfully") | Out-Null
        $listItem.ForeColor = "Green"
        $resultsListView.Items.Add($listItem) | Out-Null
        
        $global:ProcessingStats.SuccessfulFiles++
        Write-LogMessage "Successfully processed $fileName ($chunkCount chunks, $fileSize KB)" "SUCCESS"
        
    } catch {
        $processingTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
        $errorMessage = $_.Exception.Message
        
        $listItem = New-Object System.Windows.Forms.ListViewItem($fileName)
        $listItem.SubItems.Add("Failed") | Out-Null
        $listItem.SubItems.Add("0") | Out-Null
        $listItem.SubItems.Add("0") | Out-Null
        $listItem.SubItems.Add("$processingTime`s") | Out-Null
        $listItem.SubItems.Add($errorMessage) | Out-Null
        $listItem.ForeColor = "Red"
        $resultsListView.Items.Add($listItem) | Out-Null
        
        $global:ProcessingStats.FailedFiles++
        Write-LogMessage "Failed to process $fileName`: $errorMessage" "ERROR"
    }
}

function Update-StatsDisplay {
    $duration = if ($global:ProcessingStats.StartTime) {
        [math]::Round(((Get-Date) - $global:ProcessingStats.StartTime).TotalSeconds, 1)
    } else { 0 }
    
    $statsLabel.Text = "Files: $($global:ProcessingStats.ProcessedFiles) processed, $($global:ProcessingStats.SuccessfulFiles) successful, $($global:ProcessingStats.FailedFiles) failed | Chunks: $($global:ProcessingStats.TotalChunks) | Time: $($duration)s"
}

function Test-AzureConnection {
    if (-not $connectionStringTextBox.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please enter an Azure connection string.", "Missing Connection String", "OK", "Warning")
        return
    }
    
    try {
        Write-LogMessage "Testing Azure connection..." "INFO"
        $testResult = Test-AzureConnection -ConnectionString $connectionStringTextBox.Text -ContainerName $containerNameTextBox.Text
        
        if ($testResult) {
            [System.Windows.Forms.MessageBox]::Show("Azure connection successful!", "Connection Test", "OK", "Information")
            Write-LogMessage "Azure connection test successful" "SUCCESS"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Azure connection failed. Check your connection string and container name.", "Connection Test Failed", "OK", "Error")
            Write-LogMessage "Azure connection test failed" "ERROR"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Azure connection error: $($_.Exception.Message)", "Connection Test Error", "OK", "Error")
        Write-LogMessage "Azure connection error: $($_.Exception.Message)" "ERROR"
    }
}

function Save-Configuration {
    try {
        $config = @{
            Azure = @{
                ConnectionString = $connectionStringTextBox.Text
                ContainerName = $containerNameTextBox.Text
            }
            Processing = @{
                ChunkSize = [int]$chunkSizeTextBox.Text
                ChunkOverlap = [int]$chunkOverlapTextBox.Text
                CleanContent = $cleanContentCheckBox.Checked
                ExtractEntities = $extractEntitiesCheckBox.Checked
                CreateRAGChunks = $createRAGCheckBox.Checked
                UploadToAzure = $uploadAzureCheckBox.Checked
                GenerateFlattenedJSON = $flattenJSONCheckBox.Checked
            }
        }
        
        Save-Config -ConfigPath $global:ConfigPath -ConfigData $config
        [System.Windows.Forms.MessageBox]::Show("Configuration saved successfully!", "Save Configuration", "OK", "Information")
        Write-LogMessage "Configuration saved to $global:ConfigPath" "SUCCESS"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to save configuration: $($_.Exception.Message)", "Save Error", "OK", "Error")
        Write-LogMessage "Failed to save configuration: $($_.Exception.Message)" "ERROR"
    }
}

function Load-Configuration {
    try {
        if (-not (Test-Path $global:ConfigPath)) {
            [System.Windows.Forms.MessageBox]::Show("Configuration file not found. Using defaults.", "Load Configuration", "OK", "Information")
            return
        }
        
        $config = Load-Config -ConfigPath $global:ConfigPath
        
        if ($config.Azure) {
            $connectionStringTextBox.Text = $config.Azure.ConnectionString
            $containerNameTextBox.Text = $config.Azure.ContainerName
        }
        
        if ($config.Processing) {
            $chunkSizeTextBox.Text = $config.Processing.ChunkSize.ToString()
            $chunkOverlapTextBox.Text = $config.Processing.ChunkOverlap.ToString()
            $cleanContentCheckBox.Checked = $config.Processing.CleanContent
            $extractEntitiesCheckBox.Checked = $config.Processing.ExtractEntities
            $createRAGCheckBox.Checked = $config.Processing.CreateRAGChunks
            $uploadAzureCheckBox.Checked = $config.Processing.UploadToAzure
            $flattenJSONCheckBox.Checked = $config.Processing.GenerateFlattenedJSON
        }
        
        [System.Windows.Forms.MessageBox]::Show("Configuration loaded successfully!", "Load Configuration", "OK", "Information")
        Write-LogMessage "Configuration loaded from $global:ConfigPath" "SUCCESS"
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load configuration: $($_.Exception.Message)", "Load Error", "OK", "Error")
        Write-LogMessage "Failed to load configuration: $($_.Exception.Message)" "ERROR"
    }
}

function Export-ProcessingResults {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveDialog.DefaultExt = "csv"
    $saveDialog.FileName = "EmailProcessingResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $results = @()
            foreach ($item in $resultsListView.Items) {
                $results += [PSCustomObject]@{
                    FileName = $item.Text
                    Status = $item.SubItems[1].Text
                    Chunks = $item.SubItems[2].Text
                    SizeKB = $item.SubItems[3].Text
                    ProcessingTime = $item.SubItems[4].Text
                    Details = $item.SubItems[5].Text
                }
            }
            
            $results | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Results exported successfully!", "Export Results", "OK", "Information")
            Write-LogMessage "Results exported to $($saveDialog.FileName)" "SUCCESS"
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export results: $($_.Exception.Message)", "Export Error", "OK", "Error")
            Write-LogMessage "Failed to export results: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Export-LogFile {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDialog.DefaultExt = "log"
    $saveDialog.FileName = "EmailCleaner_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $global:LogTextBox.Text | Set-Content -Path $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("Log exported successfully!", "Export Log", "OK", "Information")
            Write-LogMessage "Log exported to $($saveDialog.FileName)" "SUCCESS"
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export log: $($_.Exception.Message)", "Export Error", "OK", "Error")
            Write-LogMessage "Failed to export log: $($_.Exception.Message)" "ERROR"
        }
    }
}

# Initialize application
Write-LogMessage "MSG Email Cleaner v1.0 started" "INFO"
Write-LogMessage "GUI initialized successfully" "SUCCESS"

# Load configuration on startup
if (Test-Path $global:ConfigPath) {
    Load-Configuration
}

# Show the form
$form.ShowDialog() | Out-Null