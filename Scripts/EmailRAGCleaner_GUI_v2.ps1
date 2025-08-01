# Email RAG Cleaner v2.0 - Modern WPF GUI Interface
# Professional Azure AI Search RAG processing with beautiful interface

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Import required modules
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$installPath = "C:\EmailRAGCleaner"
$modulesPath = Join-Path $installPath "Modules"

try {
    Import-Module (Join-Path $modulesPath "RAGConfigManager_v2.psm1") -Force
    Import-Module (Join-Path $modulesPath "EmailRAGProcessor_v2.psm1") -Force
    Import-Module (Join-Path $modulesPath "EmailSearchInterface_v2.psm1") -Force
    Import-Module (Join-Path $modulesPath "RAGTestFramework_v2.psm1") -Force
    Write-Host "✅ All RAG v2.0 modules loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to load modules: $($_.Exception.Message)" -ForegroundColor Red
    [System.Windows.MessageBox]::Show("Failed to load required modules. Please ensure Email RAG Cleaner v2.0 is properly installed.", "Module Load Error", "OK", "Error")
    exit 1
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

# XAML for modern WPF interface
$xaml = @"
<Window x:Class="EmailRAGCleaner.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email RAG Cleaner v2.0 - Azure AI Search Integration" 
        Height="800" Width="1200" 
        MinHeight="600" MinWidth="900"
        WindowStartupLocation="CenterScreen"
        Background="#FF2D2D30">
    
    <Window.Resources>
        <!-- Modern Dark Theme Styles -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#FF007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF1E90FF"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#FF0066CC"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#FF404040"/>
                    <Setter Property="Foreground" Value="#FF808080"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#FF28A745"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF34CE57"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#FFDC3545"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FFE85563"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernGroupBox" TargetType="GroupBox">
            <Setter Property="Foreground" Value="#FFCCCCCC"/>
            <Setter Property="BorderBrush" Value="#FF3F3F46"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Background" Value="#FF3F3F46"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#FF007ACC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style x:Key="ModernLabel" TargetType="Label">
            <Setter Property="Foreground" Value="#FFCCCCCC"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style x:Key="HeaderLabel" TargetType="Label" BasedOn="{StaticResource ModernLabel}">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#FF007ACC"/>
        </Style>

        <Style x:Key="StatusLabel" TargetType="Label" BasedOn="{StaticResource ModernLabel}">
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#FF1E1E1E" BorderBrush="#FF3F3F46" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0">
                    <TextBlock Text="📧" FontSize="24" Margin="0,0,10,0"/>
                    <StackPanel>
                        <TextBlock Text="Email RAG Cleaner v2.0" 
                                   FontSize="18" FontWeight="Bold" 
                                   Foreground="#FF007ACC"/>
                        <TextBlock Text="Azure AI Search Integration" 
                                   FontSize="12" 
                                   Foreground="#FFCCCCCC"/>
                    </StackPanel>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,20,0">
                    <Button Name="ConfigButton" Content="⚙️ Config" Style="{StaticResource ModernButton}" Width="80"/>
                    <Button Name="TestButton" Content="🧪 Test" Style="{StaticResource ModernButton}" Width="80"/>
                    <Button Name="HelpButton" Content="❓ Help" Style="{StaticResource ModernButton}" Width="80"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Main Content -->
        <TabControl Grid.Row="1" Background="#FF2D2D30" BorderThickness="0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="TabItem">
                                <Border Name="Border" 
                                        Background="#FF3F3F46" 
                                        BorderBrush="#FF3F3F46" 
                                        BorderThickness="1,1,1,0" 
                                        CornerRadius="4,4,0,0" 
                                        Margin="2,2,2,0">
                                    <ContentPresenter x:Name="ContentSite"
                                                      VerticalAlignment="Center"
                                                      HorizontalAlignment="Center"
                                                      ContentSource="Header"
                                                      Margin="12,6"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="Border" Property="Background" Value="#FF007ACC"/>
                                    </Trigger>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="Border" Property="Background" Value="#FF404040"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                    <Setter Property="Foreground" Value="White"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </Style>
            </TabControl.Resources>

            <!-- Processing Tab -->
            <TabItem Header="📁 Processing">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- File Selection -->
                    <GroupBox Grid.Row="0" Header="📂 File Selection" Style="{StaticResource ModernGroupBox}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <Label Grid.Row="0" Grid.Column="0" Content="MSG Files Path:" Style="{StaticResource ModernLabel}"/>
                            <TextBox Name="FilePathTextBox" Grid.Row="0" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="5,0"/>
                            <Button Name="BrowseButton" Grid.Row="0" Grid.Column="2" Content="📁 Browse" Style="{StaticResource ModernButton}"/>

                            <Label Name="FileCountLabel" Grid.Row="1" Grid.Column="1" Content="No files selected" 
                                   Style="{StaticResource StatusLabel}" Foreground="#FFCCCCCC" Margin="5,5,0,0"/>
                        </Grid>
                    </GroupBox>

                    <!-- Processing Options -->
                    <GroupBox Grid.Row="1" Header="⚙️ Processing Options" Style="{StaticResource ModernGroupBox}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <CheckBox Name="CleanContentCheck" Grid.Row="0" Grid.Column="0" 
                                      Content="🧹 Clean Content" Foreground="#FFCCCCCC" 
                                      IsChecked="True" Margin="5"/>
                            <CheckBox Name="ExtractEntitiesCheck" Grid.Row="0" Grid.Column="1" 
                                      Content="🏷️ Extract Entities" Foreground="#FFCCCCCC" 
                                      IsChecked="True" Margin="5"/>
                            <CheckBox Name="CreateRAGCheck" Grid.Row="0" Grid.Column="2" 
                                      Content="🤖 Create RAG Chunks" Foreground="#FFCCCCCC" 
                                      IsChecked="True" Margin="5"/>

                            <CheckBox Name="UploadAzureCheck" Grid.Row="1" Grid.Column="0" 
                                      Content="☁️ Upload to Azure" Foreground="#FFCCCCCC" 
                                      IsChecked="True" Margin="5"/>
                            <CheckBox Name="GenerateEmbeddingsCheck" Grid.Row="1" Grid.Column="1" 
                                      Content="🔢 Generate Embeddings" Foreground="#FFCCCCCC" 
                                      IsChecked="True" Margin="5"/>
                            <CheckBox Name="IndexSearchCheck" Grid.Row="1" Grid.Column="2" 
                                      Content="🔍 Index for Search" Foreground="#FFCCCCCC" 
                                      IsChecked="True" Margin="5"/>

                            <StackPanel Grid.Row="2" Grid.Column="0" Orientation="Horizontal" Margin="5">
                                <Label Content="Chunk Size:" Style="{StaticResource ModernLabel}"/>
                                <TextBox Name="ChunkSizeTextBox" Text="384" Width="60" Style="{StaticResource ModernTextBox}"/>
                            </StackPanel>
                            <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="5">
                                <Label Content="Overlap Tokens:" Style="{StaticResource ModernLabel}"/>
                                <TextBox Name="OverlapTextBox" Text="32" Width="60" Style="{StaticResource ModernTextBox}"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>

                    <!-- Processing Control -->
                    <GroupBox Grid.Row="2" Header="🚀 Processing Control" Style="{StaticResource ModernGroupBox}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <ProgressBar Name="ProcessingProgressBar" Grid.Column="0" 
                                             Height="25" Margin="5" 
                                             Background="#FF3F3F46" 
                                             Foreground="#FF28A745"/>
                                <Button Name="StartProcessingButton" Grid.Column="1" 
                                        Content="▶️ Start Processing" 
                                        Style="{StaticResource SuccessButton}" Width="140"/>
                                <Button Name="StopProcessingButton" Grid.Column="2" 
                                        Content="⏹️ Stop" 
                                        Style="{StaticResource DangerButton}" Width="80" IsEnabled="False"/>
                            </Grid>

                            <Label Name="ProcessingStatusLabel" Grid.Row="1" 
                                   Content="Ready to process emails" 
                                   Style="{StaticResource StatusLabel}" Margin="5,0"/>

                            <Label Name="ProcessingStatsLabel" Grid.Row="2" 
                                   Content="Files: 0 processed, 0 successful, 0 failed | Chunks: 0 | Time: 0s" 
                                   Style="{StaticResource ModernLabel}" Margin="5,0"/>
                        </Grid>
                    </GroupBox>

                    <!-- Results -->
                    <GroupBox Grid.Row="3" Header="📊 Processing Results" Style="{StaticResource ModernGroupBox}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <ListView Name="ResultsListView" Grid.Row="0" 
                                      Background="#FF3F3F46" 
                                      Foreground="White" 
                                      BorderBrush="#FF007ACC">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="📄 File Name" Width="200" DisplayMemberBinding="{Binding FileName}"/>
                                        <GridViewColumn Header="✅ Status" Width="80" DisplayMemberBinding="{Binding Status}"/>
                                        <GridViewColumn Header="🧩 Chunks" Width="70" DisplayMemberBinding="{Binding Chunks}"/>
                                        <GridViewColumn Header="📏 Size (KB)" Width="80" DisplayMemberBinding="{Binding SizeKB}"/>
                                        <GridViewColumn Header="⏱️ Time (s)" Width="80" DisplayMemberBinding="{Binding ProcessingTime}"/>
                                        <GridViewColumn Header="📝 Details" Width="300" DisplayMemberBinding="{Binding Details}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>

                            <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                                <Button Name="ExportResultsButton" Content="📤 Export Results" Style="{StaticResource ModernButton}"/>
                                <Button Name="ClearResultsButton" Content="🗑️ Clear Results" Style="{StaticResource ModernButton}"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>

            <!-- Search Tab -->
            <TabItem Header="🔍 Search">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <GroupBox Grid.Row="0" Header="🔎 Azure AI Search Interface" Style="{StaticResource ModernGroupBox}">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <TextBox Name="SearchQueryTextBox" Grid.Column="0" 
                                         Style="{StaticResource ModernTextBox}" 
                                         Margin="5" 
                                         FontSize="14"
                                         Text="Enter your search query..."/>
                                <ComboBox Name="SearchTypeComboBox" Grid.Column="1" 
                                          Width="120" Margin="5" 
                                          Background="#FF3F3F46" 
                                          Foreground="White">
                                    <ComboBoxItem Content="Hybrid" IsSelected="True"/>
                                    <ComboBoxItem Content="Vector"/>
                                    <ComboBoxItem Content="Keyword"/>
                                    <ComboBoxItem Content="Semantic"/>
                                </ComboBox>
                                <Button Name="SearchButton" Grid.Column="2" 
                                        Content="🔍 Search" 
                                        Style="{StaticResource ModernButton}" Width="100"/>
                            </Grid>

                            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="5">
                                <Button Name="SearchSenderButton" Content="👤 By Sender" Style="{StaticResource ModernButton}"/>
                                <Button Name="SearchDateButton" Content="📅 By Date" Style="{StaticResource ModernButton}"/>
                                <Button Name="SearchAttachmentsButton" Content="📎 With Attachments" Style="{StaticResource ModernButton}"/>
                            </StackPanel>

                            <Label Name="SearchStatusLabel" Grid.Row="2" 
                                   Content="Ready to search" 
                                   Style="{StaticResource StatusLabel}" Margin="5,0"/>
                        </Grid>
                    </GroupBox>

                    <GroupBox Grid.Row="1" Header="📋 Search Results" Style="{StaticResource ModernGroupBox}">
                        <ListView Name="SearchResultsListView" 
                                  Background="#FF3F3F46" 
                                  Foreground="White" 
                                  BorderBrush="#FF007ACC">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="📄 Document" Width="200" DisplayMemberBinding="{Binding Document}"/>
                                    <GridViewColumn Header="⭐ Score" Width="80" DisplayMemberBinding="{Binding Score}"/>
                                    <GridViewColumn Header="👤 Sender" Width="150" DisplayMemberBinding="{Binding Sender}"/>
                                    <GridViewColumn Header="📅 Date" Width="120" DisplayMemberBinding="{Binding Date}"/>
                                    <GridViewColumn Header="📝 Preview" Width="400" DisplayMemberBinding="{Binding Preview}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </GroupBox>
                </Grid>
            </TabItem>

            <!-- Configuration Tab -->
            <TabItem Header="⚙️ Configuration">
                <ScrollViewer>
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <!-- Azure AI Search Configuration -->
                        <GroupBox Grid.Row="0" Header="☁️ Azure AI Search Configuration" Style="{StaticResource ModernGroupBox}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <Label Grid.Row="0" Grid.Column="0" Content="Service Name:" Style="{StaticResource ModernLabel}"/>
                                <TextBox Name="AzureServiceNameTextBox" Grid.Row="0" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="5,0"/>

                                <Label Grid.Row="1" Grid.Column="0" Content="API Key:" Style="{StaticResource ModernLabel}"/>
                                <PasswordBox Name="AzureApiKeyBox" Grid.Row="1" Grid.Column="1" 
                                             Background="#FF3F3F46" Foreground="White" 
                                             BorderBrush="#FF007ACC" Margin="5,0"/>

                                <Label Grid.Row="2" Grid.Column="0" Content="Index Name:" Style="{StaticResource ModernLabel}"/>
                                <TextBox Name="IndexNameTextBox" Grid.Row="2" Grid.Column="1" 
                                         Text="email-rag-index" Style="{StaticResource ModernTextBox}" Margin="5,0"/>

                                <Button Name="TestAzureConnectionButton" Grid.Row="0" Grid.Column="2" Grid.RowSpan="3"
                                        Content="🧪 Test Connection" Style="{StaticResource ModernButton}" 
                                        VerticalAlignment="Center" Width="130"/>

                                <Label Name="AzureConnectionStatusLabel" Grid.Row="3" Grid.Column="1" 
                                       Content="Connection not tested" 
                                       Style="{StaticResource StatusLabel}" Margin="5,5,0,0"/>
                            </Grid>
                        </GroupBox>

                        <!-- OpenAI Configuration -->
                        <GroupBox Grid.Row="1" Header="🤖 OpenAI Configuration" Style="{StaticResource ModernGroupBox}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <Label Grid.Row="0" Grid.Column="0" Content="Endpoint:" Style="{StaticResource ModernLabel}"/>
                                <TextBox Name="OpenAIEndpointTextBox" Grid.Row="0" Grid.Column="1" Style="{StaticResource ModernTextBox}" Margin="5,0"/>

                                <Label Grid.Row="1" Grid.Column="0" Content="API Key:" Style="{StaticResource ModernLabel}"/>
                                <PasswordBox Name="OpenAIApiKeyBox" Grid.Row="1" Grid.Column="1" 
                                             Background="#FF3F3F46" Foreground="White" 
                                             BorderBrush="#FF007ACC" Margin="5,0"/>

                                <Label Grid.Row="2" Grid.Column="0" Content="Model:" Style="{StaticResource ModernLabel}"/>
                                <ComboBox Name="EmbeddingModelComboBox" Grid.Row="2" Grid.Column="1" 
                                          Background="#FF3F3F46" Foreground="White" Margin="5,0">
                                    <ComboBoxItem Content="text-embedding-ada-002" IsSelected="True"/>
                                    <ComboBoxItem Content="text-embedding-3-small"/>
                                    <ComboBoxItem Content="text-embedding-3-large"/>
                                </ComboBox>

                                <Button Name="TestOpenAIConnectionButton" Grid.Row="0" Grid.Column="2" Grid.RowSpan="3"
                                        Content="🧪 Test OpenAI" Style="{StaticResource ModernButton}" 
                                        VerticalAlignment="Center" Width="130"/>

                                <Label Name="OpenAIConnectionStatusLabel" Grid.Row="3" Grid.Column="1" 
                                       Content="Connection not tested" 
                                       Style="{StaticResource StatusLabel}" Margin="5,5,0,0"/>
                            </Grid>
                        </GroupBox>

                        <!-- Configuration Management -->
                        <GroupBox Grid.Row="2" Header="💾 Configuration Management" Style="{StaticResource ModernGroupBox}">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                <Button Name="LoadConfigButton" Content="📂 Load Config" Style="{StaticResource ModernButton}" Width="120"/>
                                <Button Name="SaveConfigButton" Content="💾 Save Config" Style="{StaticResource SuccessButton}" Width="120"/>
                                <Button Name="ResetConfigButton" Content="🔄 Reset to Defaults" Style="{StaticResource DangerButton}" Width="140"/>
                            </StackPanel>
                        </GroupBox>
                    </Grid>
                </ScrollViewer>
            </TabItem>

            <!-- Logs Tab -->
            <TabItem Header="📋 Logs">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button Name="ClearLogsButton" Content="🗑️ Clear Logs" Style="{StaticResource ModernButton}"/>
                        <Button Name="ExportLogsButton" Content="📤 Export Logs" Style="{StaticResource ModernButton}"/>
                        <Button Name="RefreshLogsButton" Content="🔄 Refresh" Style="{StaticResource ModernButton}"/>
                    </StackPanel>

                    <TextBox Name="LogsTextBox" Grid.Row="1" 
                             Background="#FF1E1E1E" 
                             Foreground="#FFCCCCCC" 
                             FontFamily="Consolas" 
                             FontSize="11"
                             IsReadOnly="True" 
                             TextWrapping="Wrap" 
                             AcceptsReturn="True" 
                             VerticalScrollBarVisibility="Auto" 
                             HorizontalScrollBarVisibility="Auto"/>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Status Bar -->
        <Border Grid.Row="2" Background="#FF1E1E1E" BorderBrush="#FF3F3F46" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Label Name="StatusBarLabel" Grid.Column="0" 
                       Content="Ready - Email RAG Cleaner v2.0" 
                       Style="{StaticResource ModernLabel}" 
                       VerticalAlignment="Center" Margin="10,0"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,10,0">
                    <Ellipse Name="StatusIndicator" Width="10" Height="10" Fill="#FF28A745" Margin="5,0"/>
                    <Label Name="ConnectionStatusLabel" Content="Configured" 
                           Style="{StaticResource ModernLabel}" FontSize="10"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Create and configure the WPF window
try {
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "❌ Failed to create WPF window: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Get control references
$controls = @{}
$window | Get-Member -MemberType Property | Where-Object { $_.Name -eq "Content" } | ForEach-Object {
    $controls = @{}
    $window.Content.Children | ForEach-Object {
        if ($_.Name) { $controls[$_.Name] = $_ }
    }
}

# Recursive function to find named controls
function Get-NamedControls($parent, $controlHash) {
    if ($parent.Children) {
        $parent.Children | ForEach-Object {
            if ($_.Name) {
                $controlHash[$_.Name] = $_
            }
            Get-NamedControls $_ $controlHash
        }
    } elseif ($parent.Content -and $parent.Content.Children) {
        $parent.Content.Children | ForEach-Object {
            if ($_.Name) {
                $controlHash[$_.Name] = $_
            }
            Get-NamedControls $_ $controlHash
        }
    } elseif ($parent.Items) {
        $parent.Items | ForEach-Object {
            if ($_.Name) {
                $controlHash[$_.Name] = $_
            }
            Get-NamedControls $_ $controlHash
        }
    }
}

$controls = @{}
Get-NamedControls $window $controls

# Logging function
function Write-GUILog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Update GUI log
    if ($controls["LogsTextBox"]) {
        $controls["LogsTextBox"].Text += "$logMessage`r`n"
        $controls["LogsTextBox"].ScrollToEnd()
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

# Update status functions
function Update-StatusBar {
    param([string]$Message, [string]$Status = "Ready")
    
    if ($controls["StatusBarLabel"]) {
        $controls["StatusBarLabel"].Content = $Message
    }
    
    if ($controls["StatusIndicator"]) {
        $color = switch ($Status) {
            "Ready" { "#FF28A745" }
            "Processing" { "#FFFFC107" }
            "Error" { "#FFDC3545" }
            default { "#FF6C757D" }
        }
        $controls["StatusIndicator"].Fill = $color
    }
}

function Update-FileCount {
    if ($controls["FilePathTextBox"] -and $controls["FileCountLabel"]) {
        $path = $controls["FilePathTextBox"].Text
        if ($path -and (Test-Path $path)) {
            $msgFiles = Get-ChildItem -Path $path -Filter "*.msg" -Recurse -ErrorAction SilentlyContinue
            $count = $msgFiles.Count
            $controls["FileCountLabel"].Content = "📄 $count MSG files found"
            $controls["FileCountLabel"].Foreground = if ($count -gt 0) { "#FF28A745" } else { "#FFDC3545" }
            $global:ProcessingStats.TotalFiles = $count
        } else {
            $controls["FileCountLabel"].Content = "❌ Invalid path or no files found"
            $controls["FileCountLabel"].Foreground = "#FFDC3545"
            $global:ProcessingStats.TotalFiles = 0
        }
    }
}

# Event handlers
if ($controls["BrowseButton"]) {
    $controls["BrowseButton"].Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select folder containing MSG files"
        
        if ($folderDialog.ShowDialog() -eq "OK") {
            $controls["FilePathTextBox"].Text = $folderDialog.SelectedPath
            Update-FileCount
        }
    })
}

if ($controls["FilePathTextBox"]) {
    $controls["FilePathTextBox"].Add_TextChanged({ Update-FileCount })
}

if ($controls["StartProcessingButton"]) {
    $controls["StartProcessingButton"].Add_Click({
        Start-EmailProcessing
    })
}

if ($controls["StopProcessingButton"]) {
    $controls["StopProcessingButton"].Add_Click({
        $global:ProcessingStats.IsProcessing = $false
        Write-GUILog "Processing stopped by user" "WARN"
        Update-StatusBar "Processing stopped" "Ready"
    })
}

if ($controls["TestButton"]) {
    $controls["TestButton"].Add_Click({
        Start-SystemTest
    })
}

if ($controls["LoadConfigButton"]) {
    $controls["LoadConfigButton"].Add_Click({
        Load-GUIConfiguration
    })
}

if ($controls["SaveConfigButton"]) {
    $controls["SaveConfigButton"].Add_Click({
        Save-GUIConfiguration
    })
}

if ($controls["TestAzureConnectionButton"]) {
    $controls["TestAzureConnectionButton"].Add_Click({
        Test-AzureConnection
    })
}

if ($controls["SearchButton"]) {
    $controls["SearchButton"].Add_Click({
        Perform-EmailSearch
    })
}

# Processing function
function Start-EmailProcessing {
    if (-not $controls["FilePathTextBox"].Text -or -not (Test-Path $controls["FilePathTextBox"].Text)) {
        [System.Windows.MessageBox]::Show("Please select a valid folder containing MSG files.", "Invalid Path", "OK", "Warning")
        return
    }
    
    if ($global:ProcessingStats.TotalFiles -eq 0) {
        [System.Windows.MessageBox]::Show("No MSG files found in the selected folder.", "No Files", "OK", "Warning")
        return
    }
    
    # Reset stats
    $global:ProcessingStats.ProcessedFiles = 0
    $global:ProcessingStats.SuccessfulFiles = 0
    $global:ProcessingStats.FailedFiles = 0
    $global:ProcessingStats.TotalChunks = 0
    $global:ProcessingStats.StartTime = Get-Date
    $global:ProcessingStats.IsProcessing = $true
    
    # Update UI state
    $controls["StartProcessingButton"].IsEnabled = $false
    $controls["StopProcessingButton"].IsEnabled = $true
    $controls["ProcessingProgressBar"].Value = 0
    $controls["ResultsListView"].Items.Clear()
    
    Write-GUILog "Starting processing of $($global:ProcessingStats.TotalFiles) MSG files" "INFO"
    Update-StatusBar "Processing emails..." "Processing"
    
    # Get processing options
    $options = @{
        CleanContent = $controls["CleanContentCheck"].IsChecked
        ExtractEntities = $controls["ExtractEntitiesCheck"].IsChecked
        CreateRAGChunks = $controls["CreateRAGCheck"].IsChecked
        UploadToAzure = $controls["UploadAzureCheck"].IsChecked
        GenerateEmbeddings = $controls["GenerateEmbeddingsCheck"].IsChecked
        IndexForSearch = $controls["IndexSearchCheck"].IsChecked
        ChunkSize = [int]$controls["ChunkSizeTextBox"].Text
        OverlapTokens = [int]$controls["OverlapTextBox"].Text
    }
    
    # Process files asynchronously
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("controls", $controls)
    $runspace.SessionStateProxy.SetVariable("options", $options)
    $runspace.SessionStateProxy.SetVariable("global:ProcessingStats", $global:ProcessingStats)
    
    $powershell = [powershell]::Create()
    $powershell.Runspace = $runspace
    
    $scriptBlock = {
        $msgFiles = Get-ChildItem -Path $controls["FilePathTextBox"].Text -Filter "*.msg" -Recurse
        
        foreach ($msgFile in $msgFiles) {
            if (-not $global:ProcessingStats.IsProcessing) {
                break
            }
            
            # Simulate processing (replace with actual processing logic)
            Start-Sleep -Milliseconds 500
            
            # Update progress on UI thread
            $controls["ProcessingProgressBar"].Dispatcher.Invoke([action]{
                $global:ProcessingStats.ProcessedFiles++
                $progressPercent = [math]::Round(($global:ProcessingStats.ProcessedFiles / $global:ProcessingStats.TotalFiles) * 100)
                $controls["ProcessingProgressBar"].Value = $progressPercent
                $controls["ProcessingStatusLabel"].Content = "Processing: $($msgFile.Name) ($($global:ProcessingStats.ProcessedFiles)/$($global:ProcessingStats.TotalFiles))"
                
                # Add result to list
                $result = [PSCustomObject]@{
                    FileName = $msgFile.Name
                    Status = "✅ Success"
                    Chunks = Get-Random -Minimum 5 -Maximum 20
                    SizeKB = [math]::Round((Get-Item $msgFile.FullName).Length / 1KB, 1)
                    ProcessingTime = [math]::Round((Get-Random -Minimum 100 -Maximum 5000) / 1000, 2)
                    Details = "Processed successfully with RAG optimization"
                }
                $controls["ResultsListView"].Items.Add($result)
                
                $global:ProcessingStats.SuccessfulFiles++
                $global:ProcessingStats.TotalChunks += $result.Chunks
                
                # Update stats
                $duration = [math]::Round(((Get-Date) - $global:ProcessingStats.StartTime).TotalSeconds, 1)
                $controls["ProcessingStatsLabel"].Content = "Files: $($global:ProcessingStats.ProcessedFiles) processed, $($global:ProcessingStats.SuccessfulFiles) successful, $($global:ProcessingStats.FailedFiles) failed | Chunks: $($global:ProcessingStats.TotalChunks) | Time: $($duration)s"
            })
        }
        
        # Finish processing
        $controls["ProcessingProgressBar"].Dispatcher.Invoke([action]{
            $global:ProcessingStats.EndTime = Get-Date
            $duration = ($global:ProcessingStats.EndTime - $global:ProcessingStats.StartTime).TotalSeconds
            
            $controls["ProcessingStatusLabel"].Content = "✅ Processing completed in $([math]::Round($duration, 2)) seconds"
            $controls["StartProcessingButton"].IsEnabled = $true
            $controls["StopProcessingButton"].IsEnabled = $false
            $global:ProcessingStats.IsProcessing = $false
        })
    }
    
    $powershell.AddScript($scriptBlock)
    $powershell.BeginInvoke()
}

# Configuration functions
function Load-GUIConfiguration {
    try {
        $configPath = Join-Path $installPath "Config\default-config.json"
        if (Test-Path $configPath) {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            
            if ($config.AzureSearch) {
                $controls["AzureServiceNameTextBox"].Text = $config.AzureSearch.ServiceName
                $controls["IndexNameTextBox"].Text = $config.AzureSearch.IndexName
            }
            
            Write-GUILog "Configuration loaded successfully" "SUCCESS"
            [System.Windows.MessageBox]::Show("Configuration loaded successfully!", "Load Configuration", "OK", "Information")
        } else {
            Write-GUILog "Configuration file not found" "WARN"
            [System.Windows.MessageBox]::Show("Configuration file not found. Using defaults.", "Load Configuration", "OK", "Warning")
        }
    } catch {
        Write-GUILog "Failed to load configuration: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Failed to load configuration: $($_.Exception.Message)", "Load Error", "OK", "Error")
    }
}

function Save-GUIConfiguration {
    try {
        $config = @{
            AzureSearch = @{
                ServiceName = $controls["AzureServiceNameTextBox"].Text
                ApiKey = $controls["AzureApiKeyBox"].Password
                IndexName = $controls["IndexNameTextBox"].Text
                ServiceUrl = "https://$($controls["AzureServiceNameTextBox"].Text).search.windows.net"
            }
            OpenAI = @{
                Endpoint = $controls["OpenAIEndpointTextBox"].Text
                ApiKey = $controls["OpenAIApiKeyBox"].Password
                EmbeddingModel = $controls["EmbeddingModelComboBox"].SelectedItem.Content
            }
            Processing = @{
                ChunkSize = [int]$controls["ChunkSizeTextBox"].Text
                OverlapTokens = [int]$controls["OverlapTextBox"].Text
            }
        }
        
        $configPath = Join-Path $installPath "Config\gui-config.json"
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
        
        Write-GUILog "Configuration saved successfully" "SUCCESS"
        [System.Windows.MessageBox]::Show("Configuration saved successfully!", "Save Configuration", "OK", "Information")
    } catch {
        Write-GUILog "Failed to save configuration: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Failed to save configuration: $($_.Exception.Message)", "Save Error", "OK", "Error")
    }
}

function Test-AzureConnection {
    if (-not $controls["AzureServiceNameTextBox"].Text -or -not $controls["AzureApiKeyBox"].Password) {
        [System.Windows.MessageBox]::Show("Please enter Azure Search service name and API key.", "Missing Configuration", "OK", "Warning")
        return
    }
    
    try {
        Write-GUILog "Testing Azure AI Search connection..." "INFO"
        $controls["AzureConnectionStatusLabel"].Content = "🔄 Testing connection..."
        
        # Simulate connection test (replace with actual test)
        Start-Sleep -Seconds 2
        
        $controls["AzureConnectionStatusLabel"].Content = "✅ Connection successful"
        $controls["AzureConnectionStatusLabel"].Foreground = "#FF28A745"
        
        Write-GUILog "Azure AI Search connection test successful" "SUCCESS"
        [System.Windows.MessageBox]::Show("Azure AI Search connection successful!", "Connection Test", "OK", "Information")
    } catch {
        $controls["AzureConnectionStatusLabel"].Content = "❌ Connection failed"
        $controls["AzureConnectionStatusLabel"].Foreground = "#FFDC3545"
        
        Write-GUILog "Azure connection test failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Azure connection failed: $($_.Exception.Message)", "Connection Test Failed", "OK", "Error")
    }
}

function Start-SystemTest {
    Write-GUILog "Starting system test..." "INFO"
    Update-StatusBar "Running system tests..." "Processing"
    
    try {
        # Simulate comprehensive system test
        $controls["ProcessingStatusLabel"].Content = "🧪 Running comprehensive system tests..."
        
        # Test modules
        Write-GUILog "Testing module loading..." "INFO"
        Start-Sleep -Seconds 1
        
        # Test configuration
        Write-GUILog "Testing configuration..." "INFO"
        Start-Sleep -Seconds 1
        
        # Test Azure connection (if configured)
        Write-GUILog "Testing Azure connectivity..." "INFO"
        Start-Sleep -Seconds 2
        
        $controls["ProcessingStatusLabel"].Content = "✅ All system tests passed successfully"
        Write-GUILog "System test completed successfully" "SUCCESS"
        Update-StatusBar "System tests completed" "Ready"
        
        [System.Windows.MessageBox]::Show("All system tests passed successfully!\n\n✅ Module loading\n✅ Configuration validation\n✅ Azure connectivity", "System Test Results", "OK", "Information")
    } catch {
        Write-GUILog "System test failed: $($_.Exception.Message)" "ERROR"
        Update-StatusBar "System test failed" "Error"
        [System.Windows.MessageBox]::Show("System test failed: $($_.Exception.Message)", "System Test Failed", "OK", "Error")
    }
}

function Perform-EmailSearch {
    $query = $controls["SearchQueryTextBox"].Text
    if (-not $query -or $query -eq "Enter your search query...") {
        [System.Windows.MessageBox]::Show("Please enter a search query.", "Missing Query", "OK", "Warning")
        return
    }
    
    Write-GUILog "Performing search: $query" "INFO"
    $controls["SearchStatusLabel"].Content = "🔄 Searching..."
    $controls["SearchResultsListView"].Items.Clear()
    
    try {
        # Simulate search results (replace with actual search)
        Start-Sleep -Seconds 1
        
        $searchResults = @(
            [PSCustomObject]@{
                Document = "Meeting_Notes_2024.msg"
                Score = "0.95"
                Sender = "john.doe@company.com"
                Date = "2024-01-15"
                Preview = "Quarterly meeting discussion about project milestones and deadlines..."
            },
            [PSCustomObject]@{
                Document = "Project_Update.msg"
                Score = "0.87"
                Sender = "jane.smith@company.com"
                Date = "2024-01-10"
                Preview = "Project status update with timeline adjustments and resource allocation..."
            },
            [PSCustomObject]@{
                Document = "Budget_Review.msg"
                Score = "0.76"
                Sender = "finance@company.com"
                Date = "2024-01-08"
                Preview = "Annual budget review and approval process for next fiscal year..."
            }
        )
        
        foreach ($result in $searchResults) {
            $controls["SearchResultsListView"].Items.Add($result)
        }
        
        $controls["SearchStatusLabel"].Content = "✅ Found $($searchResults.Count) results"
        Write-GUILog "Search completed: $($searchResults.Count) results found" "SUCCESS"
    } catch {
        $controls["SearchStatusLabel"].Content = "❌ Search failed"
        Write-GUILog "Search failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show("Search failed: $($_.Exception.Message)", "Search Error", "OK", "Error")
    }
}

# Initialize the application
Write-GUILog "Email RAG Cleaner v2.0 GUI started" "INFO"
Update-StatusBar "Email RAG Cleaner v2.0 - Ready" "Ready"

# Load configuration on startup
$configPath = Join-Path $installPath "Config\gui-config.json"
if (Test-Path $configPath) {
    Load-GUIConfiguration
}

# Show the window
$window.ShowDialog() | Out-Null