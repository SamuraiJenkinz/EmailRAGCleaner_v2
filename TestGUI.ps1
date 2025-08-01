# Email RAG Cleaner v2.0 - Simple Test GUI
# This version works without modules for testing

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Simple XAML without class name
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email RAG Cleaner v2.0 - Test GUI" 
        Height="400" Width="600" 
        WindowStartupLocation="CenterScreen"
        Background="#FF2D2D30">
    
    <Grid>
        <StackPanel Margin="20">
            <TextBlock Text="Email RAG Cleaner v2.0" 
                       FontSize="24" FontWeight="Bold" 
                       Foreground="#FF007ACC" 
                       HorizontalAlignment="Center" 
                       Margin="0,0,0,20"/>
            
            <TextBlock Text="GUI Test - Modules Not Required" 
                       FontSize="14" 
                       Foreground="White" 
                       HorizontalAlignment="Center" 
                       Margin="0,0,0,30"/>
            
            <GroupBox Header="System Status" Foreground="White" BorderBrush="#FF007ACC">
                <StackPanel Margin="10">
                    <TextBlock Name="StatusText" 
                               Text="Checking system..." 
                               Foreground="White" 
                               Margin="0,0,0,10"/>
                    <ProgressBar Name="TestProgress" Height="20" Margin="0,0,0,10"/>
                    <Button Name="TestButton" 
                            Content="Run System Test" 
                            Background="#FF007ACC" 
                            Foreground="White" 
                            Padding="10,5"
                            HorizontalAlignment="Center"/>
                </StackPanel>
            </GroupBox>
            
            <TextBlock Name="InfoText" 
                       Text="This is a test GUI to verify WPF is working correctly." 
                       Foreground="#FFCCCCCC" 
                       TextWrapping="Wrap"
                       Margin="0,20,0,0"/>
            
            <Button Name="CloseButton" 
                    Content="Close" 
                    Background="#FFDC3545" 
                    Foreground="White" 
                    Padding="20,5"
                    HorizontalAlignment="Center"
                    Margin="0,20,0,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

try {
    # Parse XAML
    [xml]$xamlXml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xamlXml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Get controls
    $statusText = $window.FindName("StatusText")
    $testProgress = $window.FindName("TestProgress")
    $testButton = $window.FindName("TestButton")
    $closeButton = $window.FindName("CloseButton")
    $infoText = $window.FindName("InfoText")
    
    # Button events
    $testButton.Add_Click({
        $statusText.Text = "Running tests..."
        $testProgress.IsIndeterminate = $true
        
        # Check installation
        $installPath = "C:\EmailRAGCleaner"
        if (Test-Path $installPath) {
            $statusText.Text = "✅ Installation found at: $installPath"
            
            # Check for modules
            $modulesPath = Join-Path $installPath "Modules"
            if (Test-Path $modulesPath) {
                $modules = Get-ChildItem -Path $modulesPath -Filter "*.psm1"
                $infoText.Text = "Found $($modules.Count) modules in the Modules directory.`n`nModules: $($modules.Name -join ', ')"
            } else {
                $infoText.Text = "⚠️ Modules directory not found. Run CopyModules.bat to copy the modules."
            }
        } else {
            $statusText.Text = "❌ Installation not found at: $installPath"
            $infoText.Text = "Please run the installer first."
        }
        
        $testProgress.IsIndeterminate = $false
        $testProgress.Value = 100
    })
    
    $closeButton.Add_Click({
        $window.Close()
    })
    
    # Initial status
    $statusText.Text = "Ready - Click 'Run System Test' to check installation"
    
    # Show window
    $window.ShowDialog() | Out-Null
    
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to create test GUI: $_",
        "Error",
        "OK",
        "Error"
    )
}