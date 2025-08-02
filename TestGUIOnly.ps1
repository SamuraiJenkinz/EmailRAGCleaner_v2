# Simple GUI test script without module loading
Add-Type -AssemblyName PresentationFramework

try {
    Write-Host "Testing GUI without modules..." -ForegroundColor Cyan
    
    # XAML for simple test window
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Test Window" Height="300" Width="400">
    <StackPanel Margin="20">
        <GroupBox Header="File Selection" Margin="0,0,0,20">
            <StackPanel>
                <TextBox Name="PathTextBox" Margin="5"/>
                <Button Name="TestButton" Content="Test Button" Margin="5"/>
            </StackPanel>
        </GroupBox>
        <TextBlock Text="If you see this, the GUI works!" Margin="10"/>
    </StackPanel>
</Window>
"@

    [xml]$xamlXml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xamlXml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    Write-Host "GUI created successfully!" -ForegroundColor Green
    Write-Host "Close the window to continue..." -ForegroundColor Yellow
    
    # Show the window
    $null = $window.ShowDialog()
    
    Write-Host "GUI test completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "GUI test failed: $_" -ForegroundColor Red
}