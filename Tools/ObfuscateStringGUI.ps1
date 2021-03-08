
<#
Solution: Microsoft Deployment Toolkit
 Purpose: Obfuscate String GUI
 Version: 1.0.0
    Date: 08 Mars 2021

  Author: Tomas Johansson
 Twitter: @deploymentnoob
     Web: https://www.4thcorner.net

This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the author
#>

#Add WPF and Windows Forms assemblies
try {
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName system.windows.forms
}
catch {
    Throw "Failed to load Windows Presentation Framework assemblies."
}

# Export Dialog
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Obfuscate Text" Height="490" Width="500" Topmost="True" ResizeMode="NoResize">
    <Grid>
        <Label Name="Labeltext" Content="" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top"/>
        <TextBox Name="TextBox_Clear" HorizontalAlignment="Left" Margin="20,40,0,0" Text="Text to obfuscate" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="24" Opacity="1"/>
        <Label Name="LabelKey" Content="Key to use for obfuscate text (Leave Empty to create New Key):" HorizontalAlignment="Left" Margin="20,70,0,0" VerticalAlignment="Top"/>
        <TextBox Name="TextBox_key" HorizontalAlignment="Left" Margin="20,95,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="24"/>
        <GroupBox Name="GroupBox_keylength" Header="Key Length" Margin="20,128,300,240">
            <Grid Width="170" Height="45" Margin="0,0,0,0">
                <RadioButton Name="RadioButton16" Content="16" HorizontalAlignment="Left" Margin="15,0,0,0" VerticalAlignment="Center"/>
                <RadioButton Name="RadioButton24" Content="24" HorizontalAlignment="Left" Margin="65,0,0,0" VerticalAlignment="Center"/>
                <RadioButton Name="RadioButton32" Content="32" HorizontalAlignment="Left" Margin="115,0,0,0" VerticalAlignment="Center"/>
            </Grid>
        </GroupBox>
        <TextBox Name="Textbox_Obfuscated" HorizontalAlignment="Left" Margin="20,217,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="181"/>
        <Button Name="Button_obfuscate" Content="Obfuscate" HorizontalAlignment="Left" Margin="300,420,0,0" VerticalAlignment="Top" Width="75" Height="20"/>
        <Button Name="Button_cancel" Content="Cancel" HorizontalAlignment="Left" Margin="395,420,0,0" VerticalAlignment="Top" Width="75" Height="20"/>
    </Grid>
</Window>
'@

# Create the XAML reader using a new XML node reader
$XAMLReader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($XAMLReader)
foreach ($Name in ($XAML | Select-Xml '//*/@Name' | ForEach-Object { $_.Node.Value})) {
    New-Variable -Name $Name -Value $Window.FindName($Name) -Force
}

# Textbox Key have key disable groupbox for selecting byte length
$textBox_key.Add_TextChanged({
    If (!([string]::IsNullOrEmpty($textBox_key.Text))) {
        $groupBox_keylength.IsEnabled = $false
    }
    Else {
        $groupBox_keylength.IsEnabled = $true
    }
})

# Close dialog when pressing Escape key
$objForm = New-Object System.Windows.Forms.Form 
$objForm.KeyPreview = $True

$objForm.Add_KeyDown({
    If ($_.KeyCode -eq "Escape") {
        $objForm.Close()
    }
})

# Close dialog window
$button_cancel.add_click({
    $Window.Close() 
})

# Launch the window
$Window.WindowStartupLocation="CenterScreen"
$Window.Add_Loaded( {
    $this.TopMost = $true
})
# Show the window
$Window.ShowDialog() | out-null