
<#
Solution: Microsoft Deployment Toolkit
 Purpose: Encrypt Text GUI
 Version: 1.0.0
    Date: 08 March 2021

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
        Title="Encrypt Text" Height="475" Width="500" Topmost="True" ResizeMode="NoResize">
    <Grid>
        <Label x:Name="Labeltext" Content="Text to Encrypt:" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top" Width="198"/>
        <TextBox Name="TextBox_Clear" HorizontalAlignment="Left" Margin="20,40,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="24" Opacity="1"/>
        <Label Name="LabelKey" Content="Key to use for obfuscate text (Leave Empty to create new Key):" HorizontalAlignment="Left" Margin="20,70,0,0" VerticalAlignment="Top"/>
        <TextBox Name="TextBox_key" HorizontalAlignment="Left" Margin="20,95,0,0" Text="" TextWrapping="NoWrap" VerticalAlignment="Top" Width="450" Height="24" />
        <GroupBox Name="GroupBox_keylength" Header="Key Length" Margin="20,130,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" VerticalContentAlignment="Center">
            <Grid Height="Auto" Margin="20,20,0,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="8*"/>
                    <RowDefinition Height="8*"/>
                </Grid.RowDefinitions>
                <RadioButton Name="RadioButton16" Content="16" HorizontalAlignment="Left" VerticalAlignment="Center" Padding="0" Margin="-19,-11,0,5"/>
                <RadioButton Name="RadioButton24" Content="24" HorizontalAlignment="Left" Margin="22,-10,0,5" VerticalAlignment="Center" Padding="0"/>
                <RadioButton Name="RadioButton32" Content="32" HorizontalAlignment="Left" Margin="65,-10,0,5" VerticalAlignment="Center" Padding="0,0,5,0"/>
            </Grid>
        </GroupBox>
        <TextBox Name="Textbox_Obfuscated" HorizontalAlignment="Left" Margin="20,197,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="181"/>
        <Button Name="Button_obfuscate" Content="Encrypt" HorizontalAlignment="Left" Margin="300,395,0,0" VerticalAlignment="Top" Width="75" Height="20"/>
        <Button Name="Button_cancel" Content="Cancel" HorizontalAlignment="Left" Margin="395,395,0,0" VerticalAlignment="Top" Width="75" Height="20"/>
    </Grid>
</Window>
'@

# Create Hashtable and Runspace for GUI
$syncHash = [hashtable]::Synchronized(@{})
$NewRunspace =[runspacefactory]::CreateRunspace()
$NewRunspace.ApartmentState = "STA"
$NewRunspace.ThreadOptions = "ReuseThread"         
$NewRunspace.Open()
$NewRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)      


# Create the XAML reader using a new XML node reader
$XAMLReader = New-Object System.Xml.XmlNodeReader $XAML
$syncHash.Window = [Windows.Markup.XamlReader]::Load($XAMLReader)
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | 
ForEach-Object {
    $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name))
}

Function New-ByteKey {
<#
.SYNOPSIS
    Create a Byte Array
.DESCRIPTION
    Create an New Byte Array
.PARAMETER Lenght
    Set key Lengt, Must be 16,24 or 32 (Maximal is 32)
.INPUTS
    Lenght : Interger
.OUTPUTS
    Key    : String
.EXAMPLE
    New-EncryptedString -PlainTextString "I am about to become encrypted!" -Key $Key
#>
    Param (
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeLine=$true)]
        [ValidateSet("16","24","32")]
        [Alias("KeyLength")]
        [INT]$Length
    )
    BEGIN  {
        
    }
    PROCESS {
        $KeyByte = New-Object Byte[] $Length # You can use 16, 24, or 32 for AES
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($KeyByte)
        $Key = $KeyByte -join ","
    }
    END {
        Return $Key
    }
}

Function New-EncryptedString {
<#
.SYNOPSIS
    Encrypt String with Secret Key
.DESCRIPTION
    Create an New String from input
.PARAMETER PlainTextString
    String to Encrypt
.PARAMETER Key
    Byte array
.INPUTS
    PlainTextString  : String
    Key              : Byte array
.OUTPUTS
    EncryptedString  : String
.EXAMPLE
    New-EncryptedString -PlainTextString "I am about to become encrypted!" -Key $Key
#>
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeLine=$true)]
        [Alias("String")]
        [String]$PlainTextString,
        [Parameter(Mandatory=$True,Position=1)]
        [Alias("Key")]
        [byte[]]$EncryptionKey
    )
    BEGIN {
    }
    PROCESS { 
        Try{
            $secureString = Convertto-SecureString $PlainTextString -AsPlainText -Force
            $EncryptedString = ConvertFrom-SecureString -SecureString $secureString -Key $EncryptionKey
        }
        Catch{
            $EncryptedString  = $null
        }
    }
    END {
        Return $EncryptedString
    }
}

# Click Encrypt buttton
$SyncHash.Button_obfuscate.add_Click({
    If ($SyncHash.RadioButton16.IsChecked) {
        $KeyLength = 16
    }
    ElseIf ($SyncHash.RadioButton24.IsChecked){
        $KeyLength = 24
    }
    ElseIf ($SyncHash.RadioButton32.IsChecked){
        $KeyLength = 32
    }
    Else {
        $KeyLength = $null
    }
    If (([string]::IsNullOrEmpty($syncHash.TextBox_key.Text))) {
        If ([string]::IsNullOrEmpty($KeyLength)) {
            [System.Windows.Forms.MessageBox]::Show("Key Length not Selected" , "Error")
        }
        Else {
            $NewKey = New-ByteKey $KeyLength
            $Key = $NewKey.Split(",")
        }
    }
    Else {
        $Key = ($syncHash.TextBox_key.Text).Split(",")

    }

    $TextString = $syncHash.TextBox_Clear.Text
    
    $Retval = New-EncryptedString -PlainTextString $TextString -EncryptionKey $Key

    $syncHash.Window.Dispatcher.Invoke(
        [action]{
        $SyncHash.RadioButton16.IsChecked = $false
        $SyncHash.RadioButton24.IsChecked = $false
        $SyncHash.RadioButton32.IsChecked = $false
        $SyncHash.Textbox_Obfuscated.Text = $Retval
        $SyncHash.TextBox_key.Text = $Key -join ","
        }
    )


})

# Textbox Key have key disable groupbox for selecting byte length
$syncHash.textBox_key.Add_TextChanged({
    If (!([string]::IsNullOrEmpty($syncHash.TextBox_key.Text))) {
        $syncHash.GroupBox_keylength.IsEnabled = $false
        $KeyLength = $null
    }
    Else {
        $syncHash.GroupBox_keylength.IsEnabled = $true
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
$syncHash.button_cancel.add_click({
    $syncHash.Window.Close() 
})

# Launch the window
$syncHash.Window.WindowStartupLocation="CenterScreen"
$syncHash.Window.Add_Loaded( {
    $this.TopMost = $true
})
# Show the window
$syncHash.Window.ShowDialog() | out-null