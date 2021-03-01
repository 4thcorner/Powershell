<#
Solution: Microsoft Deployment Toolkit
 Purpose: Group Policys Export GUI
 Version: 2.0.0
    Date: 20 July 2020

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

# Check for elevation
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    [System.Windows.MessageBox]::Show("Oupps!`nYou need to run this Script from an Elevated PowerShell!`nPlease start the PowerShell as an Administrator`nAborting script...",'Group Policys Export','OK',[System.Windows.Forms.MessageBoxIcon]::Warning)
    Break
}

# Export Dialog
[xml]$XAML = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Export Group Policy" Height="380" MinHeight="380" Width="400" MinWidth="400" MaxWidth="400" >
    <Grid Name="Grid" Margin="20,10,10,10" HorizontalAlignment="Left">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label Name="SelectionInfolabel" Content="Please make a selection from the list below:" Grid.Row="1" Grid.ColumnSpan="3"/>
        <ListBox Name="GPOlistBox" Height="200" Width="Auto" MinWidth="340" Grid.Row="2" Grid.ColumnSpan="3" SelectionMode="Multiple"/>
        <Label Name="Infolabel" Content="Exporting to:" Grid.Row="3" Grid.ColumnSpan="3" />
        <TextBox Name="ExporttextBox" Height="24" TextWrapping="NoWrap" Text="TextBox" VerticalContentAlignment="Center" Width="250" Grid.Row="4" Grid.ColumnSpan="3" Margin="0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
        <Button Name="BrowseButton" Content="Browse" Width="76" Height="24" Grid.Row="4" Grid.Column="2" Margin="0,4,0,6" VerticalAlignment="Center" HorizontalAlignment="Right"/>
        <Button Name="OKButton" Content="Export" Width="75" Margin="0,10,0,0" Grid.Row="5" Grid.Column="2" Height="24" HorizontalAlignment="Left"/>
        <Button Name="CancelButton" Content="Cancle" Width="75" Margin="0,10,0,0" Grid.Row="5" Grid.Column="3" HorizontalAlignment="Right" Height="24"/>
    </Grid>
</Window>
'@

#Create the XAML reader using a new XML node reader
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)
foreach ($Name in ($XAML | Select-Xml '//*/@Name' | foreach { $_.Node.Value})) {
    New-Variable -Name $Name -Value $Window.FindName($Name) -Force
}

# Constants
#$GPOtoExport = @()
#$GPOList =  @()

# Get Doamin
$Domain = (Get-ADDomain).DNSRoot

# Get all GPO´s in domain
$GPOList = (Get-GPO -All -Domain $Domain).Displayname | Sort

# Get ScriptPath
$Path = split-path -parent $MyInvocation.MyCommand.Path

#Get Initial Directory for browser button
$InitialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

# Set Deafult export path
$Global:ExportPath = $Null

# Show an Open Folder Dialog and return the directory selected by the user.
Function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton) {
    $BrowseForFolderOptions = 0
    If ($NoNewFolderButton) {
        $BrowseForFolderOptions += 512
    }
    $App = New-Object -ComObject Shell.Application
    $Folder = $App.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    If ($Folder) {
        $selectedDirectory = $Folder.Self.Path
    }
    Else {
        $selectedDirectory = $Null
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($App) > $null
    Return $selectedDirectory
}

# Export Selected Group Policys to selected folder
Function ExportSelectedGPO ([string]$ExportPath,[Array]$GPOtoExport )
{   
    # Start Export of GPO´s
    Foreach ($GPOName in $GPOtoExport) {
    
        $GPONamePath = $ExportPath + '\' + $GPOName

        Try {
            # Test if Directorys exist
            If (!(Test-Path -PathType Container $GPONamePath))
            {
                New-Item -ItemType Directory -Force -Path $GPONamePath | Out-Null
            }
            Else {
                Remove-Item -Path $GPONamePath -Recurse  -Force
                New-Item -ItemType Directory -Force -Path $GPONamePath | Out-Null
            }
             
            # Backup pf GPO
            Backup-GPO -Name $GPOName -Path $GPONamePath -ErrorAction SilentlyContinue | Out-Null
        }
        Catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting" , "Errors")
            Exit
        }
    }
}

# GPO Listbox content
$GPOlistBox.ItemsSource = $GPOList

# Selectionmode for ListBox
$GPOlistBox.SelectionMode = 'Multiple'

# Textbox with Path
$ExporttextBox.Text = $Path

$BrowseButton.add_Click({
    $Window.Topmost = $false
    $Global:ExportPath = Read-FolderBrowserDialog -Message "Please select a directory" -InitialDirectory $InitialDirectory
    if ([string]::IsNullOrEmpty($ExportPath)) { 
        [System.Windows.Forms.MessageBox]::Show("No Export Directory Selected!" , "Export Directory Selected")
    }
    Else {
        $ExporttextBox.Clear()
        $ExporttextBox.AppendText($ExportPath)  
    }
    $Window.Topmost = $true
})

$OKButton.add_Click({
    
    $GPOtoExport = $GPOlistBox.SelectedItems

    If ([string]::IsNullOrEmpty($GPOtoExport)) { 
        [System.Windows.Forms.MessageBox]::Show("No Group Policys Selected!" , "Group Policys")
    }
    Else {
        ExportSelectedGPO -ExportPath $ExportPath -GPOtoExport $GPOToExport
        $Window.Close()
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = “Group Policys Exported”
        $Messageboxbody = “Selected Group Policys exported"
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        $Window.Topmost = $false
        [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
        $Window.Close()
    }
})

# Close dialog window
$CancelButton.add_click({
    $Window.Close() 
}) 

# Launch the window
$Window.WindowStartupLocation="CenterScreen"
$Window.Add_Loaded({
    $this.TopMost = $true
})
$Window.ShowDialog() | Out-Null