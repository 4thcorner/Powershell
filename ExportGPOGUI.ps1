<#
Solution: Hydration System Center 2012 R2 
 Purpose: Export selected GPO´s with GUI Dialog
 Version: 1.0
    Date: 25 May 2016

  Author: Tomas Johansson
 Twitter: @CanKlant
    Blog: http://www.4thcorner.se

 This script is provided "AS IS" with no warranties, confers no rights and 
 is not supported by the author
#>
#Add WPF and Windows Forms assemblies
try {
    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
}
catch {
    Throw "Failed to load Windows Presentation Framework assemblies."
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
$GPOtoExport = @()
$GPOList =  @()

# Get Doamin
$Domain = (Get-ADDomain).DNSRoot

# Get all GPO´s in domain
$GPOList = (Get-GPO -All -Domain $Domain).Displayname | Sort

# Get ScriptPath
$Path = split-path -parent $MyInvocation.MyCommand.Path

#Get Initial Directory for browser button
$InitialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

# Set Deafult export path
$ExportPath = $Path

# Show an Open Folder Dialog and return the directory selected by the user.
function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton)
{
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
 
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
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
            Backup-GPO -Name $GPOName -Path $GPONamePath -ErrorAction Ignore | Out-Null
            #Write-Host -ForegroundColor Green "GPO: $GPOName exist in domain and backup is done to $GPONamePath"
        }
        Catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting" , "Errors")
            Exit
        }
    }
}

#EVENT Handler

# GPO Listbox content
$GPOlistBox.ItemsSource = $GPOList

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
    $Global:GPOToExport = $GPOlistBox.SelectedItems
    if ([string]::IsNullOrEmpty($GPOtoExport)) { 
        [System.Windows.Forms.MessageBox]::Show("No Group Policys Selected!" , "Group Policys")
    }
    Else {
        ExportSelectedGPO -ExportPath $ExportPath -GPOtoExport $GPOToExport
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

#Launch the window
$Window.WindowStartupLocation="CenterScreen"
$Window.Add_Loaded( {
    $this.TopMost = $true
})
$Window.ShowDialog() | Out-Null