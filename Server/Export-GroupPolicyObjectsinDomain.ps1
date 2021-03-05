<#
Solution: Microsoft Deployment Toolkit
 Purpose: Group Policys Export GUI
 Version: 3.0.0
    Date: 05 Mars 2021

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
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
     [System.Windows.MessageBox]::Show("Oupps!`nYou need to run this Script from an Elevated PowerShell!`nPlease start the PowerShell as an Administrator`nAborting script...",'Group Policys Export','OK',[System.Windows.Forms.MessageBoxIcon]::Warning)
     Break
}

# Export Dialog
[xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Export Group Policys in Domain" Height="410" Width="930" Topmost="True" ResizeMode="NoResize">
    <Grid Name="grid" Margin="20,10,20,10">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto" />
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Label Name="Information" Content="Group Policys i Domain:" Grid.Row="1" Margin="0,0,5,280" Grid.RowSpan="2" Grid.Column="1" />
    <Label Name="Domain" Content="Domain" Grid.Row="1" Margin="130,0,5,280" Grid.RowSpan="2" />
    <ListView Name="ListViewGPO" HorizontalAlignment="Center" Height="270" Width="880" VerticalAlignment="Center"  Grid.Row="2">
        <ListView.View>
            <GridView>
                <GridViewColumn Header="Group Policys" DisplayMemberBinding="{Binding GPOName}" />
                <GridViewColumn Header="Group Policys ID" DisplayMemberBinding="{Binding GPOId}" />
                <GridViewColumn Header="Group Policys Status" DisplayMemberBinding="{Binding GPOStatus}" />
                <GridViewColumn Header="WMI Filter" DisplayMemberBinding="{Binding WMIFilter}" />
            </GridView>
        </ListView.View>
    </ListView>
    <TextBox Name="ExportPathBox" HorizontalAlignment="Left" Height="24" Width="570" Grid.Row="3" Text="Browse for Export Folder" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,0,0,0" Grid.RowSpan="2"/>
    <Button Name="Browsebutton" Content="Browse" HorizontalAlignment="Left" Margin="580,0,0,0" Grid.Row="3" VerticalAlignment="Center" Width="76" Grid.RowSpan="2"/>
    <Button Name="Exportbutton" Content="Export" HorizontalAlignment="Left" Margin="696,0,0,0" Grid.Row="3" VerticalAlignment="Center" Width="75" Grid.RowSpan="2"/>
    <Button Name="Cancelbutton" Content="Cancel" HorizontalAlignment="Left" Margin="786,0,0,0" Grid.Row="3" VerticalAlignment="Center" Width="74" Grid.RowSpan="2"/>
    </Grid>
</Window>
'@


#Create the XAML reader using a new XML node reader
$XAMLReader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($XAMLReader)
foreach ($Name in ($XAML | Select-Xml '//*/@Name' | ForEach-Object { $_.Node.Value})) {
    New-Variable -Name $Name -Value $Window.FindName($Name) -Force
}
# Constants
$Global:ExportPath = $Null

# Get Initial Directory for browser button
$InitialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

# Sort event handler
$Window.Add_SourceInitialized({
    [System.Windows.RoutedEventHandler]$ColumnSortHandler = {

        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
           
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {

                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                                                
                # And now we actually apply the sort to the View
                $ListViewGPO_DefaultView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ListViewGPO.ItemsSource)
                # Change the sort direction each time they sort
                    Switch($LListViewGPOn_DefaultView.SortDescriptions[0].Direction)
                    {
                        "Decending" { $Direction = "Ascending" }
                        "Ascending" { $Direction = "Descending" }
                            Default { $Direction = "Ascending" }
                    }           
                $ListViewGPO_DefaultView.SortDescriptions.Clear()
                $ListViewGPOn_DefaultView.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $Direction))
                $ListViewGPO_DefaultView.Refresh()  
            }
        }
    }
    #Attach the Event Handler
    $ListViewGPO.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ColumnSortHandler)
})

Function Import-ListView {
    $GPOInfos = @()
    $Domain = (Get-ADDomain).DNSRoot
    $GPOLIST = Get-GPO -All -Domain $Domain | Select-Object DisplayName, Id, GpoStatus, WmiFilter | Sort-Object DisplayName
    
    foreach ($GPO in $GPOLIST) {
        $GPOInfo = New-Object -TypeName PSObject
        $GPOInfo | Add-Member -MemberType NoteProperty -Name 'GPOName' -Value $GPO.DisplayName
        $GPOInfo | Add-Member -MemberType NoteProperty -Name 'GPOId' -Value $GPO.Id
        $GPOInfo | Add-Member -MemberType NoteProperty -Name 'GPOStatus' -Value $GPO.GpoStatus
        $GPOInfo | Add-Member -MemberType NoteProperty -Name 'WMIFilter' -Value $GPO.WmiFilter
        $GPOInfos += $GPOInfo
    }
    $ListViewGPO.ItemsSource = $GPOInfos
}

# Show an Open Folder Dialog and return the directory selected by the user
Function Read-FolderBrowserDialog {
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$true)]
        [string]$InitialDirectory,
        [Parameter(Mandatory=$false)]
        [switch]$NoNewFolderButton
    )

    $BrowseForFolderOptions = 0
    If ($NoNewFolderButton) {
        $BrowseForFolderOptions += 512
    }
    $App = New-Object -ComObject Shell.Application
    $Folder = $App.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    If ($Folder) {
        $SelectedDirectory = $Folder.Self.Path
    }
    Else {
        $selectedDirectory = $Null
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($App) > $null
    Return $SelectedDirectory
}

# Export GPO to Selected Folder
Function Export-SelectedGPO () {
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [String]$ExportPath,
        [Parameter(Mandatory=$true)]
        [Array]$GPOtoExport
    )

    # Start Export of GPO´s
    Foreach ($GPOName in $GPOtoExport) {
    
        $GPONamePath = Join-Path $ExportPath $GPOName

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
            $Window.Close()
            Exit
        }
    }
}

# Verify Export Folder Path
Function Confirm-ExportPath {
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [String]$ExportPath
    )
    If ([string]::IsNullOrEmpty($ExportPath)) { 
        [System.Windows.Forms.MessageBox]::Show("No Export Directory Selected!" , "Export Directory Selected")
    }
    Return $true
}

# Export of  GPO Finish
Function Export-Finish {
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = “Group Policys Exported”
    $Messageboxbody = “Selected Group Policys exported"
    $MessageIcon = [System.Windows.MessageBoxImage]::Information
    $Window.Topmost = $false
    [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
    $Window.Close()
}

# Browse Button
$BrowseButton.add_Click({
    $Window.Topmost = $false
    $Global:ExportPath = Read-FolderBrowserDialog -Message "Please select a directory" -InitialDirectory $InitialDirectory
    If ([string]::IsNullOrEmpty($ExportPath)) { 
        [System.Windows.Forms.MessageBox]::Show("No Export Directory Selected!" , "Export Directory Selected")
    }
    Else {
        $ExportPathBox.Clear()
        $ExportPathBox.AppendText($ExportPath)  
    }
    $Window.Topmost = $true
})

# Export Button
$Exportbutton.add_Click({

    $GPOtoExport = @{}
    $GPOtoExport = $ListViewGPO.SelectedItems.GPOName

    $Result = Confirm-ExportPath $ExportPath

    If ($Result -eq $true) {
        If ($GPOtoExport) {
            Export-SelectedGPO -ExportPath $ExportPath -GPOtoExport $GPOToExport
            Export-Finish
        }
        Else {
            [System.Windows.Forms.MessageBox]::Show("No Group Policys Selected!" , "Group Policys")
        }
    }
    Else {
        [System.Windows.Forms.MessageBox]::Show("No Export Directory Selected!" , "Export Directory Selected")
    }
})

# Close dialog window
$CancelButton.add_click({
    $Window.Close() 
})

# Load Listview with data
Import-ListView

# Launch the window
$Window.WindowStartupLocation="CenterScreen"
$Window.Add_Loaded( {
    $this.TopMost = $true
})

# Show the window
$Window.ShowDialog() | out-null