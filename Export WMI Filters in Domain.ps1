<#
Solution: Microsoft Deployment Toolkit
 Purpose: Export WMI Filters in Domain
 Version: 1.0.0
    Date: 20 July 2020

  Author: Tomas Johansson
 Twitter: @deploymentnoob
     Web: https://www.4thcorner.net

This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the author
#>

# Add Needed Assemblys
Add-Type -AssemblyName PresentationFramework

# Check for elevation
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    [System.Windows.MessageBox]::Show("Oupps!`nYou need to run this Script from an Elevated PowerShell!`nPlease start the PowerShell as an Administrator`nAborting script...",'Export of WMI Filters','OK',[System.Windows.Forms.MessageBoxIcon]::Warning)
    Break
}

# Get ScriptPath
$Path = split-path -parent $MyInvocation.MyCommand.Path

# Seyr Initial Directory
$InitialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

# Show an Open Folder Dialog and return the directory selected by the user.
Function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton) {
    $browseForFolderOptions = 0
    If ($NoNewFolderButton) {
        $browseForFolderOptions += 512
    }
    $App = New-Object -ComObject Shell.Application
    $folder = $App.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    If ($Folder) {
        $selectedDirectory = $folder.Self.Path
    }
    Else {
        $selectedDirectory = $Null
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($App) > $null
    Return $selectedDirectory
}


# Set Export Path
$ExportPath = Read-FolderBrowserDialog

# If Cancel Button are Export Path
If ([string]::IsNullOrEmpty($ExportPath)) {
    Break
}

# Get WMI Filters in Domain
$WmiFilters = Get-ADObject -Filter {objectClass -eq "msWMI-Som"} -Properties "msWMI-Author","msWMI-ChangeDate","msWMI-CreationDate","msWMI-ID","msWMI-Name","msWMI-Parm1","msWMI-Parm2"

if ($WmiFilters -ne $null -and $WmiFilters.Count -gt 0) {
    Write-Host "Backing up WMI Filters."
    $WmiBackupFile = New-Item -Path "$ExportPath\WmiFilters.json" -ItemType File -Force -Credential $Credential

    $WmiFilters = $WmiFilters | Select-Object -Property "msWMI-Author","msWMI-ChangeDate","msWMI-CreationDate","msWMI-ID","msWMI-Name","msWMI-Parm1","msWMI-Parm2"

    $Json = ConvertTo-Json -InputObject $WmiFilters 
    $Json = $Json.Replace("null", "`"`"")
    Set-Content -Value $Json -Path $WmiBackupFile -Force -Credential $Credential 

    [System.Windows.MessageBox]::Show("Completed WMI Filter backup.",'Export of WMI Filters','OK',[System.Windows.Forms.MessageBoxIcon]::Asterisk)
}
Else {
    [System.Windows.MessageBox]::Show("No WMI filters in the domain.",'Export of WMI Filters','OK',[System.Windows.Forms.MessageBoxIcon]::Asterisk)
}