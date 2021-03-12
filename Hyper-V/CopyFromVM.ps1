<#
Solution: Microsoft Hyper-V Tool
 Purpose: Copy from Hyper-V Guest
 Version: 2.0.0
    Date: 12 March 2021

  Author: Tomas Johansson
 Twitter: @deploymentnoob
     Web: https://www.4thcorner.net

This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the author
#>

# Define a Microsoft .NET Core class in your PowerShell session
Add-Type -AssemblyName Microsoft.VisualBasic

# This command uses the PromptForCredential method to prompt the user for their user name and password.
# The command saves the resulting credentials in the $Credential variable.
$Credential = $Host.UI.PromptForCredential("Need credentials", "Please enter your User name and Password.", "", "NetBiosUserName")

# Server to use PSSession
$ConnectToServer = [Microsoft.VisualBasic.Interaction]::InputBox('Input Server to connect to', 'Server Name', "")

# Create new PSSession to Server with specific Credential
$Session = New-PSSession -VMName $ConnectToServer -Credential $Credential

#  start interactive session with single remote computer
Enter-PSSession  -Session $Session

# Copy Items from Temp Diretory from remote computer and if file exist ower write files
Copy-Item -FromSession $Session -Path "C:\TEMP\*" -Destination "D:\Temp" -Force

# closes the current session. also closes the connection between the local and remote computers
Remove-PSSession -Session $Session
