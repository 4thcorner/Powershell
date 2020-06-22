<#
 Solution: Hyper-V 
 Purpose: Hyper-V GUI
 Version: 1.0
    Date: 26 May 2016
  Author: Tomas Johansson
 Twitter: @CanKlant
    Blog: http://www.4thcorner.se

 This script is provided "AS IS" with no warranties, confers no rights and 
 is not supported by the author
#>

# Check for elevation
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole( `
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Oupps, you need to run this script from an elevated PowerShell prompt!`nPlease start the PowerShell prompt as an Administrator and re-run the script."
	Write-Warning "Aborting script..."
    Break
}

#Check hyper-v module installed
if(!(get-module -ListAvailable -Name "hyper-v")) {
    throw "Hyper-V module not found. Please run this tool on a server that has the Hyper-V management tools installed"
}

# Add needed Assemblys
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing  
Add-Type -AssemblyName System.Windows.Forms
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Virtual Machines in Hyper-V" Height="450" Width="740">
        <Grid Name="grid" Margin="20,0,25,0.5">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label Name="Information" Content="Hyper-V Machines on:" Grid.Row="1" />
        <Label Name="Machine" Content="Machines" Grid.Row="1" Grid.ColumnSpan="2" Margin="125,0,0,0" />
        <ListView Name="ListViewMain" HorizontalAlignment="Left" Height="280" Width="680" Margin="0" VerticalAlignment="Top"  Grid.Row="2">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Hyper-V VM Name" DisplayMemberBinding="{Binding VMName}" />
                    <!--<GridViewColumn Header="VMName" DisplayMemberBinding ="{Binding VMName}"/>-->
                    <GridViewColumn Header="State" DisplayMemberBinding ="{Binding VMState}"/>
                    <GridViewColumn Header="CPU Usage" DisplayMemberBinding ="{Binding VMCPUUsage}"/>
                    <GridViewColumn Header="Assigned Memory MB" DisplayMemberBinding ="{Binding VMMemoryAssigned}"/>
                    <GridViewColumn Header="Version" DisplayMemberBinding ="{Binding VMVersion}"/>
                    <!--<GridViewColumn Header="VM Path" DisplayMemberBinding ="{Binding VMPath}"/>-->
                    <GridViewColumn Header="VM Uptime" DisplayMemberBinding ="{Binding VMUptime}"/>
                    <!--<GridViewColumn Header="VM" DisplayMemberBinding="{Binding VM}" />-->
                </GridView>
            </ListView.View>
        </ListView>
        <StackPanel Grid.Row="3" HorizontalAlignment="Left" Grid.ColumnSpan="2" Width="300">
            <Label Name="LabelAction" Content="Actions:" />
        </StackPanel>
        <GroupBox Name="groupBox_Hyper_V_Start_Stop" Header="Start/Stop VM" HorizontalAlignment="Left" Margin="0,30,0,0" Grid.Row="3" VerticalAlignment="Top" Height="64">
            <StackPanel VerticalAlignment="Top" HorizontalAlignment="Left" Orientation="Horizontal">
                <Button Name="buttonStart" Content="Start" Width="75" Margin="10,10,0,10" />
                <Button Name="buttonStop" Content="Stop" Width="75" Margin="20,10,10,10" HorizontalAlignment="Left"/>
            </StackPanel>
        </GroupBox>
        <GroupBox Name="groupBoxConnect" Header="Console" HorizontalAlignment="Left" Margin="220,30,0,0" Grid.Row="3" VerticalAlignment="Top">
            <Button Name="buttonConnect" Content="Connect" HorizontalAlignment="Center" Width="75" Margin="10"/>
        </GroupBox>
        <Button Name="buttonRefresh" Content="Refresh" HorizontalAlignment="Left" Margin="500,65,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75"/>
        <Button Name="buttonCancel" Content="Cancel" HorizontalAlignment="Right" Margin="0,65,10,0" Grid.Row="3" VerticalAlignment="Top" Width="74"/>
        <GroupBox x:Name="groupBoxDelete" Header="Delete VM" HorizontalAlignment="Left" Margin="350,30,0,0" Grid.Row="3" VerticalAlignment="Top">
            <Button Name="buttonDeleteVM" Content="Delete" HorizontalAlignment="Left" Width="75" Margin="10"/>
        </GroupBox>
    </Grid>
</Window>
'@

#Read XAML
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Form = [Windows.Markup.XamlReader]::Load($Reader)
foreach ($Name in ($XAML | Select-Xml '//*/@Name' | ForEach-Object { $_.Node.Value})) {
    New-Variable -Name $Name -Value $Form.FindName($Name) -Force
}


Function LoadListView {
$VMInfos = @()
$VMs = get-vm #-ComputerName $ComputerName

    foreach ($VM in $VMs) {
        $VMInfo = New-Object -TypeName PSObject
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMName -Value $VM.Name
        $VMInfo | Add-Member -MemberType NoteProperty -Name Name -Value $VM.Name
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMState -Value $VM.State
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMCPUUsage -Value $VM.CPUUsage
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMMemoryAssigned -Value ($VM.MemoryAssigned/1MB)
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMVersion -Value $VM.Version
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMPath -Value $VM.Path
        $VMInfo | Add-Member -MemberType NoteProperty -Name VMUptime -Value ("{0:D1}" -f $VM.Uptime.Days +  '.' + "{0:D2}" -f $VM.Uptime.Hours +  ':' + "{0:D2}" -f $VM.Uptime.Minutes +  ':' + "{0:D2}" -f $VM.Uptime.Seconds)
        $VMInfo | Add-Member -MemberType NoteProperty -Name VM -Value $VM
        $VMInfos += $VMInfo
    }
    $ListViewMain.ItemsSource = $VMInfos
}


$Machine.Content = $env:COMPUTERNAME

#Event Handlers
$Form.Add_Loaded({

})

#Sort event handler
$Form.Add_SourceInitialized({
    [System.Windows.RoutedEventHandler]$ColumnSortHandler = {

        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
           
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {

                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                                                
                # And now we actually apply the sort to the View
                $ListViewMain_DefaultView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ListViewMain.ItemsSource)
                # Change the sort direction each time they sort
                    Switch($ListViewMain_DefaultView.SortDescriptions[0].Direction)
                    {
                        "Decending" { $Direction = "Ascending" }
                        "Ascending" { $Direction = "Descending" }
                            Default { $Direction = "Ascending" }
                    }           
                $ListViewMain_DefaultView.SortDescriptions.Clear()
                $ListViewMain_DefaultView.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $Direction))
                $ListViewMain_DefaultView.Refresh()  
            }
        }
    }
    #Attach the Event Handler
    $ListViewMain.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ColumnSortHandler)
})

# Stop selected VMs
$ButtonStop.Add_Click({
    $Form.Topmost = $false
    Stop-VMs -VM $ListViewMain.SelectedItems
})

# Start selected VMs
$buttonStart.Add_Click({
    $Form.Topmost = $false
    Start-VMs -VM $ListViewMain.SelectedItems
})

# Connect to Conole on selected VM
$buttonConnect.Add_Click({
    $Form.Topmost = $false

    if ([string]::IsNullOrEmpty($ListViewMain.SelectedItem.Name)) {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Connect to Console"
        $Messageboxbody = "Please select a Machine to connect to!"
        $MessageIcon = [System.Windows.MessageBoxImage]::Exclamation
        [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)

    }
    Else {
        Start-Process -FilePath  "$env:windir\system32\vmconnect.exe" -ArgumentList $env:COMPUTERNAME,$ListViewMain.SelectedItem.Name
    }
})

$buttonDeleteVM.Add_Click({
    $Form.Topmost = $false

    if ([string]::IsNullOrEmpty($ListViewMain.SelectedItem.Name)) {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Delete Virtula Macine"
        $Messageboxbody = "Please select a Machine to Delete!"
        $MessageIcon = [System.Windows.MessageBoxImage]::Exclamation
        [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)

    }
    Else {
        Delete-VM -VM $ListViewMain.SelectedItems
    }
})

# Refresh ListView
$ButtonRefresh.Add_Click({
    LoadListView
})

# Close dialog window
$ButtonCancel.add_click({ 
    $Form.Close() 
})

$objForm = New-Object System.Windows.Forms.Form 
$objForm.KeyPreview = $True

$objForm.Add_KeyDown({
    If ($_.KeyCode -eq "Escape") {
        $objForm.Close()
    }
})

#Actions
function Stop-VMs
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [Object] $VM
    )

    Process
    {
     write-warning "Shutting down VM: $($VM.VMName) ($($VM.Name))"
     Stop-VM -Name $VM.Name
    }
    End {
        LoadListView
    }
}

Function Start-VMs {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [Object] $VM
    )
    
    Process
    {
     write-warning "Starting VM: $($VM.VMName) "
     Start-VM -Name $VM.VMName
    }

    End {
        LoadListView
    }
}

# Delete Selected Virtual Machines
Function Delete-VM {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [Object] $VM
    )
    Process {
        write-warning "Deleting VM: $($VM.VMName) ($($VM.Name))"
        $VM2Delete  = $VM
        ForEach ($VMDelet in $VM2Delete) {
            $ButtonType = [System.Windows.MessageBoxButton]::YesNo
            $MessageboxTitle = “ Delete Virtual Machine”
            $Messageboxbody = “Do you really want to delete this machine:`n`n$($VMDelet.Name)"
            $MessageIcon = [System.Windows.MessageBoxImage]::Question
            $Result = [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
            If ($Result -eq "YES") {
                DoDeleteOfVM -MachineName $VMDelet.Name
                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageboxTitle = “ Virtual Machine Deleted”
                $Messageboxbody = “Virtual Machine: $($VMDelet.Name) Deleted!"
                $MessageIcon = [System.Windows.MessageBoxImage]::Information
                [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
            }
        }
    }
    End {
        LoadListView
    }
}

# Doing Deleting ov selected VirtualMachine
Function DoDeleteofVM {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String]$MachineName
    )
    Process {
        Try {
        # Get VM information
        $VM = Get-VM -Name $MachineName -ErrorAction stop
        Write-Host  -ForegroundColor Green "Virtual machine: $($MachineName.ToUpper()) Exist on host"

        # Shutdown machine if running
        If ($VM.State -eq "Running")  {
            Write-Host  -ForegroundColor Yellow "Virtual machine: $($MachineName.ToUpper()) Running"
            DO {
                Stop-VM -Name $MachineName -TurnOff
                $VM = Get-VM -Name $MachineName -ErrorAction stop
    
            } While ($VM.State -eq "Running")
            Write-Host  -ForegroundColor Green "Virtual machine: $($MachineName.ToUpper()) is Stopped"
        }

        #Get Path for machine to delete
        $VMPath = $VM.Path

        # Removing machine from hyper-v host
        Remove-VM –vmname $MachineName -Force

        # Remove Directorys and files
        If (Test-Path $VMPath) {
            Get-ChildItem -Path $VMPath -Include * | remove-Item -recurse
            If (Test-Path $VMPath) { 
                Remove-Item -Path $VMPath
            }
        }
            Write-Host -ForegroundColor Green "Virtual machine: $($MachineName.ToUpper()) Deleted from host"
        }
        Catch
        {
            Write-Host -ForegroundColor Yellow "No Virtual machine: $($MachineName.ToUpper()) Exist on host"
        }
    }
}

#Load Listview with data
LoadListView

#Launch the window
$Form.WindowStartupLocation="CenterScreen"
$Form.Add_Loaded( {
    $this.TopMost = $true
})

# Show the window
$Form.ShowDialog() | out-null