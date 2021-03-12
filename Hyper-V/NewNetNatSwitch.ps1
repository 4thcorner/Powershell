
<#
Solution: Microsoft Hyper-V Tool
 Purpose: Create New Swicth with NetNat GUI
 Version: 2.0.1
    Date: 12 Mars 2021

  Author: Tomas Johansson
 Twitter: @deploymentnoob
     Web: https://www.4thcorner.net

This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the author
#>

#Add WPF and Windows Forms assemblies
try {
    Add-Type -AssemblyName PresentationCore -ErrorAction Stop
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName system.windows.forms -ErrorAction Stop
}
catch {
    Throw "Failed to load Windows Presentation Framework assemblies."
}

# Check for elevation
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole( `
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "Aborting script..."
    $Messageboxbody = "Oupps, you need to run this script from an elevated PowerShell prompt! Please start the PowerShell prompt as an Administrator and re-run the script."
    $MessageIcon = [System.Windows.MessageBoxImage]::Exclamation
    [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
    Exit
}

if(!(get-module -ListAvailable -Name "hyper-v")) {
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "Hyper-V module not found"
    $Messageboxbody = "Please run this tool on a server that has the Hyper-V management tools installed"
    $MessageIcon = [System.Windows.MessageBoxImage]::Exclamation
    [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)
    Exit
}


# Windows GUIin XAML
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New Hyper-V NetNat Switch" Height="380" Width="450" Topmost="True" ResizeMode="NoResize">
        <Grid Height="350" Width="450"  HorizontalAlignment="Left" Margin="0,-1,-6,2">
        <Label x:Name="Label_SwitchName" Content="Switch Name:" HorizontalAlignment="Left" Margin="20,15,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="TextBox_SwitchName" HorizontalAlignment="Left" Height="23" Margin="25,37,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="300" TabIndex="0"/>
        <Label x:Name="Label_IPAdress_Gateway" Content="IP Adress of Gateway:" HorizontalAlignment="Left" Margin="20,59,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="TextBox_IPAdressGateway" HorizontalAlignment="Left" Height="23" Margin="25,82,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="300" TabIndex="1"/>
        <Label x:Name="Label_PrefixxLength" Content="Prefix:" HorizontalAlignment="Left" Margin="20,103,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="TextBox_PrefixLength" HorizontalAlignment="Left" Height="23" Margin="25,127,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="50" TabIndex="2"/>
        <Label x:Name="Label_NewSwitchName" Content="Switch Name:" HorizontalAlignment="Left" Margin="76,160,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_NewSwitchName_Result" Content="" HorizontalAlignment="Left" Margin="155,160,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_Gateway_IPAdress" Content="Gateway IP Address:" HorizontalAlignment="Left" Margin="42,200,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_Gateway_IPAddress_Result" Content="" HorizontalAlignment="Left" Margin="155,201,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_Network" Content="Network:" HorizontalAlignment="Left" Margin="100,180,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_IPAddressSpace" Content="" HorizontalAlignment="Left" Margin="155,180,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_InternalAdressPrefix" Content="Internal IP Adress Prefix:" HorizontalAlignment="Left" Margin="21,220,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_InternalAdressPrefix_Result" Content="" HorizontalAlignment="Left" Margin="155,220,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_Subnetmask" Content="Subnetmask:" HorizontalAlignment="Left" Margin="81,240,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_InternalAdressPrefix_Calculated" Content="" HorizontalAlignment="Left" Margin="156,240,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_Create_Information" Content="" HorizontalAlignment="Center" Margin="20,279,0,0" VerticalAlignment="Top"/>
        <Button x:Name="Button_CreateNetNat" Content="Create" HorizontalAlignment="Left" Margin="260,320,0,0" VerticalAlignment="Top" Width="75" TabIndex="3"/>
        <Button x:Name="Button_Cancel" Content="Cancel" HorizontalAlignment="Left" Margin="350,320,0,0" VerticalAlignment="Top" Width="75" TabIndex="4"/>
    </Grid>
</Window>
'@

# Create Hashtable and Runspace for GUI
$SyncHash = [hashtable]::Synchronized(@{})
$NewRunspace =[runspacefactory]::CreateRunspace()
$NewRunspace.ApartmentState = "STA"
$NewRunspace.ThreadOptions = "ReuseThread"         
$NewRunspace.Open()
$NewRunspace.SessionStateProxy.SetVariable("syncHash",$SyncHash)      

# Create the XAML reader using a new XML node reader
$XAMLReader = New-Object System.Xml.XmlNodeReader $XAML
$SyncHash.Window = [Windows.Markup.XamlReader]::Load($XAMLReader)
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | 
ForEach-Object {
    $SyncHash.Add($_.Name,$SyncHash.Window.FindName($_.Name))
}

$SyncHash.TextBox_SwitchName.Add_TextChanged({
    $NewswitchName = $SyncHash.TextBox_SwitchName.Text
    $ExistName = Get-VMSwitch | Where-Object {$_.Name -like "$NewswitchName"} | Select-Object -ExpandProperty Name
    
    If ([string]::IsNullOrEmpty($ExistName)) {
        $SwitchExist = $False
    }
    Else {
        $SwitchExist = $True
    }

    $SyncHash.Window.Dispatcher.Invoke(
        [action]{
            If ($SwitchExist -eq $True) {
                $SyncHash.Label_NewSwitchName_Result.Foreground = "Red"
                $SyncHash.Label_NewSwitchName_Result.Content = "Switch with that name already exist!"
                $SyncHash.Button_CreateNetNat.IsEnabled = $False
            }
            Else {
                $SyncHash.Label_NewSwitchName_Result.Foreground = "Black"
                $SyncHash.Label_NewSwitchName_Result.Content = $SyncHash.TextBox_SwitchName.Text
                $SyncHash.Button_CreateNetNat.IsEnabled = $True
            }
        }
    )
})

$SyncHash.TextBox_PrefixLength.Add_TextChanged({
    $Bitmask = $SyncHash.TextBox_PrefixLength.Text

    If (!([string]::IsNullOrEmpty($Bitmask))) {
        $Result = Test-Bitmask $Bitmask
    }
    If ($Result -eq $true) {
        $SubnetMask = ConvertTo-IPv4MaskString $SyncHash.TextBox_PrefixLength.Text
        
    }
    $ValidSubnet = Test-IPv4MaskString $SubnetMask

    $SyncHash.Window.Dispatcher.Invoke(
        [action]{
            $SyncHash.Label_InternalAdressPrefix_Result.Content = $SyncHash.TextBox_PrefixLength.Text
            If ($ValidSubnet -eq $true) {
                $SyncHash.Label_InternalAdressPrefix_Result.Foreground = "Black"
                $SyncHash.Label_InternalAdressPrefix_Calculated.Foreground = "Black"
                $SyncHash.Label_InternalAdressPrefix_Calculated.Content = $SubnetMask
                $SyncHash.Button_CreateNetNat.IsEnabled = $True
            }
            Else {
                $SyncHash.Label_InternalAdressPrefix_Result.Foreground = "Red"
                $SyncHash.Label_InternalAdressPrefix_Calculated.Foreground = "Red"
                $SyncHash.Label_InternalAdressPrefix_Calculated.Content = "Not Valid subnetmask"
                $SyncHash.Button_CreateNetNat.IsEnabled = $False
            }
        }
    )
})

$SyncHash.TextBox_IPAdressGateway.Add_TextChanged({
    $GWIPAddress = $SyncHash.TextBox_IPAdressGateway.Text

    $VaildIPAdress = Test-GateWayIPAdress $GWIPAddress

    If ($VaildIPAdress -eq $true) {
        $GatewayIPExist = Test-NetIPAdress $GWIPAddress
        If ($GatewayIPExist -eq $False) {
            $IPAddressSpace = Resolve-IPv4NetworkSpace $GWIPAddress
            $SyncHash.Label_Gateway_IPAddress_Result.Foreground = "Black"
            $ResultIPAdress = $GWIPAddress
        }
        Else {
            $SyncHash.Label_Gateway_IPAddress_Result.Foreground = "Red"
            $ResultIPAdress = "IP Adress already exist!"
            $IPAddressSpace = "-"
        }
    }

    $SyncHash.Window.Dispatcher.Invoke(
        [action]{
            If ($VaildIPAdress -eq $true) {
                $SyncHash.Label_Gateway_IPAddress_Result.Content = $ResultIPAdress
                $SyncHash.Label_IPAddressSpace.Content = $IPAddressSpace
                $SyncHash.Button_CreateNetNat.IsEnabled = $True
            }
            Else {
                $SyncHash.Label_Gateway_IPAddress_Result.Content = $ResultIPAdress
                $SyncHash.Button_CreateNetNat.IsEnabled = $False
            }
            If ($GatewayIPExist -eq $True) {
                $SyncHash.Button_CreateNetNat.IsEnabled = $False
            }
            Else {
                $SyncHash.Button_CreateNetNat.IsEnabled = $True
            }
        }
    )
})


# Close Dialog Window
$SyncHash.button_cancel.add_click({
    $SyncHash.Window.Close() 
})


# Creat NetNat switch
$SyncHash.Button_CreateNetNat.add_click({
    $SwitchName = $SyncHash.TextBox_SwitchName.text
    $NetworkGatewayIPAddress = $SyncHash.TextBox_IPAdressGateway.Text
    $NetworkPrefix = $SyncHash.TextBox_PrefixLength.Text

    # Network Adress Space
    $NetworkIPAddressSpace = Resolve-IPv4NetworkSpace $NetworkGatewayIPAddress

    # Check if NetNat already exist
    $NetNatExist = Test-NetNat $NetworkIPAddressSpace $NetworkPrefix

    If ($NetNatExist -eq $False) {
        # Create 
        $NetNatCreated = New-NetNatNetwork -SwitchName $SwitchName -GatewayIPAddress $NetworkGatewayIPAddress -IPAddressSpace $NetworkIPAddressSpace -PrefixLengt $NetworkPrefix

        If ($NetNatCreated -eq $True) {
            $SyncHash.Label_Create_Information.Foreground = "Black"
            $ResultText = "NetNat created for switch: " + $SwitchName
        }
        Else {
            $SyncHash.Label_Create_Information.Foreground = "Red"
            $ResultText = "Failed to Create NetNat for switch: " + $SwitchName
        }
    }
    Else {
        $ResultText = "NetNat: " + $NetNatExist + " Exist already"
    }

    $SyncHash.Window.Dispatcher.Invoke(
        [action]{
            $SyncHash.Label_Create_Information.Content = $ResultText
        }
    )
})


Function New-NetNatNetwork {
    Param (
        [string]$SwitchName,
        [string]$GatewayIPAddress,
        [string]$IPAddressSpace,
        [byte]$PrefixLengt
    )
    BEGIN {
        $Retval = $Null
        $NetNatName = $SwitchName + "-Network"
        $InterfaceAddressPrefix = $IPAddressSpace + "/" + $PrefixLengt
    }
    PROCESS {
        Try {
            $Retval = (New-VMSwitch -SwitchName $SwitchName -SwitchType Internal -ErrorAction Stop).Name
            If (!([string]::IsNullOrEmpty($Retval))) {
                Try {
                    $InterfaceIndex = $Null
                    $InterfaceIndex = (Get-NetAdapter | Where-Object {$_.Name -Like "vEthernet ($SwitchName)"}).InterfaceIndex
                    If (!([string]::IsNullOrEmpty($InterfaceIndex))) {
                        Try {
                            $Retval = $Null
                            $Retval = (New-NetIPAddress -IPAddress $GatewayIPAddress -PrefixLength $PrefixLengt -InterfaceIndex $InterfaceIndex -ErrorAction Stop).Count
                            If (!([string]::IsNullOrEmpty($Retval))) {
                                Try {
                                    $Retval = $Null
                                    Try {
                                        $Null = New-NetNat -Name $NetNatName -InternalIPInterfaceAddressPrefix $InterfaceAddressPrefix -ErrorAction Stop
                                        $Result = $True
                                    }
                                    Catch {
                                        $Result = $False
                                    }
                                }
                                Catch {
                                    $Result = $False
                                }
                            }
                            Else {
                                $Result = $False 
                            }
                        }
                        Catch {
                            $Result = $False
                        }
                    }
                    Else {
                        $Result = $False
                    }
                }
                Catch {
                    $Result = $False
                }   
            }
            Else {
                $Result = $False
            }
        }
        Catch {
            $Result = $False
        }
    }
    END {
        Return $Result
    }
}

Function Test-NetNat {
    Param (
        [String]$NetworkSpace,
        [string]$NetworkPrefix
    )
    BEGIN{
        $Retval = $Null
        $InternalIPInterfaceAddressPrefix = $NetworkSpace + "/" + $NetworkPrefix
    }
    PROCESS{
        Try {
            $Retval = Get-NetNat | Where-Object {$_.InternalIPInterfaceAddressPrefix -match $InternalIPInterfaceAddressPrefix} | Select-Object -ExpandProperty Name -ErrorAction Stop
            If(!([string]::IsNullOrEmpty($Retval))) {
                $Retval = $True
            
            }
            Else {
                $Retval = $False
            }
        }
        Catch {
            $Retval = $False
        }
    }
    END{
        Return $Retval
    }
}

Function Test-NetIPAdress {
    Param (
        [string]$GatewayIPAddress
    )
    BEGIN {
    }
    PROCESS {
        Try {
            $Null = Get-NetIPAddress -IPAddress $GatewayIPAddress -ErrorAction Stop
            $Retval = $true 
        }
        Catch {
            $Retval = $False
        }
    }
    END {
        Return $Retval
    }
}



Function Test-GateWayIPAdress {
<#
.SYNOPSIS
    Tests whether an IPv4 Gateway Adress string (e.g., "192.168.10.1") is valid.
.DESCRIPTION
    ests whether an IPv4 Gateway Adress string (e.g., "192.168.10.1") is valid.
.PARAMETER GWIPAddress
    Specifies the IPv4 Gateway IPAddress (e.g., "192.168.10.1").
#>
    Param (
        [String]$GWIPAddress
    )
    BEGIN {
        $Pattern = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"
    }
    PROCESS {
        $Retval = $GWIPAddress -match $pattern
    }
    END {
        Return $Retval
    }
}
Function ConvertTo-IPv4MaskString {
<#
.SYNOPSIS
    Converts a number of bits (0-32) to an IPv4 network mask string (e.g., "255.255.255.0").
.DESCRIPTION
    Converts a number of bits (0-32) to an IPv4 network mask string (e.g., "255.255.255.0").
.PARAMETER MaskBits
    Specifies the number of bits in the mask.
#>
    param(
    [parameter(Mandatory=$true)]
    [ValidateRange(0,32)]
    [Int] $MaskBits
    )
    BEGIN {
    }
    PROCESS {
    $mask = ([Math]::Pow(2, $MaskBits) - 1) * [Math]::Pow(2, (32 - $MaskBits))
    $bytes = [BitConverter]::GetBytes([UInt32] $mask)
    $Retval = (($bytes.Count - 1)..0 | ForEach-Object { [String] $bytes[$_] }) -join "."
    }
    END {
        Return $Retval
    }
}

Function Test-IPv4MaskString {
<#
.SYNOPSIS
    Tests whether an IPv4 network mask string (e.g., "255.255.255.0") is valid.
.DESCRIPTION
    Tests whether an IPv4 network mask string (e.g., "255.255.255.0") is valid.
.PARAMETER MaskString
    Specifies the IPv4 network mask string (e.g., "255.255.255.0").
#>
    Param (
    [string]$MaskString
    )
    BEGIN{
    }
    PROCESS {
        $bValidMask = $true
        $ArrSections = @()
        $ArrSections +=$MaskString.split(".")
        
        # Firstly, make sure there are 4 sections in the subnet mask
        if ($ArrSections.count -ne 4) {
            $bValidMask =$False
        }
    
        # Secondly, make sure it only contains numbers and it's between 0-255
        if ($bValidMask) {
            foreach ($item in $arrSections) {
                if(!($item  -match "^\d+$")) {
                   $bValidMask = $False
                }                  
            }
        }
    
        if ($bValidMask) {
            foreach ($item in $arrSections)
            {
                $item = [int]$item
                if ($item -lt 0 -or $item -gt 255) {$bValidMask = $False}
            }
        }
    
        # Lastly, make sure it is actually a subnet mask when converted into binary format
        if ($bValidMask) {
            foreach ($item in $arrSections)
            {
                $binary = [Convert]::ToString($item,2)
                if ($binary.length -lt 8)
                {
                    do {
                    $binary = "0$binary"
                    } while ($binary.length -lt 8)
                }
                $strFullBinary = $strFullBinary+$binary
            }
            if ($strFullBinary.contains("01")) {$bValidMask = $False}
            if ($bValidMask)
            {
                $strFullBinary = $strFullBinary.replace("10", "1.0")
                if ((($strFullBinary.split(".")).count -ne 2)) {$bValidMask = $False}
            }
        }
    }
    END {
        Return $bValidMask
    }
}
Function Test-Bitmask {
<#
.SYNOPSIS
    Test if bitmask is in the range between 1 and 32 
.DESCRIPTION
    Validate if Bitmask is in valid range
.PARAMETER Bitmask
    Set Bitmask, Must be between 1-32)
.INPUTS
    Lenght : Interger
.OUTPUTS
    Retval    : boolen
.EXAMPLE
    Test-Bitmask 24
#>
    Param (
    [string]$Bitmask
    )
    BEGIN {
    }
    PROCESS {
        If ($Bitmask -lt 32) {
            $Retval = $true
            }
        Else {
            $Retval = $False
        }
    }
    END {
        Return  $Retval
    }
}

Function Resolve-IPv4NetworkSpace {
<#
.SYNOPSIS
    Get IPv4 Address Space from Gateway IPAddress 
.DESCRIPTION
    Get IPv4 Address Space from Gateway IPAddress 
.PARAMETER $GWIPAddress
    IP Address  of Gateway
.INPUTS
    GWIPAddress : String
.OUTPUTS
    Retval    : String
.EXAMPLE
    Resolve-IPv4NetworkSpace 192.168.10.1
#>
    Param (
        [string]$GWIPAddress
    )
    BEGIN {
    }
    PROCESS {
        # Last IP octect for Switch IPAddress
        $LastOctect = '0'
        # Get IPAdress Space forNetwork  
        $Retval = ($GWIPAddress -replace '(\d+\.\d+\.\d+\.)(\d+)','$1')+$LastOctect
    }
    END {
        Return $Retval
    }
}

# Launch the window
$SyncHash.Window.WindowStartupLocation="CenterScreen"
$SyncHash.Window.Add_Loaded( {
    $this.TopMost = $true
})

# Set Focus on first textbox
$SyncHash.TextBox_SwitchName.Focus() | out-null

# Disable Create NetNat button
$SyncHash.Button_CreateNetNat.IsEnabled = $False

# Show the window
$SyncHash.Window.ShowDialog() | out-null