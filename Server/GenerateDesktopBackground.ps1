<#
Solution: Create Desktop Image
 Purpose: Group Policys Export GUI
 Version: 1.0.0
    Date: 05 March 2021

  Author: Tomas Johansson
 Twitter: @deploymentnoob
     Web: https://www.4thcorner.net

This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the author
#>

# Add Microsoft .NET Core class to PowerShell session
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Computer Information
$OS = Get-CimInstance Win32_OperatingSystem
$OrganizationName = ($OS.Organization).ToUpper()
$ComputerName = ($os.CSName).ToUpper()
$Domain = ($env:USERDNSDOMAIN).ToUpper()
$ComputerName = ($os.CSName).ToUpper()
$ServerOS = ($OS.Caption).ToUpper()
$Version = ($OS.Version).ToUpper()
$InstallDate = ($OS.InstallDate).ToString('yyyy-MM-dd HH:mm')
$Boot = ($os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm')

$ComputerInfo = ([ordered]@{
    ComputerName = $ComputerName
    Domain = $Domain
    ServerOS = $ServerOS
    Version = $Version
    Installed = $InstallDate
    #'Last Boot' = $Boot
})

# Get Screen Resolution 
$Resolution = [System.Windows.Forms.Screen]::AllScreens | Where-Object Primary | Select-Object -ExpandProperty Bounds | Select-Object Width,Height

# Desktop Wallpaper Settings
$WallPaperWith = $Resolution.Width
$WallPaperHeigth = $Resolution.Height
$WallPaperColorHex = "#003C50"

# Font Settings
# Fonts must be installed on the computer 
$FontName = "Forsvarsmakten Sans Stencil"
$FontSize = 16
$FontStyle = [System.Drawing.FontStyle]::Regular
$FontUnit = [System.Drawing.GraphicsUnit]::Point
$FontColor = "White"

# Text Margin from Edge of screen to text
$MarginTop = 20
$MarginLeft= 20
$MarginRight = 20
$MarginBottom = 20

# Text Padding
[FLOAT]$textPaddingRight = 10
[FLOAT]$textPaddingLeft = 10
[FLOAT]$textPaddingTop = 10
[FLOAT]$textItemSpace = 3

# Wallpaper Save Path
# Wallpaper Image Type and Path
# Types JPEG or PNG
$SaveImageType = "jpeg"
Switch ($SaveImageType) {
    jpeg {$fileExtension = "jpg" }
     png {$fileExtension = "png" }
}

# WallpaperName and Save Path
$WallPaperFileName = $OrganizationName + '_Desktop_' +  $WallPaperWith + 'x' + $WallPaperHeigth + '.' + $FileExtension
$SavePath = "D:\Powershell\Background"
$SaveFile = Join-Path $SavePath $WallPaperFileName

# For tests Draw a Rectangle Around Text 
$DrawBoxRectangle = $false
$DrawTextRectangle = $false
$ShowDiskInfo = $false

# Create new Bitmap
$Background = New-Object System.Drawing.Bitmap $WallPaperWith,$WallPaperHeigth
# Encapsulates a GDI+ drawing surface
$Graphics = [System.Drawing.Graphics]::FromImage($Background)

# Wallpaper Background Color in HEX
$WallpaperBgColor = [System.Drawing.ColorTranslator]::FromHtml($WallPaperColorHex)

# Brush for Background
$WallpaperBrush = New-Object Drawing.SolidBrush($WallpaperBgColor)
# Brush for Text
$TextBrush = [System.Drawing.Brushes]::$FontColor

# Create Background for Desktop Wallpaper
$Graphics = [System.Drawing.Graphics]::FromImage($Background)
$Graphics.FillRectangle($WallpaperBrush,0,0,$Background.Width,$Background.Height)

# Encapsulates text layout information (such as alignment, orientation and tab stops) 
$LeftTextFormat = new-object system.drawing.stringformat
$LeftTextFormat.Alignment = [system.drawing.StringAlignment]::Far
$LeftTextFormat.LineAlignment = [system.drawing.StringAlignment]::Near

$RightTextFormat = new-object system.drawing.stringformat
$RightTextFormat.Alignment = [system.drawing.StringAlignment]::Near
$RightTextFormat.LineAlignment = [system.drawing.StringAlignment]::Near

# Image Quality
$Graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$Graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
$Graphics.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAlias
$Graphics.CompositingQuality = [Drawing.Drawing2D.CompositingQuality]::HighQuality

# Current Width and Hight of Taskbar
$Taskbar = [System.Windows.Forms.Screen]::AllScreens | Where-Object {$_.Primary -eq $true }
$TaskbarOffset = $Taskbar.Bounds.Height - $Taskbar.WorkingArea.Height

# Set Default Values
$MaxKeyWidth = 0
$MaxValueWidth = 0
$TextBgHeight = 0
$TextBgWidth = 0

ForEach ($Info in $ComputerInfo.GetEnumerator()) {
    
    $KeyString = $Info.Key + ": "
    $KeyFont = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle,$FontUnit)
    $KeySize = [system.windows.forms.textrenderer]::MeasureText($keyString, $keyFont)
    $MaxKeyWidth = [math]::Max($MaxKeyWidth, $keySize.Width)

    $ValueString = $Info.Value
    $ValueFont = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle,$FontUnit)
    $ValueSize = [system.windows.forms.textrenderer]::MeasureText($ValueString, $ValueFont)
    $MaxValueWidth = [math]::Max($MaxValueWidth, $ValueSize.Width)

    $MaxItemHeight = [math]::Max($valueSize.Height, $keySize.Height)
    $TextBgHeight += ($MaxItemHeight)

}

If ($ShowDiskInfo -eq $True ) {

    # Get Volume onformation for Computer
    $Disks = Get-Volume | Where-Object { $_.FileSystemLabel -notMatch"Recovery" -and ($_.DriveType -notmatch "CD-ROM")}

    # Create Hashtable
    $DiskData =  [ordered]@{}

    # Insert in to Hashtable
    ForEach ($Disk in $Disks) {
        $Drive = "Disk " + $Disk.DriveLetter + ":  "
        $DriveSize  = "Disk Size: " + $([math]::round($Disk.Size/1GB,2)) + " GB" + " Free Space: " + $([math]::round($Disk.SizeRemaining/1GB,2)) + " GB"

        $diskData.Insert(0, $Drive, $DriveSize)
    }

    ForEach ($Disk in $DiskData.GetEnumerator()) {
        $KeyString = $Disk.Name
        $KeyFont = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle,$FontUnit)
        $KeySize = [system.windows.forms.textrenderer]::MeasureText($keyString, $keyFont)
        $MaxKeyWidth = [math]::Max($MaxKeyWidth, $keySize.Width)
    
        $ValueString = $Info.Value
        $ValueFont = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle,$FontUnit)
        $ValueSize = [system.windows.forms.textrenderer]::MeasureText($ValueString, $ValueFont)
        $MaxValueWidth = [math]::Max($MaxValueWidth, $ValueSize.Width)

        $MaxItemHeight = [math]::Max($valueSize.Height, $keySize.Height)
        $TextBgHeight += ($MaxItemHeight)
    
    }
}

$TextBgWidth = $MaxKeyWidth + $MaxValueWidth + $textPaddingRight
$TextBgX = $Resolution.Width - $TextBgWidth - $MarginRight
$textBgY = $Resolution.Height - $TextBgHeight - $MarginBottom - $TaskbarOffset


If ($DrawBoxRectangle -eq $true) {
    $BoxRectangle = New-Object System.Drawing.RectangleF($TextBgX, $textBgY, $TextBgWidth, $TextBgHeight)
    $Pen = New-Object System.Drawing.Pen Yellow, 1
    $Graphics.DrawRectangle($Pen, $BoxRectangle.Left, $BoxRectangle.Top, $BoxRectangle.Width, $BoxRectangle.Height)
}


# Set CumulativeHeight to Empty
$CumulativeHeight = 0

ForEach ($Info in $ComputerInfo.GetEnumerator()) {

    $KeyString = $Info.Name + ": "
    $ValueString = $info.Value

    $KeyX = $TextBgX 
    $KeyY = $textBgY + $CumulativeHeight

    $ValueX = $KeyX + $MaxKeyWidth
    $ValueY = $KeyY

    # Calculate Rectangel for Text
    $KeyRectangle = New-Object System.Drawing.RectangleF($KeyX, $KeyY, $MaxKeyWidth, $KeySize.Height)
    $ValueRectangle = New-Object System.Drawing.RectangleF($ValueX, $ValueY, $($MaxValueWidth + $textPaddingRight),  $MaxItemHeight)

    # Insert Rectangel in Image
    If ($DrawTextRectangle -eq $true ) {
        $Pen = New-Object System.Drawing.Pen White, 1
        $Graphics.DrawRectangle($Pen, $KeyRectangle.Left, $KeyRectangle.Top, $KeyRectangle.Width, $KeyRectangle.Height)
        $Graphics.DrawRectangle($Pen, $ValueRectangle.Left, $ValueRectangle.Top, $ValueRectangle.Width, $ValueRectangle.Height)
    }

    # Insert Text in Image
    $Graphics.DrawString($KeyString, $KeyFont, $TextBrush, $KeyRectangle, $LeftTextFormat)
    $Graphics.DrawString($ValueString, $ValueFont, $TextBrush, $ValueRectangle , $RightTextFormat)

    $CumulativeHeight += $MaxItemHeight
}


If ($ShowDiskInfo -eq $True ) {

    ForEach ($Disk in $DiskData.GetEnumerator()) {

        $KeyString = $Disk.Name
        $ValueString = $Disk.Value

        $KeyX = $TextBgX 
        $KeyY = $textBgY + $CumulativeHeight

        $ValueX = $KeyX + $MaxKeyWidth
        $ValueY = $KeyY

        # Calculate Rectangel for Text
        $KeyRectangle = New-Object System.Drawing.RectangleF($KeyX, $KeyY, $MaxKeyWidth, $MaxItemHeight)
        $ValueRectangle = New-Object System.Drawing.RectangleF($ValueX, $ValueY, $($MaxValueWidth + $textPaddingRight),  $MaxItemHeight)

        # Insert Rectangel in Image
        If ($DrawRectangel -eq $true ) {
            $Pen = New-Object System.Drawing.Pen White, 1
            $Graphics.DrawRectangle($Pen, $KeyRectangle.Left, $KeyRectangle.Top, $KeyRectangle.Width, $KeyRectangle.Height)
            $Graphics.DrawRectangle($Pen, $ValueRectangle.Left, $ValueRectangle.Top, $ValueRectangle.Width, $ValueRectangle.Height)
        }

        # Insert Text in Image
        $Graphics.DrawString($KeyString, $KeyFont, $TextBrush, $KeyRectangle, $LeftTextFormat)
        $Graphics.DrawString($ValueString, $ValueFont, $TextBrush, $ValueRectangle , $RightTextFormat)

        $CumulativeHeight += $MaxItemHeight

    }
}


# Close Graphics
$Graphics.Dispose()

# Save and close Bitmap
$Background.Save($SaveFile, [system.drawing.imaging.imageformat]::$SaveImageType)
$Background.Dispose()


# Modify Path to the picture accordingly to reflect your infrastructure
$ImgPath = $SaveFile
$Code = @' 
using System.Runtime.InteropServices; 
namespace Win32{ 
    
     public class Wallpaper{ 
        [DllImport("user32.dll", CharSet=CharSet.Auto)] 
         static extern int SystemParametersInfo (int uAction , int uParam , string lpvParam , int fuWinIni) ; 
         
         public static void SetWallpaper(string thePath){ 
            SystemParametersInfo(20,0,thePath,3); 
         }
    }
 } 
'@

add-type $code 

#Apply the Change on the system 
[Win32.Wallpaper]::SetWallpaper($imgPath)