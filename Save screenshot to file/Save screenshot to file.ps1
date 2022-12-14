$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will take a screenshot of all displays of the user session.

    .DESCRIPTION
    This script gets the dimensions of the users working display area and take a screenshot. The screenshot can be saved
    as a BMP, JPG or PNG in a location of choice. 

    .PARAMETER strScreenShotPath
    The location the screenshot should be saved, without a trailing backslash (ie. \\server\sahare\screenshots). Environment
    variables such as %USERPROFILE% may be used.

    .PARAMETER strImageType
    The desired screenshot format.
    BMP --> Large, no compression (only included for compatibility)
    JPG --> Smallest, compression artifacts
    PNG --> Larger than JPG, but losslesscompression, default

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    The working display area of a session is affected by the Scaling choice in Display settings. For example, a user
    display may be 1920x1080 with 125% scaling will result in a 1536x864 screenshot
    If the screenshots are written to a shared location make sure users can write and modify in that location, but not
    read to make sure they cannot open other users's screenshots.
#>

[string]$strScreenShotPath = $args[0]
[string]$strImageType = $args[1]

Function Feedback {
    Param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]$Message,
        [Parameter(Mandatory = $false,
        Position = 1)]
        [string]$Exception,
        [switch]$Oops
    )

    # This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If (!$error -and !$Oops) {
        # Write content of feedback string
        Write-Host $Message -ForegroundColor 'Green'
    }

    # If an error occured report it, and exit the script with ErrorLevel 1
    Else {
        # Write content of feedback string but to the error stream
        $Host.UI.WriteErrorLine($Message) 
        
        # Display error details
        If ($Exception) {
            $Host.UI.WriteErrorLine("Exception detail:`n$Exception")
        }

        # Exit errorlevel 1
        Exit 1
    }
}

Function Take-Screenshot {
    # This function takes screenshots of the passed screen. 
    # Load the assemblies required for using screen and form calls
    try {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    }
    catch {
        Feedback -Message "Required assemblies could not be loaded." -Exception $_
    }

    # Get the bounds of the entire display area
    try {
        $Screens = [System.Windows.Forms.Screen]::AllScreens

        [int]$intRightMostMonitorStart = $screens.Bounds.X | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
        [int]$intWidth = $intRightMostMonitorStart + ($Screens | Where-Object {$_.Bounds.X -eq $intRightMostMonitorStart}).Bounds.Width

        [int]$intTopMostMonitorStart = $screens.Bounds.Y | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
        [int]$intHeight = $intTopMostMonitorStart + ($Screens | Where-Object {$_.Bounds.X -eq $intTopMostMonitorStart}).Bounds.Height
    }
    catch {
        Feedback -Message 'The dimensions of the entire display area could not be retreived.' -Exception $_
    }

    # Take the screenshot 
    try {

        $bmpScreenShot = New-Object System.Drawing.Bitmap $intWidth, $intHeight
        $sizScreenShot = New-object System.Drawing.Size $intWidth, $intHeight
        $imgScreenshot = [System.Drawing.Graphics]::FromImage($bmpScreenShot)
        $imgScreenShot.CopyFromScreen(0, 0, 0, 0, $sizScreenShot)
        $bmpScreenShot
    }
    catch {
        Feedback -Message 'There was a problem taking the screenshot.' -Exception $_
    }
}

function Save-BitmapObject {
    param (
        [parameter(Mandatory = $true,
            position = 0)]
        [System.Drawing.Bitmap]$Bitmap,
        [parameter(Mandatory = $true,
            position = 1)]
        [string]$FilePathAndName,
        [parameter(Mandatory = $true,
            position = 2)]
        [string]$FileType
    )
    switch ($FileType) {
        "BMP" {$Bitmap.Save("$FilePathAndName", ([System.Drawing.Imaging.ImageFormat]::Bmp))}
        "PNG" {$Bitmap.Save("$FilePathAndName", ([System.Drawing.Imaging.ImageFormat]::Png))}
        "JPG" {$Bitmap.Save("$FilePathAndName", ([System.Drawing.Imaging.ImageFormat]::Jpeg))}
    }
}

# Check all the arguments have been passsed
if ($args.Count -ne 2) {
    Feedback -Message "The script did not get enough arguments from the Console." -Oops
}

# Set filename
[string]$strNow = (Get-Date).ToString("yyyyMMdd-HHmmss")
$strScreenShotFileName = "$strNow-%COMPUTERNAME%-%USERNAME%.$strImageType"

# Take screenshot and save it
[System.Drawing.Bitmap]$bmpScreenShot = Take-Screenshot

# Save the screenshot to desired path
try {
    $strFilePathAndName = ([System.Environment]::ExpandEnvironmentVariables("$strScreenShotPath\$strScreenShotFileName"))
    Save-BitmapObject -Bitmap $bmpScreenShot -FilePathAndName $strFilePathAndName -FileType $strImageType
    Feedback -Message "Screenshot written to $strFilePathAndName"
}
catch {
    # Failure perhaps because of filepath problem, revert to %TEMP% if this was not the filepath already
    if ($strScreenShotPath -ne '%TEMP%') {
        # Clear the error raised by first save attempt failure
        $error.clear()
        # Try saving to %TEMP% instead
        try {
            $strFilePathAndName = ([System.Environment]::ExpandEnvironmentVariables("%TEMP%\$strScreenShotFileName"))
            Save-BitmapObject -Bitmap $bmpScreenShot -FilePathAndName $strFilePathAndName -FileType $strImageType
            Feedback -Message "Screenshot written to $strFilePathAndName"
        }
        catch {
            Feedback -Message 'Screenshot could not be saved.' -Exception $_
        }
    }
    else {
        # The file path the screenshot was supposed to be saved to was already %TEMP%, so something else was going on.
        Feedback -Message 'Screenshot could not be saved.' -Exception $_
    }
}
