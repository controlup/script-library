$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will take multiple screenshots of all displays of the user session.

    .DESCRIPTION
    This script gets the dimensions of the users working display area and takes screenshots. The screenshots can be saved
    as a BMP, JPG or PNG in a location of choice. 

    .PARAMETER strScreenShotPath
    The location the screenshot should be saved, without a trailing backslash (ie. \\server\sahare\screenshots). Environment
    variables such as %USERPROFILE% may be used.

    .PARAMETER strImageType
    The desired screenshot format.
    BMP --> Large, no compression (only included for compatibility)
    JPG --> Smallest, compression artifacts
    PNG --> Larger than JPG, but losslesscompression, default

    .PARAMETER intScreenShotAmount
    The amount of screenshots to take.

    .PARAMETER intScreenShotInterval
    The time between each screenshot, in seconds.

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
[int]$intScreenShotAmount = $args[2]
# Sleeptimer uses milliseconds, convert now
[int]$intScreenShotInterval = ([int]$args[3] * 1000)

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
        "BMP" { $Bitmap.Save("$FilePathAndName", ([System.Drawing.Imaging.ImageFormat]::Bmp)) }
        "PNG" { $Bitmap.Save("$FilePathAndName", ([System.Drawing.Imaging.ImageFormat]::Png)) }
        "JPG" { $Bitmap.Save("$FilePathAndName", ([System.Drawing.Imaging.ImageFormat]::Jpeg)) }
    }
}

# Check all the arguments have been passsed
<#if ($args.Count -ne 4) {
    Feedback -Message 'The script did not get enough arguments from the Console.' -Oops
}#>

# Set filename
[string]$strNow = (Get-Date).ToString("yyyyMMdd-HHmmss")
$strScreenShotFileName = "$strNow-%COMPUTERNAME%-%USERNAME%"

# Take screenshot(s) and save them
# Load the assemblies required for using screen and form calls
try {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
}
catch {
    Feedback -Message 'Required assemblies could not be loaded.' -Oops
}

# Get the bounds of the entire display area
try {
    $Screens = [System.Windows.Forms.Screen]::AllScreens

    [int]$intRightMostMonitorStart = $screens.Bounds.X | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
    [int]$intWidth = $intRightMostMonitorStart + ($Screens | Where-Object { $_.Bounds.X -eq $intRightMostMonitorStart }).Bounds.Width

    [int]$intTopMostMonitorStart = $screens.Bounds.Y | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
    [int]$intHeight = $intTopMostMonitorStart + ($Screens | Where-Object { $_.Bounds.X -eq $intTopMostMonitorStart }).Bounds.Height
}
catch {
    Feedback -Message 'The dimensions of the entire display area could not be retreived.' -Oops
}

# Take the screenshot 
try {

    $bmpScreenShot = New-Object System.Drawing.Bitmap $intWidth, $intHeight
    $sizScreenShot = New-Object System.Drawing.Size $intWidth, $intHeight
    $imgScreenshot = [System.Drawing.Graphics]::FromImage($bmpScreenShot)
    [datetime]$dtScreenShot = Get-Date
    $imgScreenShot.CopyFromScreen(0, 0, 0, 0, $sizScreenShot)        
}
catch {
    Feedback -Message 'There was a problem taking the screenshot.' -Oops
}

# Save the first screenshot to desired path. Change path to %TEMP% if chosen path can't be written to
try {
    $strFilePathAndName = ([System.Environment]::ExpandEnvironmentVariables("$strScreenShotPath\$strScreenShotFileName"))
    Save-BitmapObject -Bitmap $bmpScreenShot -FilePathAndName "$strFilePathAndName-1.$strImageType" -FileType $strImageType
}
catch {
    # Failure perhaps because of filepath problem, revert to %TEMP% if this was not the filepath already
    if ($strScreenShotPath -ne '%TEMP%') {
        # Clear the error raised by first save attempt failure
        $error.clear()

        # Try saving to %TEMP% instead
        try {
            $strFilePathAndName = ([System.Environment]::ExpandEnvironmentVariables("%TEMP%\$strScreenShotFileName"))
            Save-BitmapObject -Bitmap $bmpScreenShot -FilePathAndName "$strFilePathAndName-1.$strImageType" -FileType $strImageType
        }
        catch {
            Feedback -Message 'Screenshots could not be saved.' -Oops
        }
    }
    else {
        # The file path the screenshot was supposed to be saved to was already %TEMP%, so something else was going on.
        Feedback -Message 'Screenshots could not be saved.' -Oops
    }
}

# Everything seems to be working fine, take more screenshots and save them if more than one was specified
If ($intScreenShotAmount -gt 1) {
    for ($i = 2; $i -le $intScreenShotAmount; $i++) {
        # Calculate the interval, based on the time the CopyFromScreen and Save file took. If it took longer than then screenshot interval set the sleep timer to 0 seconds.
        [int]$intSleepMilliseconds = ([timespan](New-TimeSpan -Start $dtScreenShot -End $(Get-Date))).TotalMilliseconds

        # Test if less time has passed than sepcified interval, if so sleep remainder of time
        if ($intSleepMilliseconds -lt $intScreenShotInterval) { Start-Sleep -Milliseconds ($intScreenShotInterval - $intSleepMilliseconds) }

        # Interval has passed, set datetime of screenshot, take the screenshot and save it
        $dtScreenShot = Get-Date
        $imgScreenShot.CopyFromScreen(0, 0, 0, 0, $sizScreenShot)

        # Save it
        Save-BitmapObject -Bitmap $bmpScreenShot -FilePathAndName "$strFilePathAndName-$($i).$($strImageType)" -FileType $strImageType
    }
}
else {
    Feedback -Message "Screenshot saved to $strFilePathAndName-1.$strImageType"
    Exit 0
}

# Script has not exited, so multiple screenshots have been taken.
Feedback -Message "Screenshots saved to $([System.Environment]::ExpandEnvironmentVariables("$strScreenShotPath"))"
