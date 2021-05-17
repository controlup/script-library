#requires -version 3
$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    This script will create a video of all displays of the user session.

    .DESCRIPTION
    This script creates a video of the entire user display area using ffmpeg and saves it in a location of choice.

    .PARAMETER strVideoFilePath
    The location the video should be saved, without a trailing backslash (ie. \\server\share\videos). Environment variables such as %USERPROFILE% may be used.

    .PARAMETER intVideoDuration
    Duration of the video recording in seconds.

    .PARAMETER intVideoFPS
    The framerate (frames per second) of the video. 15 fps is minimum due to some video players having issues with low frame rate videos

    .PARAMETER strffMPEGLocation
    The location of ffmpeg.exe, without a trailing backslash (ie. C:\Temp\ffmpeg). Environment variables such as %USERPROFILE% may be used.

    .PARAMETER bolScale
    If set to True, video will be half the resolution of the desktop dimensions, almost halving CPU use and file size. Should be fine for most use cases.

    .PARAMETER bolUseNVIDIAHardwareEncoding
    Use NVidia GPU for mpeg encoding of the video. If hardware decoding cannot be used, ffpmeg wil exit with an error.

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    ffmpeg.exe is used for creating the video, it can be downloaded from here:
    https://www.ffmpeg.org/

    Hardware encoding should work with NVidia GPUs that support NVENC, but may fail due to errors in ffMPEG encoder detection. Do not assume your system does not support NVENC on the basis of ffMPEG hardware encoding failure.
#>

[string]$strVideoFilePath = $args[0]
[int]$intVideoDuration = $args[1]
[int]$intVideoFPS = $args[2]
[string]$strffMPEGLocation = $args[3]
If ($args[4] -eq 'True') {[bool]$bolScale = $true} Else {[bool]$bolScale = $false}
If ($args[5] -eq 'True') {[bool]$bolUseNVIDIAHardwareEncoding = $true} Else {[bool]$bolUseNVIDIAHardwareEncoding = $false}

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

function Remove-EmptyVideoFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$VideoFile
    )
    
    # try to remove the empty video file
    try {
        Remove-Item $VideoFile
        [string]$out = "The empty video file $VideoFile was removed."
    }
    catch {
        [string]$out = "The empty video file $VideoFile could NOT be removed."
    }
    $out
}

# Check all the arguments have been passsed
    if ($args.Count -ne 6) {
    Feedback -Message "The script did not get enough arguments from the Console." -Oops
    }

# Set filename and test for writing.
[string]$strNow = (Get-Date).ToString("yyyyMMdd-HHmmss")
$strVideoFilename = "$strNow-%COMPUTERNAME%-%USERNAME%.mkv"

# Test if ffMPEG.exe is available
if (!(Test-Path $strffMPEGLocation\ffmpeg.exe)) {
    Feedback -Message "FFMPEG.EXE was not found in $strffMPEGLocation" -Oops
}

# Set path and test write
try {
    [string]$strFilePathAndName = ([System.Environment]::ExpandEnvironmentVariables("$strVideoFilePath\$strVideoFileName"))
    [io.file]::OpenWrite($strFilePathAndName).close()
}
catch {
    # Trying something else, clear the error
    $error.Clear()
    # Failure perhaps because of filepath problem, revert to %TEMP% if this was not the filepath already
    if ($strVideoFilePath -ne '%TEMP%') {
        # Try saving to %TEMP% instead
        try {
            [string]$strFilePathAndName = ([System.Environment]::ExpandEnvironmentVariables("%TEMP%\$strVideoFileName"))
            [io.file]::OpenWrite($strFilePathAndName).close()
        }
        catch {
            Feedback -Message "There was an issue trying to write to both $strVideoFilePath and the fallback path %TEMP%" -Oops
        }
    }
    else {
        # The file path the screenshot was supposed to be saved to was already %TEMP%, so something else was going on.
        Feedback -Message 'The location for saving the video could not be written to.' -Oops
    }
}

# Create array of arguments for ffMPEG
[array]$arrArgumentList = @(
    # Set log level to errors only
    "-loglevel", "error"
    # Overwrite existing file
    "-y",
    # Grab from GDI
    "-f", "gdigrab",
    # Set framerate
    "-framerate", "$intVideoFPS",
    # Grab entire desktop
    "-i", "desktop",
    # Run for x seconds
    "-t", "$intVideoDuration"
)

# If NVidia hardware encoding is to be used add it to the argument array. If it is not used, add desired compression ratio for default codec
if ($bolUseNVIDIAHardwareEncoding) {
    $arrArgumentList += @("-c:v", "h264_nvenc", "-qp", "0")
}
else {
    # Compression, between 0 and 51
    # Lower numbers are larger files with more CPU stress, higher numbers smaller files and less CPU stress but lower quality.
    # 28 seems fine for screenshot type videos, if you are recording a video playing on the desktop you may to lower this number.
    $arrArgumentList += @("-crf", "28")
}

# If half resolution video file is specified, add scaling
if ($bolScale) {
    $arrArgumentList += @("-vf", "scale=trunc(iw/2):trunc(ih/2)")
}

# Add the output file as last parameter
$arrArgumentList += $strFilePathAndName

# Specify amount of threads ffmpeg can use to limit CPU stress this is set to 1
$arrArgumentList += @("-threads", "1")

# Create process object for ffMPEG
$prcStartInfo = New-object System.Diagnostics.ProcessStartInfo 
$prcStartInfo.CreateNoWindow = $true 
$prcStartInfo.UseShellExecute = $false 
$prcStartInfo.RedirectStandardOutput = $true 
$prcStartInfo.RedirectStandardError = $true 
$prcStartInfo.FileName = "$strffMPEGLocation\ffmpeg.exe" 
$prcStartInfo.Arguments = $arrArgumentList

# Create the process object
$prcFFMPEG = New-Object System.Diagnostics.Process 

# Add the process start info
$prcFFMPEG.StartInfo = $prcStartInfo

# Start the process, read the error stream (ffMPEG writes to error stream only, not standard ouput) and wait for process to exit
try {
    [void]$prcFFMPEG.Start()
    $logFFMPEG = $prcFFMPEG.StandardError.ReadToEnd()
    $prcFFMPEG.WaitForExit()
}
catch {
    [string]$strRemoveFile = Remove-EmptyVideoFile -VideoFile $strFilePathAndName
    Feedback -Message "There was an error running ffMPEG.exe. $strRemoveFile" -Oops
}

# Measure CPU impact. Remove the comment on the next line to get the total processor time used by ffmpeg.exe. Use this if you want to test best optimization of ffmpeg
# Write-Host "Total CPU use time in seconds: $($prcFFMPEG.TotalProcessorTime.TotalSeconds)"

if ($logffmpeg -ne '') {
    [string]$strRemoveFile = Remove-EmptyVideoFile -VideoFile $strFilePathAndName
    Feedback -Message "There was a problem running ffMPEG.exe. $strRemoveFile`n$logFFMPEG." -Oops
}

# Test if the video file contained any output
$filVideoFile = Get-Item $strFilePathAndName
if ($filVideoFile.Length -eq 0) {
    # Something went wront, the file is empty. Try to clean up the trash
    [string]$strRemoveFile = Remove-EmptyVideoFile -VideoFile $strFilePathAndName
    Feedback -Message "A video file was created but the length was 0 bytes so ffMPEG had a problem. $strRemoveFile" -Oops
}

# We got through all that, output the result
Feedback -Message "Video file $filVideoFile was created, size $([math]::Round($($filVideoFile.Length /1kb))) KB."
