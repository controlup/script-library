#requires -Version 3.0
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    This scripts starts a recording of the Horizon Session Recording fling.

    .DESCRIPTION
    This script uses the Powershell module of the Horizon Session Recording fling to record the BLAST session of a user.

    .PARAMETER sessionId
    Id of the user session that needs to be recorded

    .PARAMETER Username
    Username of who's session to record

    .PARAMETER computer
    Name of the computer the user is working on

    .NOTES
    At least version 2.2.0 of the Horizon Session Recording fling and its agent need to be installed.

    .LINK
    https://flings.vmware.com/horizon-session-recording

#>

$sessionId=$args[0]
$username=$args[1]
$computer=$args[2]

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
      If only Message is passed this message is displayed
      If Warning is specified the message is displayed in the warning stream (Message must be included)
      If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )
    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

Function Test-ArgsCount {
    <# This function checks that the correct amount of arguments have been passed to the script. As the arguments are passed from the Console or Monitor, the reason this could be that not all the infrastructure was connected to or there is a problem retreiving the information.
      This will cause a script to fail, and in worst case scenarios the script running but using the wrong arguments.
      The possible reason for the issue is passed as the $Reason.
      Example: Test-ArgsCount -ArgsCount 3 -Reason 'The Console may not be connected to the Horizon View environment, please check this.'
      Success: no ouput
      Failure: "The script did not get enough arguments from the Console. The Console may not be connected to the Horizon View environment, please check this.", and the script will exit with error code 1
      Test-ArgsCount -ArgsCount $args -Reason 'Please check you are connectect to the XXXXX environment in the Console'
    #>    
    Param (
        [Parameter(Mandatory = $true)]
        [int]$ArgsCount,
        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    # Check all the arguments have been passed
    if ($args.Count -ne $ArgsCount) {
        Out-CUConsole -Message "The script did not get enough arguments from the Console. $Reason" -Stop
    }
}

# Test correct ammount of arguments was passed
Test-ArgsCount -ArgsCount 3 -Reason 'The Console or Monitor may not be connected to the Horizon environment or agent, please check this.'

try{
    $InstallDir = Get-ItemPropertyValue -path "hklm:\SOFTWARE\VMware, Inc.\VMware Blast\SessionRecordingAgent" -Name installdir
}
catch{
    Out-CUConsole -Message "Error determining the Horizon Session recording installation location. Please make sure the Horizon Sesison recording Agent is Installed." -Stop
}
try{
    import-module "$($InstallDir)\api\horizon.sessionrecording.powershell.dll"
}
catch{
    Out-CUConsole -Message "Error loading the Horizon Session Recording PowerShell Module. Make sure the latest version of the Horizon Session Recording Agent is installed." -Stop
}
try{
    if((Get-HSRSessions -SessionID $sessionId).isrecording -eq $true){
        Out-CUConsole -Message "There is already an active recording for $username."
        break
        }
}
Catch{
    Out-CUConsole -Message "There was an error retreiving current recordings." -Stop
}
try{
    Out-CUConsole -Message "Started session recording for $username on $computer."
    Start-HSRSessionRecording -SessionID $sessionId
}
Catch{
    Out-CUConsole -Message "There was an error starting the recording." -Stop
}

