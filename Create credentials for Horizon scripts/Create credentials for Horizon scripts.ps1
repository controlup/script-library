#requires -Version 3.0
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    Prepares the a PSCredential object on the target device for running ControlUp Horizon View scripts

    .DESCRIPTION
    This script creates an encrypted PSCredential object on the target machine in order to allow running of script for Horizon View without having to authenticate manually.

    .EXAMPLE
    This script should be run on any machines that will run Horizon View scripts. In general these are the machines that run the ControlUp Console or ControlUp Monitors. 
    
    .NOTES
    Connecting to a Horizon View Connection server is required for running Horizon View scripts. The server does not allow passthrough (Active Directory) authentication. In order to allow scripts to run without asking for a password each time (such as in Automated Actions) a PSCredential
    object needs to be stored on each target device (ie. each machine that will be used for running Horizon View scripts). This script can create this PSCredential object on the targets.
    PSCREDENTIAL OBJECTS CAN ONLY BE USED BY THE USER THAT CREATED THE OBJECT AND ON THE MACHINE THE OBJECT WAS CREATED.
    - The User that creates the file is required to have a local profile when creating the file. This is a limitation from Powershell
    
    Modification history:   20/08/2019 - Anthonie de Vreede - First version
                            03/06/2020 - Wouter Kursten - Second version
                            10/09/2020 - WOuter Kursten - Third Version

    Changelog ;
        Second Version
            - Added check for local profile
            - changed error message when failing to create the xml file
            - Fixed issue where users without local admin rights and no active session on the target machine couldn't create a credentrials file ($env:USERPROFILE returns c:\users\default)

    .PARAMETER strUsername
    The username for the PSCredential object
    
    .PARAMETER strPassword1
    The password for the credential object

    .PARAMETER strPassword2
    Double check the password
    
    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    https://github.com/vmware/PowerCLI-Example-Scripts/tree/master/Modules/VMware.Hv.Helper
#>

[string]$strUserName = $args[0]
[string]$strPassword1 = $args[1]
[string]$strPassword2 = $args[2]

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
            Write-Error "$Message`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)`n" -ErrorAction Stop
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

Function New-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The username to be stored in the PSCredential object.")]
        [string]$Username,
        [parameter(Mandatory = $true,
            HelpMessage = "The password to be stored in the PSCredential object.")]
        [string]$Password,
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )
    # Username and password correct, check if target folder exists and create it if necessary
    $strCredTargetFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

    # Create the folder if it does not exist
    If (!(Test-Path $strCredTargetFolder)) {
        Write-Host "Folder does not exist"
        try {
            New-Item -Path $strCredTargetFolder -ItemType Directory
        }
        catch {
            Out-CUConsole -Message "There was a problem creating the folder used to store the credentials object ($strCredTargetFolder). Please make sure you have permission to write to the parent folder." -Exception $_
        }
    }

    # Create the PSCredential object
    try {
        [System.Security.SecureString]$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePassword)
    }
    catch {
        Out-CUConsole -Message "There was a problem creating the PSCredential object." -Exception $_
    }
    $credsfile = $strCredTargetFolder+"\"+$($Env:Username)+"_"+$($System)+"_Cred.xml"
    # Store the PSCredential object
    try {

        $Cred | Export-CliXml -Path $credsfile -Force  
        Out-CUConsole -Message "PSCredential object created and stored in $strCredTargetFolder." 
    }
    catch {
        Remove-Item -path $credsfile -force
        Out-CUConsole -Message "There was a problem saving the PSCredential object to $strCredTargetFolder. This may be a permission issue or there is no local profile." -Exception $_
    }
}

$userprofile = $env:USERPROFILE

if ($userprofile -like "*default*"){
    if(!(Get-WmiObject -class win32_userprofile | where-object {$_.localpath -like "$env:systemdrive\Users\$env:username"})){
        Out-CUConsole -message "User $Env:Username has no profile on this system. This is a requirement for creating the credentials file. Please log onto this machine once in order to create your user profile."  -exception "No local profile found" # this is a limitation of Powershell
    }
}



# Check if both username and password were filled in and passwords match
If (!([string]::IsNullOrWhiteSpace($strUserName)) -and !([string]::IsNullOrWhiteSpace($strPassword1)) -and ($strPassword1 -eq $strPassword2)) {
    New-CUStoredCredential -Username $strUserName -Password $strPassword1 -System 'HorizonView'
}
else {
    If ($strPassword1 -ne $strPassword2) {
        Out-CUConsole -Message "The passwords do not match. Please enter the same password in both password fields." -Stop
    }
}

