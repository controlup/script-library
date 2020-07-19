$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Check Ivanti Workspace Control DB Cache and transaction logs

    .DESCRIPTION
    Get the size of the Ivanti Workspace Control DB Cache, and if (and how many) transation logs exist

    .EXAMPLE
    Can be used to see if the local Ivanti Workspace Control agent is updating the DB and running OK.
    
    .NOTES
    Context: This script can be triggered from: Computer

Modification history: Original script by Chris Twiest
                      29062019 - Ton de Vreede - Added simple error handling, changed output method

    .LINK
#>

Function Feedback {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$Message,
        [Parameter(Mandatory = $false,
            Position = 1)]
        $Exception,
        [switch]$Oops
    )

    # This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If (!$Exception -and !$Oops) {
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

# Get the DB location
try {
    [string]$Key = 'HKLM:\SOFTWARE\WOW6432Node\RES\Workspace Manager'
    [string]$dbcachepath = (Get-ItemProperty -Path $key -Name LocalCachePath).LocalCachePath
}
catch {
    Feedback -Message 'The location of the Workspace Control database could not be retreived from the registry. Using default location.'
    $dbcachepath = "$ENV:RESPFDIR\Data\DBcache"
}

# Get Database size and transactions
try {
    $size = "{0:N2} MB" -f ((Get-ChildItem $dbcachepath -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB)
    $transactions = Get-ChildItem "$dbcachepath\Transactions"
}
catch {
    Feedback -Message 'The size and content of the Workspace Control database folder could not be read.' -Exception $_
}

# Output the results
Feedback -Message "The Ivanti Workspace Control DB cache folder located at $dbcachepath is $size"

if ($transactions -eq $null) {
    Feedback -Message "The transactions folder is empty"
}
else {
    Feedback -Message "The transactions folder is not empty and contains the following files:`n$transactions.name"
}
