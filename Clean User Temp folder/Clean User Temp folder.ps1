$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Deletes all files and folders in user TEMP folder.

    .DESCRIPTION
    The script deletes all files in the user's TEMP folder only

    .EXAMPLE
    - The user's TEMP folder needs to be cleaned because it has become too large
    - A program has left data in the TEMP folder that causes it to malfunction
    
    .NOTES
    Some programs may place files in the TEMP folder but not lock them. These files will be deleted as well, which may cause a program to malfunction.
    It is preferred to run this script only when a user has closed all applications.
    
    Based on a community script submission by Andy Gresbach

    Context: User session
    Modification history: 29062019 - Ton de Vreede - Created PS code instead of batch script, added error handling and some feedback
    
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

# Delete all files and folders in the user TEMP folder. Using -ErrorAction SilentlyContinue as there will almost always be files that can't be deleted because they are locked.
try {
    $TempContents = Get-ChildItem "$([System.Environment]::ExpandEnvironmentVariables('%TEMP%'))\" -Recurse
    Foreach ($item in $TempContents) {
        Remove-Item -Path $Item.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}
catch {
    Feedback -Message 'There was a problem deleting all the folders and files in the user TEMP folder.' -Exception $_
}

Feedback -Message 'Folders and files in the user TEMP folder have been deleted. Some folders and files that were locked or the user does not have NTFS access to may remain.'
