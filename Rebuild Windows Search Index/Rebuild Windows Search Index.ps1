$ErrorActionPreference = 'Stop'
<#
    .SYNOPSIS
    Deletes the Windows Search Index file in order to rebuild it.

    .DESCRIPTION
    The file can only be deleted if the Windows Search service has been stopped. It also has a tendency to restart too fast, so this can cause errors while deleting the Index file. So the service is first Disabled and stopped. After deleting the Index file
    the service is set back to Automatic start. No attempt is made to restart the service as this often causes errors if done straight away. Due to the (default) Automatic(DelayedStart) setting the Search service will start by itself when it is good and ready.

    .EXAMPLE
    This script can be used as a quick fix if corruption of the Search Index file is suspected.
    
    .NOTES
    No backup of the Index file is made, as this will be rebuild automatically anyway. The target machine will get some extra CPU and Disk activity as the Index file is rebuild
    
    Context: Computer
    Modification history: 29062019 - Ton de Vreede - Created PS code instead of batch script, added error handling and some feedback. Removed backup of Index file as this will rarely be useful and could leave a large unused file on the system.
    
    Script inspired by the community script submission by Tas Smith
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

# Reconfigure Windows Search Service so that it will not restart straight away. Then stop the service.
try {
    Set-Service -Name 'wsearch' -StartupType Disabled
    Stop-Service -Name 'wsearch' -Force
}
catch {
    Feedback -Message 'There was a problem reconfiguring and stopping the Windows Search Sevice.' -Exception $_
}

# Delete the index file
try {
    Remove-Item -Path "$([System.Environment]::ExpandEnvironmentVariables("%programdata%\microsoft\search\data\applications\windows\Windows.edb"))" -Force
}
catch {
    # If it fails, try to restart the service and inform the user
    try {
        Set-Service -Name 'wsearch' -StartupType Automatic
    }
    catch {
        Feedback -Message 'The Windows Search Index file could not be deleted. The Search service could not be set to start Automtically and restarted. Please check the machine to ensure the Search Service is configured as it should be again.' -Exception $_
    }
    Feedback -Message 'The Windows Search Index file could not be deleted. The service has been set back to Automatic start and will start again in 2 minutes.' -Exception $_
}

# Set the Windows Search Service startup type back to Automatic.
try {
    Set-Service -Name 'wsearch' -StartupType Automatic
}
catch {

    Feedback -Message 'The Windows Search Index file has been deleted but the Windows Search service could not be set to start automtically and restarted. Please check the machine to ensure the Search Service is configured as it should be again.' -Exception $_
}

Feedback -Message "The Windows Search index file has been deleted. The service has been set back to Automtically start and should start in 2 minutes."
