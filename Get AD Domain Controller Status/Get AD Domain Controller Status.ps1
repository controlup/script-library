<#
.SYNOPSIS
    List the synchronization status and replication errors of all domain controllers in the domain.

.DESCRIPTION    
    When this script is executed it will retrieve the synchronization status of all domain controllers within the domain.
    The domain controllers will also be queried for replication errors to help with the primary troubleshooting of synchronization issues.

    This script can be executed on a ControlUp monitor and will request the required data via a PSSession to the domain controller(s).

.EXAMPLE
    This script is intended to be used within ControlUp as an action. No parameters are required.

.NOTES
    Author: Rein Leen
    Contributor(s): Gillian Stravers
    Context: Machine
    Modification_history:
        Rein Leen       23-05-2023      Version ready for release
#>

#region [parameters]
[CmdletBinding()]
Param ()
#endregion [parameters]

#region [variables]
# Required dependencies
#Requires -Version 5.1
#Requires -Modules ActiveDirectory

# Import modules (required in .NET engine)
Import-Module -Name ActiveDirectory

# Setting error actions
$ErrorActionPreference = 'Stop'
$DebugPreference = 'SilentlyContinue'

# Scriptblock
$domainControllerReplicationScriptBlock = {
    $ErrorActionPreference = 'Stop'
    $returnDataHashtable = @{}
    # Get Replication Status
    $domainControllers = Get-AdDomainController -Filter *
    foreach ($domainController in $domainControllers) {
        Try {
            $returnDataHashtable[('{0}-ReplicationPartnerMetadata' -f $domainController)] = Get-AdReplicationPartnerMetadata -Target $domainController.HostName -ErrorAction Stop | Select-Object Server, LastReplicationAttempt, LastReplicationSuccess, @{ Name = 'LastReplicationResult'
            Expression = {if ($_.LastReplicationResult -eq 0) {$true} else {$false}}}
        } Catch {
            # Ignore errors, completeness of data is handled outside the scriptblock.
        }
        $returnDataHashtable[('{0}-ReplicationErrors' -f $domainController)] = repadmin /showrepl $domainController.HostName /errorsonly
    }
    return $returnDataHashtable
}
#endregion [variables]

#region [actions]
# Get Replication status from the first domain controller that returns all data.
$domainControllers = Get-AdDomainController -Filter *
foreach ($domainController in $domainControllers) {
    Write-Verbose ('Retrieving data from {0}' -f $domainController.HostName)
    # Setup and enter PSSession
    Try {
        # Set session options in case of self-signed certificates.
        # ErrorAction is set to Stop to allow fallback to non-SSL option.
        $sessionOptions = New-PSSessionOption -SkipCACheck -OpenTimeout 2000
        $session = New-PSSession -ComputerName $domainController.HostName -UseSSL -SessionOption $sessionOptions -ErrorAction Stop
    } Catch {
        $session = New-PSSession -ComputerName $domainController.HostName
    }

    # Try the next domain controller if no session has been set up.
    if ([string]::IsNullOrWhiteSpace($session)) {
        Write-Verbose ('Failed to retrieve data from {0}' -f $domainController.HostName)
        continue
    }

    # Run the scriptblock if a session has been set up.
    $invokeCommandResponse = Invoke-Command -Session $session -ScriptBlock $domainControllerReplicationScriptBlock
    Write-Verbose ('Finished retrieving data from {0}' -f $domainController.HostName)
    # Break the loop since no further domain controllers have to be queried.
    break
}

# Remove all sessions
Get-PSSession | Remove-PSSession

# Output data
Write-Output ('')('*' * 80)('Replication Partner Metadata:')
foreach ($domainController in $domainControllers.Name) {
    if ([string]::IsNullOrWhiteSpace($invokeCommandResponse[('{0}-ReplicationPartnerMetadata' -f $domainController)])) {
        Write-Warning ('Replication Partner Metadata for {0} is incomplete.' -f $domainController)
    } else {
        Write-Output $invokeCommandResponse[('{0}-ReplicationPartnerMetadata' -f $domainController)]
    }
}

Write-Output ('')('*' * 80)('Replication Errors:')('')
foreach ($domainController in $domainControllers.Name) {    
    if ([string]::IsNullOrWhiteSpace($invokeCommandResponse[('{0}-ReplicationErrors' -f $domainController)])) {
        Write-Warning ('Replication Errors data for {0} is incomplete.' -f $domainController)
    } else {
        # Only output error lines with a hex-code
        [int]$count = 0
        foreach ($line in $invokeCommandResponse[('{0}-ReplicationErrors' -f $domainController)] -match '^.*\d*\(0x\d*\):.*$') {
            ('{0}: {1}' -f $domainController, $line.Trim())
            $count++
        }
        if ($count -eq 0) {
            Write-Output ('{0}: No replication errors found' -f $domainController)
        }
    }
}
#endregion [actions]
