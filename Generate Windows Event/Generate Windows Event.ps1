#requires -Version 3.0
<#
.SYNOPSIS     Generates a windows event

.DESCRIPTION  Uses write-eventlog to create a windows event that can be used to test triggers.

.CONTEXT      Machine

.TAGS         $Machine, $event, $Trigger

.HISTORY      31-03-2021 -Wouter Kursten - First Version
#>

# Name of the event log to write to.
[string]$logname = $args[0]
# type of the event for ControlUp this can be Error,Warning or FailureAudit
[string]$EventType = $args[1]
# The EventID
[int]$EventID = $args[2]
# The message to write to the item
[string]$Message = $args[3]
# Source, default is "ControlUp Agent"
[string]$Source = $args[4]

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
        write-error "The script did not get enough arguments from the Console. $Reason"
    }
}

Test-ArgsCount -ArgsCount 5 -Reason 'Not all arguments had a value, please check the arguments.'

write-eventlog -LogName $logname -EntryType $EventType -EventId $EventID -Message $Message -Source $Source
