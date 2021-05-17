<#
    .SYNOPSIS
    This script will return logging information about any ControlUp actions.

    .DESCRIPTION
    This script is a (minor) modification of David Falkus's original script for getting AppV events.  He documented everything that went into making
    this work here:  https://blogs.technet.microsoft.com/virtualshell/2016/08/25/app-v-5-troubleshooting-the-client-using-the-event-logs/

    This script takes one arguments.  
    $args[0] is how far back in time to check for events (in hours).  No argument is all events.

    AUTHOR: Trentent Tye, David Falkus
    LASTEDIT: 01/26/2017
    VERSI0N : 1.0

    modified by Ze'ev Eisenberg to better adapt to the SBA format
    
#>

# Adding threading culture change so that get-winevent picks up the messages, if PS culture is set to none en-US then the script will fail
[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"

if (!($args[0]) -or $args[0] -eq "0") {
$FilterXML = @"
<QueryList>
  <Query Id="0" Path="Application">
    <Select Path="Application">*[System[Provider[@Name='ControlUp action auditing']]]</Select>
  </Query>
</QueryList>
"@
} else {
    #convert time into milliseconds
    $time = [int]$args[0]*1000*60*60
    
$FilterXML = @"
<QueryList>
  <Query Id="0" Path="Application">
    <Select Path="Application">
    *[System[Provider[@Name='ControlUp action auditing'] 
    and TimeCreated[timediff(@SystemTime) &lt;= $($time)]]]
    </Select>
  </Query>
</QueryList>
"@
}

Try {
    $GWE_All = Get-WinEvent -FilterXml $FilterXML -ErrorAction SilentlyContinue
} Catch {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error "Unable to pull the event log" -ErrorAction Continue
    Write-Error $Error[1] -ErrorAction Continue
    Exit 1
}

#create a new object because previous events may not be defined if this is a non-persistent system
#the event contains all the data but without it defined on the server so the message property is blank
$Events = @()

If ($GWE_All -ne $null) {
    ForEach ($event in $GWE_All) {
        [xml]$eventXML = $event.ToXML()
        $prop = New-Object System.Object
        $prop | Add-Member -type NoteProperty -name TimeCreated -value ([datetime]$eventXML.Event.System.TimeCreated.SystemTime)
        $prop | Add-Member -type NoteProperty -name Message -value $eventXML.Event.EventData.Data
        $Events += $prop
    }
}

If ($Events.Count -eq "0" -or $Events -eq $null) {
    Write-Host "There are no events in the time frame requested."
    Exit
}

$Events | sort TimeCreated -Descending | select TimeCreated,Message | fl
Write-Host "$($events.count) actions in total."
