<#
.SYNOPSIS
   Get Load Index for a server in the site.
.PARAMETER Identity
   The name of the server you want to get the Load Index of - automatically supplied by CU
.NOTES
    Requires PowerShell 3.0
#>

if (!(Get-PSSnapin -Name Citrix.Broker.Admin.*)) {
    try {
        Add-PsSnapin Citrix.Broker.Admin.*
    } 
catch {
        # capture any failure and display it in the error section, then end the script with a return
        # code of 1 so that CU sees that it was not successful.
        Write-Error "Unable to load the snapin" -ErrorAction Continue
        Write-Error $Error[1] -ErrorAction Continue
        Exit 1
    }
}

# machineName has to be supplied from CU in the form of Computer FQDN.
$machineName = $args[0]

try {
    $Machine = Get-BrokerMachine -MachineName "*\$machineName"
    if ($Machine.gettype()) {
        $LIArray = $Machine.LoadIndexes
    }
}
catch {
    Throw "This site has no machine named $machineName, Probably not a XenDesktop machine."
    Exit 1
}

$LIObject = New-Object -TypeName psobject
$LIObject | Add-Member -MemberType NoteProperty -Name LoadIndex -Value $Machine.LoadIndex

foreach ($CurrentLoadIndex in $LIArray) {
    if ($CurrentLoadIndex -match "CPU") {
        $LIObject | Add-Member -MemberType NoteProperty -Name CPU -Value ($CurrentLoadIndex).Substring(4)
    }
    elseif ($CurrentLoadIndex -match "Memory") {
        $LIObject | Add-Member -MemberType NoteProperty -Name Memory -Value ($CurrentLoadIndex).Substring(7)
    }
    elseif ($CurrentLoadIndex -match "Disk") {
        $LIObject | Add-Member -MemberType NoteProperty -Name Disk -Value ($CurrentLoadIndex).Substring(5)
    }
    elseif ($CurrentLoadIndex -match "Session") {
        $LIObject | Add-Member -MemberType NoteProperty -Name "User Sessions" -Value ($CurrentLoadIndex).Substring(15)
    }
}
Write-Host "Load Index Summary`n------------------" -NoNewline
$LIObject | Format-List
$percent = ($LIObject.LoadIndex/10000)
"Utilization of $machineName is {0:P0}" -f $percent
