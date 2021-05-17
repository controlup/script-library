<#
   .NAME:      Check Citrix User Device license checkout
   .AUTHOR: Marcel Calef & Dennis Geerlings
       non-parametrizied command:
        & 'C:\Program Files (x86)\Citrix\Licensing\ls\udadmin.exe' -list -times
   .Documentation:
https://docs.citrix.com/en-us/licensing/current-release/admin-no-console/license-administration-commands.html#display-or-release-licenses-for-users-or-devices-udadmin
#>

#Inputs
$start = Get-Date
$mem = $args[0]                     # to verify the agent is isntalled on the session host
$input= $args[1]    ;   $domain,$username = $input.split('\')
$clientname      = $args[2]
$receiverVersion = $args[3]   # to verify this is a connected Citrix Session


# Verify the Agent is running in the session and not just the Citrix Site reporting on it
if ($mem.length -lt 1) {Write-Host "The ControlUp Agent must be installed on the session host" ; exit}
if ($receiverVersion.length -lt 1) {Write-host "Not a Citrix Session" ; exit}

Write-Output "User is $username and Clientname is $clientname "  

$output = (& ("${env:ProgramFiles(x86)}" + '\Citrix\Licensing\LS\udadmin.exe') -list -times)
$results = New-Object -TypeName System.Collections.ArrayList
$referenceObject = New-Object -TypeName psobject
$referenceObject | Add-Member -MemberType NoteProperty -Name Consumer -Value 0
$referenceObject | Add-Member -MemberType NoteProperty -Name CheckOut_date -Value 0
$referenceObject | Add-Member -MemberType NoteProperty -Name LicenseName -Value 0
$referenceObject | Add-Member -MemberType NoteProperty -Name LicenseExp -Value 0

foreach($line in $output)
{
    $currentObject = $referenceObject.psobject.Copy()
    $dateStart = $line.LastIndexOf("(")

    # If the line does not contain a valid Date ignore it
    if($dateStart -eq -1)   { continue }

    #Parse the relevant data into a table
    $currentObject.CheckOut_date = $line.substring($dateStart,$line.Length-$dateStart)
    $currentObject.LicenseName,$currentObject.licenseExp = $line.Replace($currentObject.CheckOut_date,"").trim().split(" ")[-1..-2]
    $consumer = $line.Replace($currentObject.CheckOut_date,"").replace($currentObject.LicenseName,"").replace($currentObject.licenseExp,"").trim()
    $currentObject.Consumer = [regex]::Replace($consumer,"\(.*\)","")    
    

    [void]$results.add($currentObject)
}

#PRint out the results
write-host $output[0]
$consumedLic = $results |where-object {$_.consumer -eq $username -or $_.consumer-eq $clientname}
$end = Get-Date
write-debug "script took: $(($end - $start).totalseconds)"
Write-debug $output.Count

$consumedLic | Format-Table
if ($consumedLic.count -gt 1) {write-host "User $username on $computer $clientname consumsed $($consumedLic.count) licenses"}
