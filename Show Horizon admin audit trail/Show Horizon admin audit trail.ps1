<#
    .SYNOPSIS
    Pulls all admin events from the Horizon Event Database

    .DESCRIPTION
    Uses the Horizon REST api's to pull all admin related events from the Horizon Event database for all pods. If there is no cloud pod setup it will only process the local pod.
    After pulling the events it will translate the id's for the various objects to names to show the proper names where needed.

    .PARAMETER ConnectionserverFQDN
    Passes as the Primary Connection server object from a machine

    .PARAMETER daysback
    Amount of days to go back to gather the audit logs

    .PARAMETER csvlocation
    Path to where the output csv file will be stored ending in \.

    .PARAMETER csvfilename
    Name of the csv file where the output will be stored.


    .NOTES

    Modification history:   06/05/2020 - Wouter Kursten - First version
                            10/09/2020 - Wouter Kursten - second version
                            10/19/2023 - Wouter Kursten - Third Version
    Changelog -
            10/09/2020 - Added check for export folder
            10/09/2020 - Added check if the user is able to write to te logfile
            10/09/2020 - Added check for ending \ in SBA
            10/19/2023 - Converted to use REST API and other changes

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = 'FQDN of the connectionserver' )]
    [string]$ConnectionServerFQDN,

    [Parameter(Mandatory = $true, HelpMessage = 'Amount of hours to look back for events' )]
    [int]$daysback,

    [Parameter(Mandatory = $true, HelpMessage = 'folder to export csv to' )]
    [string]$csvlocation,

    [Parameter(Mandatory = $true, HelpMessage = 'filename to export csv to' )]
    [string]$csvfilename
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Get-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )

    # Get the stored credential object
    $strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    try {
        Import-Clixml -LiteralPath $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
    }
    catch {
        write-error $_
    }
}

function Get-HRHeader() {
    param($accessToken)
    return @{
        'Authorization' = 'Bearer ' + $($accessToken.access_token)
        'Content-Type'  = "application/json"
    }
}

function Open-HRConnection() {
    param(
        [string] $username,
        [string] $password,
        [string] $domain,
        [string] $url
    )

    $Credentials = New-Object psobject -Property @{
        username = $username
        password = $password
        domain   = $domain
    }

    return invoke-restmethod -Method Post -uri "$url/rest/login" -ContentType "application/json" -Body ($Credentials | ConvertTo-Json)
}

function Close-HRConnection() {
    param(
        $refreshToken,
        $url
    )
    return Invoke-RestMethod -Method post -uri "$url/rest/logout" -ContentType "application/json" -Body ($refreshToken | ConvertTo-Json)
}

function Get-HorizonRestData() {
    [CmdletBinding(DefaultParametersetName = 'None')] 
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'url to the server i.e. https://pod1cbr1.loft.lab' )]
        [string] $ServerURL,

        [Parameter(Mandatory = $true,
            ParameterSetName = "filteringandpagination",
            HelpMessage = 'Array of ordered hashtables' )]
        [array] $filters,

        [Parameter(Mandatory = $true,
            ParameterSetName = "filteringandpagination",
            HelpMessage = 'Type of filter Options: And, Or' )]
        [ValidateSet('And', 'Or')]
        [string] $Filtertype,

        [Parameter(Mandatory = $false,
            ParameterSetName = "filteringandpagination",
            HelpMessage = 'Page size, default = 500' )]
        [int] $pagesize = 1000,

        [Parameter(Mandatory = $true,
            HelpMessage = 'Part after the url in the swagger UI i.e. /external/v1/ad-users-or-groups' )]
        [string] $RestMethod,

        [Parameter(Mandatory = $true,
            HelpMessage = 'Part after the url in the swagger UI i.e. /external/v1/ad-users-or-groups' )]
        [PSCustomObject] $accessToken,

        [Parameter(Mandatory = $false,
            ParameterSetName = "filteringandpagination",
            HelpMessage = '$True for rest methods that contain pagination and filtering, default = False' )]
        [switch] $filteringandpagination,

        [Parameter(Mandatory = $false,
            ParameterSetName = "id",
            HelpMessage = 'To be used with single id based queries like /monitor/v1/connection-servers/{id}' )]
        [string] $id,

        [Parameter(Mandatory = $false,
            HelpMessage = 'Extra additions to the query url that comes before the paging/filtering parts like brokering_pod_id=806ca in /rest/inventory/v1/global-sessions?brokering_pod_id=806ca&page=2&size=100' )]
        [string] $urldetails
    )

    if ($filteringandpagination) {
        if ($filters) {
            $filterhashtable = [ordered]@{}
            $filterhashtable.add('type', $filtertype)
            $filterhashtable.filters = @()
            foreach ($filter in $filters) {
                $filterhashtable.filters += $filter
            }
            $filterflat = $filterhashtable | convertto-json -Compress
            if ($urldetails) {
                $urlstart = $ServerURL + "/rest/" + $RestMethod + "?" + $urldetails + "&filter=" + $filterflat + "&page="
            }
            else {
                $urlstart = $ServerURL + "/rest/" + $RestMethod + "?filter=" + $filterflat + "&page="
            }
        }
        else {
            if ($urldetails) {
                $urlstart = $ServerURL + "/rest/" + $RestMethod + "?" + $urldetails + "&page="
            }
            else {
                $urlstart = $ServerURL + "/rest/" + $RestMethod + "?page="
            }
        }
        $results = [System.Collections.ArrayList]@()
        $page = 1
        $uri = $urlstart + $page + "&size=$pagesize"
        $response = Invoke-webrequest $uri -Method 'GET' -Headers (Get-HRHeader -accessToken $accessToken)
        $data = $response.content | convertfrom-json
        $responseheader = $response.headers
        $data.foreach({ $results.add($_) }) | out-null
        if ($responseheader.HAS_MORE_RECORDS -contains "TRUE") {
            do {
                $page++
                $uri = $urlstart + $page + "&size=$pagesize"
                $response = Invoke-webrequest $uri -Method 'GET' -Headers (Get-HRHeader -accessToken $accessToken) 
                $data = $response.content | convertfrom-json
                $responseheader = $response.headers
                $data.foreach({ $results.add($_) }) | out-null
            } until ($responseheader.HAS_MORE_RECORDS -notcontains "TRUE")
        }
    }
    elseif ($id) {
        $uri = $ServerURL + "/rest/" + $RestMethod + "/" + $id
        $data = Invoke-webrequest $uri -Method 'GET' -Headers (Get-HRHeader -accessToken $accessToken)
        $results = $data.content | convertfrom-json
    }
    else {
        if ($urldetails) {
            $uri = $ServerURL + "/rest/" + $RestMethod + "?" + $urldetails
        }
        else {
            $uri = $ServerURL + "/rest/" + $RestMethod
        }
        $data = Invoke-webrequest $uri -Method 'GET' -Headers (Get-HRHeader -accessToken $accessToken)
        $results = $data.content | convertfrom-json
    }

    return $results
}

function Get-HorizonCloudPodStatus{
    param(
        $accesstoken,
        $ConnectionServerURL
    )
    write-verbose "Getting Horizon Cloud Pod Status"
    return  (Get-HorizonRestData -ServerURL $ConnectionServerURL -accessToken $accesstoken -RestMethod "federation/v1/cpa").local_connection_server_status
}

function Get-HorizonCloudPodDetails{
    param(
        $accesstoken,
        $ConnectionServerURL
    )
    write-verbose "Getting Horizon Cloud Pod details"
    $CloudPodDetails = @()
    $pods = Get-HorizonRestData -ServerURL $ConnectionServerURL -accessToken $accesstoken -RestMethod "federation/v1/pods"
    foreach ($pod in $pods){
        $PodName = $pod.name
        write-verbose "Getting details for $podname"
        $restmethod = "federation/v1/pods/"+$pod.id+"/endpoints"
        $PodDetails = Get-HorizonRestData -ServerURL $ConnectionServerURL -accessToken $accesstoken -RestMethod $restmethod
        foreach ($PodDetail in $PodDetails){
            $name = $PodDetail.name 
            write-verbose "Creating object with details for $name"
            $server_address = $PodDetail.server_address
            $ServerDNS = ($server_address.Replace("https://","")).split(":")[0]
            $DetailObject = [PSCustomObject]@{
                PodName     = $PodName
                Name = $name
                ServerDNS    = $ServerDNS
            }
            $CloudPodDetails+=$DetailObject
        }
    }
    return $CloudPodDetails
}

function Get-HorizonNonCloudPodDetails{
    param(
        $accesstoken,
        $ConnectionServerURL,
        $ServerName
    )
    write-verbose "Getting Horizon Non Cloud Pod details"
    $EnvDetails = Get-HorizonRestData -ServerURL $ConnectionServerURL -accessToken $accesstoken -RestMethod "config/v2/environment-properties"
    $ConDetails = Get-HorizonRestData -ServerURL $ConnectionServerURL -accessToken $accesstoken -RestMethod "monitor/v2/connection-servers"
    $PodName = $EnvDetails.cluster_name
    write-verbose "Getting details for $podname"
    $PodDetails = @()
    foreach ($ConDetail in $ConDetails){
        $name = $ConDetail.name
        write-verbose "Creating object with details for $name"
        $ServerDNS = $servername.replace(($servername.split(".")[0]),$name)
        write-verbose "Server DNS should be $ServerDNS"
        $DetailObject = [PSCustomObject]@{
            PodName     = $PodName
            Name = $name
            ServerDNS    = $ServerDNS
        }
            $PodDetails+=$DetailObject
    }
    return $PodDetails
}

function connect-hvpod{
    param($PodName)
    write-verbose "trying to find a working connection server for $podname"
    [array]$conservers = $ServerArray | Where-Object {$_.PodName -eq "$podname"} | Select-Object -expandproperty ServerDNS
    $count = $conservers.count
	write-verbose $count
    $start=0
    $result= "None"
    do{
        write-verbose "Attempt: $start"
        $conserver = $conservers[$start]
        $serverurl = "https://$conserver"
        if((Test-netConnection $conserver -port 443 -InformationLevel quiet) -eq $true){
            $login_tokens = Open-HRConnection -username $username -password $UnsecurePassword -domain $Domain -url $serverurl
        }
        else{
            $login_tokens = $null
        }
        if($null -ne $login_tokens){
            $result= "Success"
            $Script:LastSelectedConnectionServer = $conserver
            write-verbose "Connected using $LastSelectedConnectionServer"
            return $login_tokens
        }
        else{
            $result= "None"
            $start++
            if($start -lt $count){
                write-verbose "Failed connecting to $conserver, trying the next one."
            }
            else{
                write-verbose "Cannot find a working connection server for $podname"
                $result= "Failure"
                return $result
            }
        }

    }
    until(($result -eq "Success") -or ($result -eq "Failure"))
}
# we need the following to ignore invalid certificates

Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
    ServicePoint srvPoint, X509Certificate certificate,
    WebRequest request, int certificateProblem) {
        return true;
    }
}
"@

[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

# Set Tls versions
$allProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $allProtocols

# Get the stored credentials for running the script

[PSCredential]$creds = Get-CUStoredCredential -System 'HorizonView'
$date = get-date
$thentime = (get-date).adddays(-$daysback)

$nowepoch = ([DateTimeOffset]$date).ToUnixTimeMilliseconds()
$thenepoch = ([DateTimeOffset]$thentime).ToUnixTimeMilliseconds()

$username = ($creds.GetNetworkCredential()).userName
$domain = ($creds.GetNetworkCredential()).Domain
$UnsecurePassword = ($creds.GetNetworkCredential()).password

$url = "https://$ConnectionServerFQDN"
$login_tokens = Open-HRConnection -username $username -password $UnsecurePassword -domain $Domain -url $url
$AccessToken = $login_tokens | Select-Object Access_token
$RefreshToken = $login_tokens | Select-Object Refresh_token

$CloudPodStatus = Get-HorizonCloudPodStatus -ConnectionServerURL $url -accesstoken $AccessToken
write-output "Cloud Pod Status: $cloudPodStatus"
if($CloudPodStatus -eq "ENABLED"){
    $CloudPodDetails = Get-HorizonCloudPodDetails -ConnectionServerURL $url -accesstoken $AccessToken
    [array]$ServerArray = $CloudPodDetails
}
else{
    $PodDetails = Get-HorizonNonCloudPodDetails -ConnectionServerURL $url -accesstoken $AccessToken -ServerName $connectionserveraddress
    [array]$ServerArray = $PodDetails
}
Close-HRConnection -refreshToken $RefreshToken -url $url
$auditevents = New-Object System.Collections.ArrayList

foreach($podname in ($serverarray | select-object -expandproperty Podname -unique)){
    write-output "Getting data for $podname"
    try{
        $login_tokens = connect-hvpod -podname $podname
        if($login_tokens -eq "Failure"){
            write-output "failed to connect so can't apply secondary image."
            break
        }
        $AccessToken = $login_tokens | Select-Object Access_token
        $RefreshToken = $login_tokens | Select-Object Refresh_token
        $serverURL = "https://$LastSelectedConnectionServer"
    }
    catch{
        write-output "Error Connecting."
        break
    }
    $rawauditevents = @()
    $auditfilters = @()
    $timefilter = [ordered]@{}
    $timefilter.add('type', 'Between')
    $timefilter.add('name', 'time')
    $timefilter.add('fromValue', $thenepoch)
    $timefilter.add('toValue', $nowepoch)
    $auditfilters += $timefilter
    $auditfilter = [ordered]@{}
    $auditfilter.add('type', 'Equals')
    $auditfilter.add('name', 'module')
    $auditfilter.add('value', 'Vlsi')
    $auditfilters += $auditfilter
    $rawauditevents += Get-HorizonRestData -ServerURL $serverURL -RestMethod "/external/v1/audit-events" -accessToken $accessToken -filteringandpagination -Filtertype "And" -filters $auditfilters
    $count = $rawauditevents.count
    write-verbose "Found $count events for $podname"
    foreach ($event in $rawauditevents) {
        $readabletimestamp = ([datetimeoffset]::FromUnixTimeMilliseconds(($event).Time)).ToLocalTime()
        $readabletimestamputc = [datetimeoffset]::FromUnixTimeMilliseconds(($event).Time)
        $timeepoch = ($event).time
        $event.psobject.Properties.Remove('Time')
        $event | Add-Member -MemberType NoteProperty -Name Pod -Value $podname
        $event | Add-Member -MemberType NoteProperty -Name time -Value $readabletimestamp
        $event | Add-Member -MemberType NoteProperty -Name time_utc -Value $readabletimestamputc
        $event | Add-Member -MemberType NoteProperty -Name time_epoch -Value $timeepoch
        $auditevents.add($event) | out-null
    }

    Close-HRConnection -refreshToken $RefreshToken -url $serverURL
}
# checks if this connect

$auditevents | sort-object -property podname, time_epoch | format-table time,machine_dns_name,message -autosize -wrap

if ($csvlocation -eq "%temp%") {
    $csvlocation = $env:TEMP + "\" 
}
elseif (!($csvlocation -match '\\$')) {
    $csvlocation = $csvlocation + "\"
}

if (test-path $csvlocation) {
    $csvpath = $csvlocation + $csvfilename
}
else {
    write-output "Unable to find folder $csvlocation reverting to default $strCUCredFolder"
    $csvpath = $strCUCredFolder + "\" + $csvfilename
}
Try { [io.file]::OpenWrite($csvpath).close() }
Catch {
    $csvpathnew = $strCUCredFolder + "\" + $csvfilename
    write-output "Unable to write to output file $csvpath reverting to default folder $csvpathnew"
    $csvpath = $csvpathnew
}
write-output "Writing Logfile"
$auditevents | sort-object -property Podname, time_epoch | ConvertTo-Csv -NoTypeInformation | Out-File $csvpath -Encoding utf8 -force
write-output "Logfile written to $csvpath"

