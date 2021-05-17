<#
    .SYNOPSIS
        Runs a command against an IGEL device

    .DESCRIPTION
        Instruct the IGEL UMS to reboot the target device using the IGEL IMI Rest API. 
		Note: Given that the common practice for the UMS server is to use the self signed Certificate, 
				this script will ignore SSL errors
 
    .CONTEXT
        Session

    .MODIFICATION_HISTORY
        Created  TTYE : 2019-12-12
        Modified Marcel Calef  : 2020-02-25


    AUTHOR: Trentent Tye
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='IGEL Client name of a session')][ValidateNotNullOrEmpty()]     [string]$Device,
    [Parameter(Mandatory=$true,HelpMessage='FQDN of the IGEL UMS server')][ValidateNotNullOrEmpty()]       [string]$UMSServer,
    [Parameter(Mandatory=$true,HelpMessage='Command to execute')][ValidateSet('Reboot')]                   [string]$Command
    
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
###$VerbosePreference = "continue" # better invoked by passing the Debug parameter:  -verbose

function Invoke-IGELRestAPI {
    Param
    (
        [Parameter(Mandatory=$true)] [string]$URI,
        [Parameter(Mandatory=$true)] [hashtable]$Headers,
        [Parameter(Mandatory=$true)][ValidateSet("GET","POST","PUT","DELETE")] [string]$Method,
        [Parameter(Mandatory=$false)] $Body,
        [Parameter(Mandatory=$false)] [string]$ContentType,
        [Parameter(Mandatory=$false)] [switch]$Session
    )

    if ([bool]($body -as [xml])) {
        $body = [xml]$body
    }
    

    if ($body) {
        try {
        if ($session) {
            Write-Verbose "Executing REST API with a body and a new session variable"
            Invoke-WebRequest -Uri $URI -Method $method -Headers $headers -SessionVariable script:session -ContentType $ContentType -Body $body -UseBasicParsing -OutVariable webResult | Out-Null
            } else {
            Write-Verbose "Executing REST API with a body and a webSession variable"
            Invoke-WebRequest -Uri $URI -Method $method -Headers $headers -WebSession $script:session -ContentType $ContentType -Body $body -UseBasicParsing -OutVariable webResult  | Out-Null
            }
        } catch {
            $Failure = $_.Exception.Response
            return $Failure
        }
        Write-Verbose "Return Result with Body sent"
        return $webResult
    } else {
        try {
        if ($session) {
            Write-Verbose "Executing REST API and a new session variable"
            Invoke-WebRequest -Uri $URI -Method $method -Headers $headers -SessionVariable script:session -UseBasicParsing -OutVariable webResult  | Out-Null
            } else {
            Write-Verbose "Executing REST API with a webSession variable"
            Invoke-WebRequest -Uri $URI -Method $method -Headers $headers -WebSession $script:session -UseBasicParsing -OutVariable webResult  | Out-Null
            }
        } catch {
            $Failure = $_.Exception.Response
            return $Failure
        }
        return $webResult
    }
}

function Create-RESTBody {
    Param
    (
        [Parameter(Mandatory=$false,HelpMessage='ID of the device')] [string]$ID,
        [Parameter(Mandatory=$false,HelpMessage='objectType')] [string]$objectType
    )

 
        $JSONObject = New-Object -TypeName PSObject
        if ($ID)         { $JSONObject | add-member -name id -value $ID -MemberType NoteProperty           }
        if ($ObjectType) { $JSONObject | add-member -name type -value $objectType -MemberType NoteProperty }
        "[$($JSONObject | ConvertTo-Json)]"
}

function Create-Cookie($name, $value, $domain, $path="/"){
    $c=New-Object System.Net.Cookie;
    $c.Name=$name;
    $c.Path=$path;
    $c.Value = $value
    $c.Domain =$domain;
    return $c;
}

Write-Verbose "Variables:"
Write-Verbose "   Device :   $Device"
Write-Verbose "   UMSServer: $UMSServer"
Write-Verbose "   Command:   $Command"


if (-not(Test-Path "$env:temp\IGELCreds.xml")) {
    Write-Verbose "Saving new credentials..."
    $IGELCreds = Get-Credential -Message "Enter the credentials of an account that can authenticate to the UMS server"
    $IGELCreds | Export-Clixml "$env:temp\IGELCreds.xml"
} else {
    Write-Verbose "Found existing credentials. Importing..."
    $IGELCreds = Import-Clixml "$env:temp\IGELCreds.xml"
}

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($IGELCreds.Password)

# Step 1. Encode the credentials to Base64 string
$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($IGELCreds.UserName):$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))"))

# Step 2. Form the header and add the Authorization attribute to it
$headers = @{ "Authorization" = "Basic $encodedCredentials" 
            "User-Agent"="ControlUp Powershell"
            }

# Most IGEL UMS servers (igelrmserver) dont have valid SSL Cert, thus ignore them
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
# https://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
add-type @"
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

#Set the UMS Server to contain the port number. Set to the defaul 8443 if it was not specified on the parameter
if ($UMSServer -notlike "*:*") {
    $UMSServer = $UMSServer + ":8443"
}
Write-Verbose "UMSServer set to $UMSServer"


    
#test if IGEL server will respond
$serverResponse = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/serverstatus" -headers $headers -method GET -Session
if ($serverResponse.StatusCode -ne 200) 
{
    Write-Error "The script was unable to communicate to the IGEL UMS Server successfully."
    Remove-Item "$env:temp\IGELCreds.xml"
}

#region Setup Session
$IGELSessionId = ($session.Cookies.GetCookies("https://$UMSServer/umsapi/v3/serverstatus")).value
$cookiedomain = ($session.Cookies.GetCookies("https://$UMSServer/umsapi/v3/serverstatus")).Domain
$cookiePath = ($session.Cookies.GetCookies("https://$UMSServer/umsapi/v3/serverstatus")).Path
$IGELAuthCookie = Create-Cookie -name "JSESSIONID" -value "$IGELSessionId" -domain "$cookieDomain" -path "$cookiePath"
Write-Verbose "Created cookie: JSESSIONID:$IGELSessionId"
$script:session.Cookies.Add($IGELAuthCookie)
#endregion Setup Session
 
#region authenticate
$auth = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/login" -headers $headers -method POST

if ($auth.StatusCode -ne "200") {
    Write-Error "Failed to login to the IGEL UMS via REST API.  Please retype your username/password"
    Remove-Item "$env:temp\IGELCreds.xml"
}

Write-Verbose "Authenticated Sucessfully with user $($IGELCreds.UserName).  Setting session information"
#endregion



$selectedDevice = $null
#Select Thin Client
$thinClients = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients" -headers $headers -method GET 
foreach ($thinClient in ($thinClients.Content | ConvertFrom-Json)) {
    if ($Device -like $thinClient.Name) {
        $selectedDevice = $thinClient
        Write-Verbose "Found a match!"
        Write-Verbose "$thinClient"
    }
}

if ($selectedDevice -eq $null) {
    Write-Error "Unable to find device : $Device"
}

$PostBody = Create-RESTBody -ID $selectedDevice.id -objectType $selectedDevice.objectType
$tcID = $selectedDevice.id

switch ($command) {
    "Wakeup"    { $commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/?command=wakeup"                  -headers $headers -method POST -Body $PostBody -ContentType "application/json"}
    "Reboot"    { $commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/?command=reboot"                  -headers $headers -method POST -Body $PostBody -ContentType "application/json"}
    "Shutdown"  { $commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/?command=shutdown"                -headers $headers -method POST -Body $PostBody -ContentType "application/json"}
    "Update"    { $commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/?command=settings2tc"             -headers $headers -method POST -Body $PostBody -ContentType "application/json"}
    "Details"   { $commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/$($tcId)?facets=details"          -headers $headers -method GET }
    "Assest"    { $commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/assetinfo/thinclients/$($tcId)"               -headers $headers -method GET }
    "Profile"   { 
        $commandResult = New-Object System.Collections.Arraylist
        $deviceProfiles = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/$($tcId)/assignments/profiles"             -headers $headers -method GET 
        $allProfiles    = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/profiles"    -headers $headers -method GET 
        foreach ( $deviceProfileId in $(($deviceProfiles.content | ConvertFrom-Json).assignee.id)) {
            foreach ($profileId in $($allProfiles.Content | ConvertFrom-Json)) {
                if ($profileId.id -eq $deviceProfileId) {
                    $commandResult += $profileId.name
                }
            }
        }
    }
}


if ([bool]($commandResult.PSobject.Properties.name -match "Content")) {
    Write-Verbose "$($commandResult.Content)"
    if ($commandResult.Content -like "*CommandExecList*") {
        Write-Output "$(($commandResult.Content | ConvertFrom-Json).commandExecList | Select "message","state" | fl | Out-String)"
    } else {
        Write-Output "$(($commandResult.Content | ConvertFrom-Json) |Out-String)"
    }
} else {
    Write-Output "$($commandResult | Out-String)"
}

