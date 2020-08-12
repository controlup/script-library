<#
    .SYNOPSIS
        Starts a vnc program to connect to a IGEL endpoint.

    .DESCRIPTION
        Starts a vnc program to connect to a IGEL endpoint.

    .EXAMPLE
        . .\IGEL_Shadow.ps1 -device IGEL01 -UMSServer igelums.acme.local -VNCExe "C:\Program Files\vnviewer.exe"
        Starts a VNC Shadow operation on the targetted IGEL device.

    .CONTEXT
        Machine

    .MODIFICATION_HISTORY
        Created TTYE : 2019-12-12


    AUTHOR: Trentent Tye
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the machine name of the endpoint')][ValidateNotNullOrEmpty()]   [string]$Device,
    [Parameter(Mandatory=$true,HelpMessage='FQDN of the IGEL UMS server')][ValidateNotNullOrEmpty()]              [string]$UMSServer,
    [Parameter(Mandatory=$true,HelpMessage='Path to vnc.exe')][ValidateNotNullOrEmpty()]                          [string]$VNCExe
    
)


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$VerbosePreference = "continue"

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
Write-Verbose "   VNC Path:   $VNCExe"

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
    Write-Error "Failed to login to the IGEL UMS via REST API.  Please check your username/password"
}

Write-Verbose "Authenticated Sucessfully.  Setting session information"
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


$commandResult = Invoke-IGELRestAPI -URI "https://$UMSServer/umsapi/v3/thinclients/$($tcId)?facets=shadow"               -headers $headers -method GET 

$lastIP = ($commandResult.Content | ConvertFrom-Json).lastIp
$VNCHostPort = "$lastIP`:5900"

if (-not(Test-Path -Path $VNCExe)) {
    Write-Output "Unable to find $VNCExe. Check your path and try again" | Msg *
}


$VNCProperties = Get-Item $VNCExe
$VNCProductName = $VNCProperties.VersionInfo.ProductName

## https://kb.igel.com/endpointmgmt-5.08/en/external-vnc-viewer-22459975.html
switch ($VNCProductName) {
    "TightVNC" { Start-Process -FilePath $VNCExe -ArgumentList ("$VNCHostPort") }
    "RealVNC"  { Start-Process -FilePath $VNCExe -ArgumentList ("$VNCHostPort") }
    "TigerVNC" { Start-Process -FilePath $VNCExe -ArgumentList ("$VNCHostPort") }
    "UltraVNC" { Start-Process -FilePath $VNCExe -ArgumentList ("-connect","$VNCHostPort") }
    default { Write-Output "Unable to determine VNC type. Trying default arguments"
              Start-Process -FilePath $VNCExe -ArgumentList ("$VNCHostPort") 
            }
}
