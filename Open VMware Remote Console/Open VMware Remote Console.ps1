<#  
.SYNOPSIS     Open VMware Remote Console using REST API
.DESCRIPTION  Leverage the vCenter REST API To get a console ticket and open the VMRC for the provided VM name
              Credentials are requested and stored in a PSCredential object in the user TEMP directory

.REFERENCES   https://vmware.github.io/vsphere-automation-sdk-rest/vsphere/operations/com/vmware/vcenter/vm/console/tickets.create-operation.html

.CONTEXT      Machines
.COMPONENT    vCenter 7.0 or newer and VMRC
.TAGS         $VirtualMachine,$VMware
.HISTORY      Marcel Calef     - 2020-11-20 - Initial release
              Guy Leech        - 2021-09-22 - Improvements
              Guy Leech        - 2021-10-04 - Changed check for vCenter version from 6.7 to 7.0 as console/tickets API does not exist before 7.0
                                              Added vcenter name to credential file nname to allow to work with different vCenters
                                              Error if no credentials entered
                                              Added code to deal with deprecation of /rest/ in URI prior to 7.0U2
              Guy Leech        - 2021-10-12 - Fallback to PS cmdlets if vCenter version before 7.0 as doesn't have REST API to get required vmrc ticket
              Guy Leech        - 2022-08-15 - Fix for not deleting creds on bad auth via REST
              Guy Leech        - 2022-08-16 - Fix for fix not deleting creds on bad auth via REST, better handling of vmrc.exe not found
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='vmName')]                   [string]$vmName,
    [Parameter(Mandatory=$true,HelpMessage='apiHost')]                  [string]$ConnectURL
      )

$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { 'Continue' } else { 'SilentlyContinue' })
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { 'Continue' } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'ErrorAction' ] ) { $ErrorActionPreference } else { 'Stop' })

Function Open-VMRC {
    <#
    .SYNOPSIS     Launch a VMware Remote Console for a named VM
    .DESCRIPTION  TBA
    .COMPONENT    http://www.vmware.com/go/download-vmrc
    .NOTES
           Based on Open-VMRC from: Allen Derusha
           Simplified for use with a ticket from the REST API: Marcel Calef
    #>

    Param( 
        [Parameter(Mandatory=$True,HelpMessage="VMRC ticket")]
        [String]$vmrcTicket )
    
    [string]$vmrcexe = 'vmrc.exe'

    ## Should be able run it as installer adds to HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\vmrc.exe
    try
    {
        $process = Start-Process -FilePath $vmrcexe -ArgumentList $vmrcTicket -ErrorAction SilentlyContinue -PassThru
    }
    catch
    {
        $process = $null
    }

    if( -Not $process )
    {
        Write-Verbose -Message "Looking for $vmrcexe"

        $VMRCPath = $null

        # Check to see if we can find VMRC.exe and tell the user where to download it if we can't find it
        If (Test-Path "${env:ProgramFiles(x86)}\VMware\VMware Remote Console\vmrc.exe") {
          $VMRCpath = "${env:ProgramFiles(x86)}\VMware\VMware Remote Console\vmrc.exe"
        } ElseIf (Test-Path "$env:ProgramFiles\VMware\VMware Remote Console\vmrc.exe") {
          $VMRCpath = "$env:ProgramFiles\VMware\VMware Remote Console\vmrc.exe"
        } Else {
            ## TODO search for it and check signed
        }
        if( $VMRCPath ) {
            $process = Start-Process -FilePath $VMRCpath -ArgumentList $vmrcTicket -PassThru ## "vmrc://clone:$Ticket@$Server`:443/?moid=$VMmoRef"
        }
        else {
            Write-Error "Could not find $vmrcexe.  Download and install the VMRC package from VMware at https://www.vmware.com/go/download-vmrc"
        }
    }
    else
    {
        $process ## return
    }    
}


# Remove prefix and suffix from the Hypervisor Connection URL from ControlUp
try {
    $vCenter = $ConnectURL.Split("/")[2]
} 
catch {
    Throw "Please provide ConnectURL as https://vCenter/sdk."
}

Write-Verbose "Variables:"
Write-Verbose "         vmName :  $vmName"
Write-Verbose "        apiHost :  $ConnectURL"
Write-Verbose "        vCenter :  $vCenter"

# Most vCenter dont have valid SSL Cert, thus ignore them
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

## put vcenter in the credential file name so can use different credentials with different vcenters
[string]$vCenterCredsFile = Join-Path -Path $env:temp -ChildPath "vCenterCreds.$vcenter.xml"

#### Connect to vCenter REST API & gather the information needed

# Step 1. Get previously saved creds or request new ones
if (-not(Test-Path -Path $vCenterCredsFile )) {
    Write-Verbose "Saving new credentials..."
    if( $vCenterCreds = Get-Credential -Message "Enter the credentials of an account for vCenter $vCenter" ) {
        $vCenterCreds | Export-Clixml -Path $vCenterCredsFile
    } else {
        Throw "Failed to get vCenter credentials for $ConnectURL"
    }
} else {
    Write-Verbose "Found existing credentials. Importing..."
    if( ! ( $vCenterCreds = Import-Clixml -Path $vCenterCredsFile ) ) {
        Throw "Cannot use credentials in $vCenterCredsFile"
    }
}

# Step 2. Encode the credentials to Base64 string
$base64AuthInfo = [System.Convert]::ToBase64String( [System.Text.Encoding]::ASCII.GetBytes( "$($vCenterCreds.UserName):$($vCenterCreds.GetNetworkCredential().password)" ))

# Step 3. Form the header and add the Authorization attribute to it
[hashtable]$authHeaders = @{
    "Authorization" = "Basic $base64AuthInfo"
}

## /rest/ to call REST API is deprecated as of 7.0U2
[string]$RESTAPI = 'api'
[string]$filter = ''

# Step 4. Get a API Session ID for further Authentication
    #  https://developer.vmware.com/docs/vsphere-automation/latest/
    
    try {
        $serverResponse = (Invoke-RestMethod -Method Post -Uri "https://$vCenter/$RESTAPI/session" -Headers $authHeaders)
    }
    catch 
    {
        ## try the deprecated URI
        $authHeaders.Add( "vmware-use-header-authn" , 'string' )
        $RESTAPI = 'rest'
        $filter = 'filter.'
        try {
            $serverResponse = Invoke-RestMethod -Method Post -Uri "https://$vCenter/$RESTAPI/com/vmware/cis/session" -Headers $authHeaders
        }
        catch {
            $serverResponse = $null
        }
        if( -Not $serverResponse ) {
            Remove-Item -Path $vCenterCredsFile -Force     ## remove the PSCredential as it failed
            Throw "The script was unable to communicate to vCenter $vcenter successfully. $_"
        }
    }

    try {
        if( $serverResponse.PSobject.Properties[ 'value' ] )
        {
            $apiSessionID = $serverResponse.value ## pre 7.0U2
        }
        else
        {
            $apiSessionID = $serverResponse ## 7.0U2
        }
        Write-Verbose "        apiSessionID :  $apiSessionID"}
    catch {
        Remove-Item -Path $vCenterCredsFile -Force          ## remove the PSCredential as it failed
        Throw "The script was unable to communicate to obtain a vCenter SessionID successfully.`n$_"
    }

# Step 5. Prepare a header with the Session ID for further requests

[hashtable]$headers = @{ "vmware-api-session-id" = $apiSessionID }

### Confirm vCenter version
Try {
    $vCenterVersion = (Invoke-RestMethod -Method Get -Uri "https://$vCenter/rest/appliance/system/version" -Headers $headers).value.version
}
Catch { 
    Remove-Item -Path $vCenterCredsFile -Force         ## remove the PSCredential as it failed
    Throw "The script was unable to retrive the vCenter version`n$_"
}

Write-Verbose "        vCenterVersion :  $vCenterVersion"

$ticket = $null

if( [version]$vCenterVersion -lt [version]'7.0' )
{
    ##Write-Warning -Message "vCenter needs to be at least 7.0 - $vmName is on vCenter $vcenter which is version $vCenterVersion - trying PowerShell cmdlets"
    if( -Not ( Import-Module -Name VMware.VimAutomation.Core -Verbose:$false -PassThru ) -or -Not (Get-Command -Name Get-View -ErrorAction SilentlyContinue ) )
    {
        Write-Warning -Message "Unable to import VMware PowerCLI module VMware.VimAutomation.Core - is it installed? See https://docs.vmware.com/en/VMware-vSphere/7.0/com.vmware.esxi.install.doc/GUID-F02D0C2D-B226-4908-9E5C-2E783D41FE2D.html"
    }
    else
    {
        [hashtable]$connectParameters = @{
            Server = $vCenter
            Credential = $vCenterCreds
            Protocol = ($ConnectURL -split ':')[0]
            Force = $true
        }
        if( ! ( $connection = Connect-VIServer @connectParameters ) )
        {
            Write-Warning "Failed to connect to vCenter server $($connectParameters['Server'])"
        }
        if( ! ( $vm = Get-VM -Name $vmName ) )
        {
            Write-Warning -Message "Failed to get VM $vmName from vCenter $($connectParameters['Server'])"
        }
        elseif( $Session = Get-View -Id Sessionmanager -ErrorAction SilentlyContinue )
        {
            if( $sessionTicket = $Session.AcquireCloneTicket() )
            {
                $ticket = "vmrc://clone:$sessionTicket@$($connection.ToString())/?moid=$($vm.ExtensionData.MoRef.value)"
            }
            else
            {
                Write-Warning "Failed to get ticket from vCenter session on $($connectParameters['Server'])"
            }
        }
        else
        {
            Write-Warning "Failed to get session from vCenter $($connectParameters['Server'])"
        }
        if( $connection )
        {
            $connection | Disconnect-VIServer -Force -Confirm:$false
            $connection = $false
        }
    }
}
else
{
    ## Get vmID from VMname - note that the VM name is case sensitive

    try { 
        if( ( $vm = Invoke-RestMethod -Method Get -Uri  "https://$vCenter/$RESTAPI/vcenter/vm?$($filter)names=$vmName" -Headers $headers ) )
        {
            if( -Not ( $vmID = $vm | Select-Object -ExpandProperty value -ErrorAction SilentlyContinue | Select-Object -ExpandProperty vm -ErrorAction SilentlyContinue ) )
            {
                $vmID = $vm.vm
            }
        }
        Write-Verbose "        vmID :  $vmID"
    }
    catch {
        Throw "The script was unable to resolve the VM name $vmName successfully."
    }

    ## build the VMRC ticker request JSON body/data
    $data = $(if( $RESTAPI -eq 'api' )
    {
        @{
            "type" = "VMRC"
        }
    }
    else
    {
        @{
            "spec" = @{
                        "type" = "VMRC"
            }
        }
    })

    $body = $data | ConvertTo-Json

    Write-Verbose "    request body :  $body"

    ## Get the VMRC Ticket
    ## https://developer.vmware.com/docs/vsphere-automation/latest/vcenter/api/vcenter/vm/vm/console/tickets/post/

    if( $request = Invoke-RestMethod -Method Post -Uri  "https://$vCenter/$RESTAPI/vcenter/vm/$($vmID)/console/tickets" -ContentType 'application/json' -Headers $headers -Body $body ) {
        $ticket = $(if( $request.PSObject.Properties[ 'value' ] ) { $request.value.ticket } else { $request.ticket } )
    } else {
        Throw "Failed to get ticket for VM id $vmID from $vcenter"
    }
}

if( $ticket )
{
    Write-Verbose -Message "Ticket is $ticket"

    if( -Not ( $vmrcProcess = Open-VMRC -vmrcTicket $ticket ) ) {
        Throw "Failed to run vmrc for VM $vmname via vCenter $vCenter"
    } else {
        Write-Output "vmrc launched ok to $vmName as pid $($vmrcProcess.Id)"
    }
}

