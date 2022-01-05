#requires -Version 3.0

<#
    .SYNOPSIS
    This script will run the vCheck for VMware Horizon with settings based on the arguments provided in the Script Action

    .DESCRIPTION
    The script checks if the folder for the vCheck exists and if not downloads the check en unpacks it.
    It will remove the globalvariables and connetcion scripts and re-create them based on the provided arguments
    vCheck uses the VMware Horizon SOAP api's so PowerCLI is required.
    There are two required parameters, if any of the other are supplied all of them need to be given or they will be ignored.

    .PARAMETER connectionserver
    [REQUIRED]fqdn of the connectionserver to connect to.

    .PARAMETER outputpath
    [REQUIRED]Drive and folder where the output html file will be saved i.e. c:\vcheckreports.

    .PARAMETER sendmail
    Needs to be supplied true or false to send emails or not.

    .PARAMETER emailusessl
    true or false to use SSL for connection to the smtp server.

    .PARAMETER smtpserver
    smtp server address.

    .PARAMETER fromaddress
    Email address for the from field

    .PARAMETER toaddress
    Email address for the to field

    .PARAMETER emailsubject
    Subject of the email

    .NOTES
    

    .COMPONENT
    VMware PowerCLI 12

    .AUTHOR
    Wouter Kursten
#>

[CmdletBinding(DefaultParameterSetName = 'noemail')]
Param (
    [Parameter(ParameterSetName = 'noemail', Mandatory = $true)]
    [Parameter(ParameterSetName = 'email')]
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'Minimum days since last modification of the file(s) to be deleted, in days. 0 = all modification dates.')]
    [string]$connectionserver,
    [Parameter(Position = 1, Mandatory = $true, HelpMessage = 'Path where the html file will be saved.')]
    [string]$outputpath,
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'true to send emails false for not sending them.')]
    [string]$sendmail,
    [Parameter(Position = 3, Mandatory = $false, HelpMessage = 'Use SSL for the smtp connection.')]
    [string]$emailusessl,
    [Parameter(ParameterSetName = 'email', Position = 4, Mandatory = $false, HelpMessage = 'Address of the smtp server.')]
    [string]$smtpserver,
    [Parameter(ParameterSetName = 'email', Position = 5,Mandatory = $false, HelpMessage = 'E-Mail address the email is send from.')]
    [string]$fromaddress,
    [Parameter(ParameterSetName = 'email', Position = 6,Mandatory = $false, HelpMessage = 'E-Mail address the email is send to.')]
    [string]$toaddress,
    [Parameter(ParameterSetName = 'email', Position = 7,Mandatory = $false, HelpMessage = 'Subject of the email to send')]
    [string]$emailsubject
)

$ErrorActionPreference = 'Stop'
Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If only Message is passed this message is displayed
    If Warning is specified the message is displayed in the warning stream (Message must be included)
    If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
    If an Exception is passed a warning is displayed and the exception is thrown
    If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )

    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

function Load-VMWareModules {
    <# Imports VMware PowerCLI modules
    NOTES:
    - The required modules to be loaded are passed as an array.
    - If the PowerCLI versions is below 12 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
    #>

    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware module to be loaded. Can be single or multiple values (as array).")]
        [array]$Components
    )

    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    # Import each module, if Import-Module fails try Add-PSSnapin
    foreach ($component in $Components) {
        try {
            $null = Import-Module -Name VMware.$component
        }
        catch {
            try {
                $null = Add-PSSnapin -Name VMware
            }
            catch {
                Out-CUConsole -Message 'The required VMWare PowerShell components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 12 or higher required) is installed and available for the user running the script.' -Warning $_
            }
        }
    }
}



$header='$Title = "Connection settings for Horizon"
$Author = "Wouter Kursten"
$PluginVersion = 0.3
$Header = "Connection Settings"
$Comments = "Connection Plugin for connecting to Horizon"
$Display = "None"
$PluginCategory = "View"'

$footer = '$creds = import-clixml $credsfile

# Loading 
Import-Module VMware.VimAutomation.HorizonView
Import-Module VMware.VimAutomation.Core

# --- Connect to Horizon Connection Server API Service ---
$hvServer1 = Connect-HVServer -Server $server -credential $creds

# --- Get Services for interacting with the Horizon API Service ---
$Services1= $hvServer1.ExtensionData

# --- Get Desktop pools
$poolqueryservice=new-object vmware.hv.queryserviceservice
$pooldefn = New-Object VMware.Hv.QueryDefinition
$pooldefn.queryentitytype="DesktopSummaryView"
$poolqueryResults = $poolqueryService.QueryService_Create($Services1, $pooldefn)
$pools = foreach ($poolresult in $poolqueryResults.results){$services1.desktop.desktop_get($poolresult.id)}
$poolqueryservice.QueryService_DeleteAll($services1)

# --- Get RDS Farms

$Farmqueryservice=new-object vmware.hv.queryserviceservice
$Farmdefn = New-Object VMware.Hv.QueryDefinition
$Farmdefn.queryentitytype="FarmSummaryView"
$FarmqueryResults = $FarmqueryService.QueryService_Create($Services1, $Farmdefn)
$farms = foreach ($farmresult in $farmqueryResults.results){$services1.farm.farm_get($farmresult.id)}
$Farmqueryservice.QueryService_DeleteAll($services1)
'
Load-VMwareModules -Components @('VimAutomation.HorizonView')

$username = $env:USERNAME
$machinename = $env:COMPUTERNAME

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

if(!(test-path $strCUCredFolder)){
    write-error "No ControlUp Scriptsupport folder found, please create a Horizon Credentials file first for the account $username"
}
if(!(test-path $strCUCredFolder"\vCheck-HorizonView-master")){
    invoke-webrequest -uri "https://github.com/vCheckReport/vCheck-HorizonView/archive/refs/heads/master.zip" -outfile $strCUCredFolder"\vcheck.zip"
    Expand-Archive $strCUCredFolder"\vcheck.zip" -DestinationPath $strCUCredFolder
    get-item $strCUCredFolder"\vcheck.zip" | remove-item
}

$connectionplugin = $strCUCredFolder+"\vCheck-HorizonView-master\Plugins\00 Initialize\00 Connection Plugin for View.ps1"
$globalvariablesfile = $strCUCredFolder+"\vCheck-HorizonView-master\GlobalVariables.ps1"

if(test-path $connectionplugin){
    get-item  $connectionplugin | remove-item
}
if(test-path $globalvariablesfile){
    get-item $globalvariablesfile | remove-item
}

new-item $connectionplugin -type "file" | out-null
new-item $globalvariablesfile -type "file" | out-null

$credsfile = $strCUCredFolder+'\'+$username+'_horizonView_Cred.xml'

$serverline = '$Server = '+'"'+$connectionserver+'"'
$credsline = '$credsfile = '+'"'+$credsfile+ '"'

$header | out-file $connectionplugin -append
$serverline | out-file $connectionplugin -append
$credsline | out-file $connectionplugin -append
$footer | out-file $connectionplugin -append

# Makes sure the setup wizard doesn't run
write-output '$SetupWizard = $False' | out-file $globalvariablesfile -append
# Name of the header
write-output '$reportHeader = "vCheck"' | out-file $globalvariablesfile -append
# We run this remotely so no opening the file directly
write-output '$DisplaytoScreen = $false' | out-file $globalvariablesfile -append
# Display the report even if it is empty?
write-output '$DisplayReportEvenIfEmpty = $false' | out-file $globalvariablesfile -append

# Use the following item to define if an email report should be sent once completed

if($sendmail -eq "true" -and $emailsubject -ne ""){
    $sendmailstatus = "Email send from $fromaddress to $toaddress. The html file has been saved  to $outputpath on $machinename"
    write-output '$SendEmail = $true' | out-file $globalvariablesfile -append
    $smtpline = '$SMTPSRV = "'+$smtpserver+'"'
    $emailfromline = '$EmailFrom = "'+$fromaddress+'"'
    $emailtoline = '$EmailTo = "'+$toaddress+'"'
    write-output '$EmailCc = ""' | out-file $globalvariablesfile -append
    $emailsubjectline = '$EmailSubject = "'+$EmailSubject+'"'
    write-output $smtpline | out-file $globalvariablesfile -append
    if($emailusessl -eq "true"){
        write-output '$EmailSSL = $true' | out-file $globalvariablesfile -append
    }
    else{
        write-output '$EmailSSL = $false' | out-file $globalvariablesfile -append
    }
    $smtpline | out-file $globalvariablesfile -append
    $emailfromline | out-file $globalvariablesfile -append
    $emailtoline | out-file $globalvariablesfile -append
    $emailccline | out-file $globalvariablesfile -append
    $emailsubjectline | out-file $globalvariablesfile -append
}
else {
    $sendmailstatus = "Sending email not enabled or not all arguments have been provided. The html file has been saved to $outputpath on $machinename" 
    write-output '$SendEmail = $false' | out-file $globalvariablesfile -append
    # Please Specify the SMTP server address (and optional port) [servername(:port)]
    write-output '$SMTPSRV = "mysmtpserver.mydomain.local"' | out-file $globalvariablesfile -append
    # Would you like to use SSL to send email?
    write-output '$EmailSSL = $false' | out-file $globalvariablesfile -append
    # Please specify the email address who will send the vCheck report
    write-output '$EmailFrom = "me@mydomain.local"' | out-file $globalvariablesfile -append
    # Please specify the email address(es) who will receive the vCheck report (separate multiple addresses with comma)
    write-output '$EmailTo = "me@mydomain.local"' | out-file $globalvariablesfile -append
    # Please specify the email address(es) who will be CCd to receive the vCheck report (separate multiple addresses with comma)
    write-output '$EmailCc = ""' | out-file $globalvariablesfile -append
    # Please specify an email subject
    write-output '$EmailSubject = "$Server vCheck Report"' | out-file $globalvariablesfile -append

}
# Send the report by e-mail even if it is empty?
write-output '$EmailReportEvenIfEmpty = $true' | out-file $globalvariablesfile -append
# If you would prefer the HTML file as an attachment then enable the following:
write-output '$SendAttachment = $false' | out-file $globalvariablesfile -append
# Set the style template to use.
write-output '$Style = "Clarity"' | out-file $globalvariablesfile -append
# Do you want to include plugin details in the report?
write-output '$reportOnPlugins = $true' | out-file $globalvariablesfile -append
# List Enabled plugins first in Plugin Report?
write-output '$ListEnabledPluginsFirst = $true' | out-file $globalvariablesfile -append
# Set the following setting to $true to see how long each Plugin takes to run as part of the report
write-output '$TimeToRun = $true' | out-file $globalvariablesfile -append
# Report on plugins that take longer than the following amount of seconds
write-output '$PluginSeconds = 30' | out-file $globalvariablesfile -append
$ErrorActionPreference = 'Continue'

& $strCUCredFolder"\vCheck-HorizonView-master\vcheck.ps1" -Outputpath $outputpath
write-output $sendmailstatus
