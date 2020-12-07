    <#
    .SYNOPSIS
        Opens the selected server in Windows Admin Center.

    .DESCRIPTION
        Opens the selected server in Windows Admin Center.

    .PARAMETER  <WAC <string>>
        FQDN of your Windows Admin Center server
		
    .PARAMETER  <ComputerName <string>>
        Name of the server to open in WAC

    .PARAMETER  <Workstation <string>>
        True for a workstation/desktop OS or False for a Server OS

    .EXAMPLE
        . .\OpenWithWAC.ps1 -WAC WAC.bottheory.local -ComputerName HYPPVS2019-001 -Workstation $false
        Opens Windows Admin Center with a focus on the server HYPPVS2019-001

    .NOTES
        A HTML5 compatible browser must be your default.  IE11 is not supported with WAC.

    .CONTEXT
        Console

    .MODIFICATION_HISTORY
        Created TTYE : 2020-09-11


    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the FQDN of the Windows Admin Center Gateway or WAC server')][ValidateNotNullOrEmpty()] [string]$WAC,
    [Parameter(Mandatory=$true,HelpMessage='Enter the SamAccountName of the Server')][ValidateNotNullOrEmpty()]                           [string]$ComputerName,
    [Parameter(Mandatory=$true,HelpMessage='Workstation OS or Server OS')][ValidateNotNullOrEmpty()]                                      [string]$Workstation
)

if ($Workstation -like "True") {
    $DesktopOS = $true
} else {
    $DesktopOS = $false
}

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher.Filter = "(&(objectCategory=Computer)(SamAccountname=$($ComputerName)`$))"
$objSearcher.SearchScope = "Subtree"
$ComputerObj = $objSearcher.FindOne()
$dnsHostName = $ComputerObj.Properties.dnshostname


if ($DesktopOS) {
   Start-Process "https://$WAC/computerManagement/connections/computer/$dnsHostName/tools/overview"
} else {
   Start-Process "https://$WAC/servermanager/connections/server/$dnsHostName/tools/overview"
}
