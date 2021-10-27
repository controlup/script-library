<#
.SYNOPSIS
    Retreives client information for VMware Horizon Clients.

.DESCRIPTION
    Retreives client information for VMware Horizon Clients from the Volatile Environment.

.PARAMETER sessionid
    Id of the session to retreive the information for.

.PARAMETER Username
    This adds the UserName to the overview

.EXAMPLE
    To get client information use .\show_horizon_client_info.ps1 -sessionid 3 -username cuemea\wouterk

.COMPONENT

.NOTES


    Version:        0.1
    Author:         Wouter Kursten
    Creation Date:  30-08-2021
    Purpose:        Retreiving Client information for VMware Horizon Clients

    MODIFICATION_HISTORY

    Wouter Kursten,         30-0-2021 - Original Code

    LINKS
    Information on the registry keys in the Volatile Environment can be found here: https://docs.vmware.com/en/VMware-Horizon/2103/horizon-remote-desktop-features/GUID-86ED59AD-3A2C-4B71-8CFE-19B33E76E571.html

#>

[CmdletBinding()]
Param
(
    [Parameter(
        Position = 0,
        Mandatory=$true,
        HelpMessage='User Session ID'
    )]
    [ValidateNotNullOrEmpty()]
    [int] $sessionid,

    [Parameter(
        Position = 1,
        Mandatory=$true,
        HelpMessage='User Name'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $UserName
)

$volatileenv = Get-ItemProperty "HKCU:\Volatile Environment\$sessionid"
$displaydata = (((($volatileenv | select-object -expandproperty 'ViewClient_Displays.TopologyV2').replace('},{','};{')).replace("{","")).replace("}","")).split(";")

if ($volatileenv -like "*viewclient*"){

    $myObject = [PSCustomObject]@{
        "Session User" = $UserName
        "Client Machine Name"       =   ($volatileenv).ViewClient_Machine_Name
        "Client Machine Domain"     =   ($volatileenv).ViewClient_Machine_Domain
        "Client Type"       =   ($volatileenv).ViewClient_Type
        "Client MAC Address"     =   ($volatileenv).ViewClient_MAC_Address
        "Client Logged On User"     =   ($volatileenv).ViewClient_LoggedOn_Username
        "Client Logged On Domain Name"     =   ($volatileenv).ViewClient_LoggedOn_Domainname
        "Client Logged On FQDN"     =   ($volatileenv).ViewClient_LoggedOn_FQDN
        "Client IP Address"     =   ($volatileenv).ViewClient_IP_Address
        "Client System DPI"     =   ($volatileenv).'ViewClient_Displays.SystemDPI'
        "Client Protocol" = ($volatileenv).ViewClient_Protocol
        "Client Launch ID" = ($volatileenv).ViewClient_Launch_ID
        "Client Launch Session Type" = ($volatileenv).ViewClient_Launch_SessionType
        "Client Time Zone ID" = ($volatileenv).ViewClient_TZID
        "Client Windows Time Zone" = ($volatileenv).ViewClient_Windows_Timezone
        "Client Keyboard Language" = ($volatileenv).'ViewClient_Keyboard.Language'
        "Client Keyboard Layout" = ($volatileenv).'ViewClient_Keyboard.Layout'
        "Client Language" = ($volatileenv).ViewClient_Language
        "Client Version" = ($volatileenv).ViewClient_Client_Version
    }

    $index = 0

    foreach ($display in $displaydata){
        $data = $display.split(",")
        $myObject | Add-Member -NotePropertyName "Display $index Width" -NotePropertyValue $data[0]
        $myObject | Add-Member -NotePropertyName "Display $index Height" -NotePropertyValue $data[1]
        $myObject | Add-Member -NotePropertyName "Display $index left" -NotePropertyValue $data[2]
        $myObject | Add-Member -NotePropertyName "Display $index top" -NotePropertyValue $data[3]
        $myObject | Add-Member -NotePropertyName "Display $index Bits Per Pixel" -NotePropertyValue $data[4]
        $myObject | Add-Member -NotePropertyName "Display $index Is Primary" -NotePropertyValue $data[5]
        $myObject | Add-Member -NotePropertyName "Display $index DPI" -NotePropertyValue $data[6]
        $index += 1
    }
    Write-output "VMware Horizon Client information:"
    write-output $myobject
}
else{
    write-output "No ViewClient Data found, are you sure this is an active VMware Horizon Session?"
}
