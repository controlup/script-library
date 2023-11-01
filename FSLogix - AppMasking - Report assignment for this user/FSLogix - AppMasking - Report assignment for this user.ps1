<#
    .SYNOPSIS
        Reports the assigned applications to either the user or the machine

    .DESCRIPTION
        Reports the assigned applications to either the user or the machine

    .EXAMPLE
        . .\Get-FSLogixAssignedAppMaskingDetails.ps1 -username "CN=amttye,OU=Administrators,OU=Jupiter Lab Users,DC=jupiterlab,DC=com"
        Reports on which FSLogix AppMasking rules apply to this user

    .EXAMPLE
        . .\Get-FSLogixAssignedAppMaskingDetails.ps1
        Reports on which FSLogix AppMasking rules apply to which users on this machine

    .PARAMETER username
        distinguishedName of user

        .NOTES
        Returns AppMasking assigned applications for the machine or user
        CONTEXT : Session/Machine
        MODIFICATION_HISTORY
        Created TTYE : 2023-08-16
        AUTHOR: Trentent Tye
#>


[CmdLetBinding()]
Param (
    [parameter(Mandatory=$false)][string]$username
)

$ProgramFiles =[Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)

[string]$FSLogixRulePath = Join-Path -Path $ProgramFiles -ChildPath "FSLogix\Apps\Rules"
[string]$FRXPath = Join-Path -Path $ProgramFiles -ChildPath 'FSLogix\Apps\frx.exe'

[object[]]$FSLogixRules = Get-ChildItem -Path $FSLogixRulePath -Filter *.fxa
foreach ($rule in $FSLogixRules) {
    Write-Output "============================================================================"
    Write-Output "Rule: $($rule.fullname)"
    if ($PSBoundParameters.ContainsKey('username')) {
        $FRXOutputText = & $FRXPath report-assignment -filename "$($rule.fullname)" -username "$username"
        $FRXOutputText.Replace("userOperation","user`nOperation")
    } else {
        & $FRXPath report-assignment -filename "$($rule.fullname)" -verbose
    }
}

Write-Output "$($FSLogixRules.Count) rules found in $FSLogixRulePath"



