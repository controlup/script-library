<#
.SYNOPSIS
    Send an e-mail message to the selected user(s) using Microsoft Outlook.
.DESCRIPTION
    Sends an email; requires the AD PowerShell module and that Outlook is installed on the computer running the CU console.
.PARAMETER User
    Specifies an Active Directory user object - automatically supplied by CU
#>

$ErrorActionPreference = "Stop"

If ( (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Import-Module ActiveDirectory
    } Catch {
        # capture any failure and display it in the error section, then end the script with a return
        # code of 1 so that CU sees that it was not successful.
        Write-Error "Unable to load the module" -ErrorAction Continue
        Write-Error $Error[1] -ErrorAction Continue
        Exit 1
    }
}

$user = $args[0]
$username = ($user -split "\\",2)[1]
$usermail = (Get-AdUser -Identity $username -Properties mail).mail

If ($usermail -ne $null)
{
    $ol = New-Object -comObject Outlook.Application  
    $mail = $ol.CreateItem(0)  
    $mail.Recipients.Add($usermail) | Out-Null
    Write-Host "Sending e-mail to $user at $usermail"
    $mail.Display()
}
Else
{
    Write-Error 'The target user does not have an email address defined in AD' -ErrorAction Continue
    Exit 1
}

