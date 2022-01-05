<#
    .SYNOPSIS
        Get password for LAPS protected machine

    .DESCRIPTION
        Retrieves the password for a machine protected by the Local Administrator Password Solution

    .EXAMPLE
        . .\LAPS_GetPassword.ps1 -ComputerName W2019-001
        Gets the LAPS password for the target machine

    .NOTES
        Designed to run as the CONSOLE context on the target machine so the user running the script requires full rights to get/set the password

    .CONTEXT
        CONSOLE

    .MODIFICATION_HISTORY
        Created TTYE : 2020-09-27


    AUTHOR: Trentent Tye
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Enter the SamAccountName of the machine')][ValidateNotNullOrEmpty()]  [string]$ComputerName
)

#Use native ADSI queries to avoid using ActiveDirectory powershell modules (which might not be installed on the target machines)
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher.Filter = "(&(objectCategory=Computer)(SamAccountname=$($COMPUTERNAME)`$))"
$objSearcher.SearchScope = "Subtree"
$ComputerObj = $objSearcher.FindOne()
$password = $ComputerObj.Properties["ms-Mcs-AdmPwd"]

#find local administrator account
$account = Get-WmiObject -ComputerName $ComputerName -Class Win32_UserAccount -Filter "LocalAccount='True' And Sid like '%-500'"

#find password expiration for LAPS account
$PasswordExpiration = $([datetime]::FromFileTime([convert]::ToInt64($ComputerObj.Properties['ms-MCS-AdmPwdExpirationTime'],10)))


Write-Output "Account          : $($account.caption)"
Write-Output "Password         : $password"
Write-Output "`nPassword Expires : $PasswordExpiration"
