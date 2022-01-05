<#
    .SYNOPSIS
        Sets a new password for LAPS protected machine

    .DESCRIPTION
        Sets a new password password for a machine protected by the Local Administrator Password Solution

    .PARAMETER ComputerName
        Specify the computer name of the target machine to reset the password

    .EXAMPLE
        . .\LAPS_ResetPassword.ps1 -ComputerName W2019-001
        Gets the LAPS password for the target machine

    .NOTES
        Designed to run as the CONSOLE context so the user requires rights to get/set the password

    .CONTEXT
        Console

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
$oldPassword = $ComputerObj.Properties["ms-Mcs-AdmPwd"]

#ADSI object has put method to update AD Attribute
$ADSICompObj = [adsi]$ComputerObj.Path
$password = $ComputerObj.Properties["ms-Mcs-AdmPwd"]
$ADSICompObj.Put("ms-Mcs-AdmPwdExpirationTime","0")
#update attribute
$ADSICompObj.SetInfo()

#tell LAPS to query AD for expiration time and set new password
Invoke-Command -ComputerName $ComputerName -ScriptBlock { Start-Process -FilePath gpupdate.exe -ArgumentList @("/Target:Computer") -Wait -WindowStyle Hidden }

#requery AD for updated attributes
$ComputerObj = $objSearcher.FindOne()
$newPassword = $ComputerObj.Properties["ms-Mcs-AdmPwd"]

if ($newPassword -ne $oldPassword) {
    Write-Output "LAPS Account password reset"
} else {
    Write-Output "Password update failed."
}

##It might be worth enabling LAPS verbose logging, gpupdate, and then look in the log for validation that password was updated, then disable logging, gpupdate.
