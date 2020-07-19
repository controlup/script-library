<#
.SYNOPSIS
    Unlocks the Active Directory account of the selected user(s).
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

Try {
    Unlock-ADAccount -Identity $args[0].split("\")[1]
} Catch {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error $Error[0] -ErrorAction Continue
    Exit 1
}

Write-Host "The user account is unlocked."

