<#
.SYNOPSIS
    Resets the password for the Active Directory account of the selected user.
.DESCRIPTION
    Resets the user password. Best to run this on one user at a time unless you want to assign the same password to all target accounts.
.PARAMETER Identity
    Specifies an Active Directory user object - automatically supplied by CU
.PARAMETER NewPassword
    The new password - manually entered by user.
#>

$ErrorActionPreference = "Stop"

If ($args[1] -ne $args[2])
{
    Write-Error "The new password does not match the confirmation. Please re-enter the password and try again." -ErrorAction Continue
    Exit 1
}
Else {
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

    # Because this is the main function of the script it is put into a try/catch frame so that any errors will be 
    # handled in a ControlUp-friendly way.

    Try {
        Set-ADAccountPassword -Identity $args[0].split("\")[1] -NewPassword (ConvertTo-SecureString -AsPlainText $args[1] -Force) -Reset
    } Catch {
        # capture any failure and display it in the error section, then end the script with a return
        # code of 1 so that CU sees that it was not successful.
        Write-Error $Error[0] -ErrorAction Continue
        Exit 1
    }

    Write-Host "The operation completed successfully."
}

