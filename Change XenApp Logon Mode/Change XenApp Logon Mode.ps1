﻿<# 
.SYNOPSIS
    Changes the XenApp Logon Mode for the selected server
.PARAMETER LogonMode
   Logon behavior of the server - manually entered by user
.PARAMETER ServerName
   The name of the computer being modified - automatically supplied by CU
#>

$ErrorActionPreference = "Stop"

If ( (Get-PSSnapin -Name Citrix.XenApp.Commands -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Add-PsSnapin Citrix.XenApp.Commands
    } Catch {
        # capture any failure and display it in the error section, then end the script with a return
        # code of 1 so that CU sees that it was not successful.
        Write-Error "Not able to load the SnapIn" -ErrorAction Continue
        Write-Error $Error[1] -ErrorAction Continue
        Exit 1
    }
}

# Because this is the main function of the script it is put into a try/catch frame so that any errors will be 
# handled in a ControlUp-friendly way.

Try {
    Set-XAServerLogOnMode -LogOnMode $args[0] -ServerName $args[1]
} Catch {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error $Error[0] -ErrorAction Continue
    Exit 1
}

Write-Host "The operation completed successfully."

