<# 
.SYNOPSIS
    Adds the selected user(s) to an Active Directory group or groups
.PARAMETER UserName
   The name of the user to be added to the group - automatically supplied by CU
.PARAMETER GroupName
   The name of the group being modified - manually typed in
#>

$ErrorActionPreference = "Stop"

If ( (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Import-Module ActiveDirectory
    } Catch {
        # capture any failure and display it in the error section, then end the script with a return
        # code of 1 so that CU sees that it was not successful.
        Write-Error "Not able to load the Module" -ErrorAction Continue
        Write-Error $Error[1] -ErrorAction Continue
        Exit 1
    }
}

# Because this is the main function of the script it is put into a try/catch frame so that any errors will be 
# handled in a ControlUp-friendly way.

Try {
    Add-ADGroupMember -Identity $args[1] -Members $args[0].split("\")[1]
} Catch {
    # capture any failure and display it in the error section, then end the script with a return
    # code of 1 so that CU sees that it was not successful.
    Write-Error $Error[0] -ErrorAction Continue
    Exit 1
}

Write-Host "The operation completed successfully."

