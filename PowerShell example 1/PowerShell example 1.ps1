# suggested template for PowerShell scripts

<#
Error trapping may be more useful here than what is used normally. A PoSh script can exit 'successfully' 
even when important commands in the script fail. Therefore, more liberal use of try/catch
and if/then/else may be more helpful to capture those script errors.

Comment your scripts! Everyone appreciates self-documented scripts.


# Date     : 31/12/2013
# Author   : John Doe
# Version  : 0.2
#
# Description
# ===========
#
# This script produces ___________
#
# Version History
# ===============
#
# Version  | Date        | Description of Change
# ---------+-------------+------------------------------------------------------
# 0.1      | 00/00/2000  | Initial Version
# 0.2      | 00/00/2000  | Added some tracing / error handling
#          |             |
#          |             |
#
######################################################

#>

<#  another way to document the script...
.SYNOPSIS
        <A brief description of the script>
.DESCRIPTION
        <A detailed description of the script>
.PARAMETER <paramName>
        <Description of script parameter>
.EXAMPLE
        <An example of using the script>
.INPUTS
        Hard work
.OUTPUTS
        Satisfaction
.LINK
        See http://www.google.com
#>

$ErrorActionPreference = "Stop"     #   another way to try to stop the script in case of errors. Important for Try/Catch usage.

If ( (Get-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue) -eq $null )
{
        # using try/catch can stop the script completely if needed with "Exit with error" - 'Exit 1' (or some other non-zero exit code)
        # and avoid a long string of errors because the first statement was not successful.
        Try {
                Add-PsSnapin Citrix.Broker.Admin.*
        } Catch {
                Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                Exit 1
        }
}

# You could also use [CmdletBinding()] and param() to help qualify arguments if the regex validation was not sufficient.
# See the other PowerShell example script.

$machineName = $args[0]

try {
        Some-Command $machineName
}
catch {
        Write-Host " <error with initial command> "
        Exit 1
}

If ($condition -ne $null) {
        Some-Command2
        Write-Host " <condition1 successful> "
} else {
        Some-Command3
        If ($condition -ne $null) {
                Some-Command4
                Write-Host " <condition2 successful> "
        } else {
                Write-Host "Could not complete the action"
                Exit 1
        }
} 

