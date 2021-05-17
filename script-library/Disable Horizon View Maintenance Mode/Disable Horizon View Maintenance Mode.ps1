<#  
.SYNOPSIS
        This script uses the View Connection server and PowerCLI to take a VDI VM out of maintenance mode
.LINK
        Adapted from http://myvirtualcloud.net/?p=3925
#>

$DesktopName = $args[0]
$state="READY"

$ErrorActionPreference = "Stop"
Function Load-HorizonView-PowerCLI ()
{
    $PCLI = "VMware.VimAutomation.HorizonView"

	If ((Find-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null) 
	{
	            Try {
	                  Import-Module $PCLI
	            } Catch {
	                  Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
	                  Exit 1
	            }
	}
	elseIf ((Get-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null) 
	{
	            Try {
	                  Import-Module $PCLI
	            } Catch {
	                  Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
	                  Exit 1
	            }
	}
	elseif( (Get-PSSnapin -Name Vmware.View.Broker -ErrorAction SilentlyContinue) -eq $null )
	{
	        # using try/catch can stop the script completely if needed with "Exit with error" - 'Exit 1' (or some other non-zero exit code)
	        # and avoid a long string of errors because the first statement was not successful.
	        Try {
	                Add-PsSnapin Vmware.View.Broker
	        } Catch {
	                Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
	                Exit 1
	        }
	}
	Else 
	{
        Write-Host "This version of PowerCLI seems to be unsupported. Please upgrade to the latest version of PowerCLI and try again."
    }
}
Load-HorizonView-PowerCLI

Try {
    $ldapBaseURL = 'LDAP://127.0.0.1:389/'
    $Machine_Id = (Get-DesktopVM -Name $DesktopName).machine_id
    $objMachine_Id = [ADSI]($ldapBaseURL + "cn=" + $Machine_Id + ",ou=Servers,dc=vdi,dc=vmware,dc=int")
    $objMachine_Id.put("pae-vmstate", $state)
    $objMachine_Id.setinfo()
    Write-Host "Successfully exited maintenance mode"
} Catch {
    Write-Host "An error on the View server prevented this action from completing successfully."
    Exit 1
}

