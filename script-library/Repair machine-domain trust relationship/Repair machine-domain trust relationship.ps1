<#
 .SYNOPSIS      Repair trust relationship between a machine and the domain 
 .DESCRIPTION
   There are many situations for which a machine (server OS or workstation) will lose domain trust. 
   If the ControlUp agent is installed on the machine, this script will execute locally and repair the domain trust.
   Domain credentials with permission to reset the computer account must be provided.
   
 .EXAMPLE      repair-domain-trust.ps1 -userName "Domain\user" -userPassword "clearTextPasswd" 
 .CONTEXT      Machine
 .CREDIT
               https://thinkpowershell.com/fix-trust-relationship-workstation-primary-domain-failed/
 .MOD_HISTORY
               2020-05-05 -  Marcel Calef  - created

#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Domain user with admin rights')][ValidateNotNullOrEmpty()]  [string]$userName,
    [Parameter(Mandatory=$true,HelpMessage='clear text password')][ValidateNotNullOrEmpty()]            [string]$userPassword
    )

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$VerbosePreference = "continue"


# Convert to SecureString and create PSCredential object
[securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
[pscredential]$cred = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)

Write-Output "Test-ComputerSecureChannel result before repair:"
$trustOK = Test-ComputerSecureChannel -Credential $cred 

if ($trustOK -like 'True'){Write-Output "Trust test passed, no need to repair"; exit }

# Run repair command up to 5 times or until the repair is succesful
$i = 0
if ($trustOK -like 'False' -and $i -le 4){
        sleep 1
        Write-Output "Trust test failed, need repair"
        $repairAttempt = Test-ComputerSecureChannel -Credential $cred -Repair
        if ($repairAttempt -like 'True'){Write-Output "Repair worked"; $i = 5}
        $i++
        }
sleep 5
Write-Output "Running Test-ComputerSecureChannel again"
Test-ComputerSecureChannel -Credential $cred 
