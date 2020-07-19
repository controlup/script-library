<#  
    .NAME:    Disable or Enable SSH on an VMware ESXi host

                Using PowerCLI, Connect to vCenter and start or stop the SSH service on a host.
                The computer running this script requires PowerCLI and vCenter credentails will be PROMPTED

    .CREDITS: https://vmguru.com/2016/01/powershell-friday-enabling-ssh-with-powercli/
    .AUTHOR:  Marcel Calef  2019-10-04
    .TAGS:    $HypervisorPlatform="VMware"
#>

$sshDesired = $args[0] 
$hypervisor = $args[1]
$input= $args[2] ; $p1,$p2,$vCenter,$sdk= $input.split('/')     ## Remove the https:// and /sdk
$ESXhost = $args[3]

# Checks
    If ($hypervisor -ne "VMware") { Write-host "This script is designed for VMware ESXi hosts only" ; exit }


Connect-VIServer $vCenter   # Credentials will be prompted

Get-VMHostService -VMHost "$ESXhost" | Where-Object {$_.Key -eq "TSM-SSH"}  # list current state 

if ($sshDesired -eq "Stop")  {Get-VMHostService -VMHost "$ESXhost" | Where-Object {$_.Key -eq "TSM-SSH"} |  Stop-VMHostService -Confirm:$false}
if ($sshDesired -eq "Start") {Get-VMHostService -VMHost "$ESXhost" | Where-Object {$_.Key -eq "TSM-SSH"} | Start-VMHostService -Confirm:$false}

