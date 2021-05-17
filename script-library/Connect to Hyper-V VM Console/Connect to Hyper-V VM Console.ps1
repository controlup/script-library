<#
    .SYNOPSIS
    A simple script that will open the VM console on a Hyper-V host. This script assumes you have all necessary permissions
    to do this action

    .DESCRIPTION
    Takes two parameters from ControlUp (Hostname and VM Name) and gets the necessary properties to open
    a console.  This script was written assuming you had permissions to do this.

    AUTHOR: Trentent Tye
    LASTEDIT: 2017-06-13
    VERSI0N : 1.0
    
#>

#GetVM Properties
$vm = Get-VM -name $args[1] -ComputerName $args[0]

#launch VM connect and connect to console
Start-Process VmConnect.exe -argumentList "$args[1] -G $($vm.Id)"

