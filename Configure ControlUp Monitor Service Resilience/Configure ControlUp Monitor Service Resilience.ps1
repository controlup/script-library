#Source: https://evotec.xyz/set-service-recovery-options-powershell/
$svcname = "cuMonitor"
$machinename = $args[0]

if( ! ( Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\ -Name ServicesPipeTimeout  -ErrorAction SilentlyContinue ) )
{
    if( ! ( $newValue = Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\ -Name ServicesPipeTimeout -Value 90000 -PassThru ) )
    {
        Write-Error -Exception "Failed to create reg's tree value"
    }
}

function Set-ServiceRecovery{
    [alias('Set-Recovery')]
    param
    (
        [string] [Parameter(Mandatory=$true)] $ServiceName,
        [string] [Parameter(Mandatory=$true)] $Server,
        [string] $action1 = "restart",
        [int] $time1 =  60000, # in miliseconds
        [string] $action2 = "restart",
        [int] $time2 =  60000, # in miliseconds
        [int] $resetCounter = 4000 # in seconds
    )
    $serverPath = "\\" + $server
    $services = Get-CimInstance -ClassName 'Win32_Service' -ComputerName $Server| Where-Object {$_.Name -imatch $ServiceName}
    $action = $action1+"/"+$time1+"/"+$action2+"/"+$time2+"/"+$actionLast+"/"+$timeLast
    foreach ($service in $services){
        # https://technet.microsoft.com/en-us/library/cc742019.aspx
        $output = sc.exe $serverPath failure $($service.Name) actions= $action reset= $resetCounter
        $output = sc.exe $serverPath failureflag $($service.Name) 1
    }
}
Set-ServiceRecovery -ServiceName $svcname -Server $machinename
