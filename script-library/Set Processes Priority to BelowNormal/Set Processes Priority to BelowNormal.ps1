$sessionId = $args[0]

$userProcesses = get-process | Where {$_.SI -eq $sessionId}
$priorityhash = @{-2="Idle";-1="BelowNormal";0="Normal";1="AboveNormal";2="High";3="RealTime"} 

foreach ($process in $userProcesses) {
    try {
        if ($process.priorityclass -eq $priorityhash[0]) {
            (Get-Process -Id $process.id).priorityclass = $priorityhash[-1] 
        }
    } catch {
        Write-Host "Skipped process $($process.name)"
    }
}
