$priorityhash = @{-2="Idle";-1="BelowNormal";0="Normal";1="AboveNormal";2="High";3="RealTime"} 
 
    (Get-Process -Id $args[0]).priorityclass = $priorityhash[1] 
