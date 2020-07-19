$target = $args[0]
$cmds = "-Noexit","-command enter-pssession -computername $target"
Start-Process powershell.exe -ArgumentList $cmds
