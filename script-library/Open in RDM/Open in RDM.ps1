#Convert the FQDN into a flat name
If ($args[0].IndexOf(".") -gt -1) { 
    $target = $args[0].Substring(0,$args[0].IndexOf("."))
} else { 
   $target = $args[0] 
}

(New-Object -Com Shell.Application).Open("rdm://find/?host=$target")

