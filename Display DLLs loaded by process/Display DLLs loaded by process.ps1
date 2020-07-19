Try {
    $Modules = (Get-Process -PID $args[0] -ErrorAction Stop).Modules
}
Catch {
    $_ | fl *
    Exit 1
}

$Modules | ft FileName

