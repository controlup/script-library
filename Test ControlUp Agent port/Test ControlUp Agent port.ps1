if ($args[2] -eq "False") {
$WarningPreference = 'SilentlyContinue'
Test-netconnection -computername $args[0] -port $args[1] | Select-Object PingSucceeded,TcpTestSucceeded
}

if ($args[2] -eq "True") {
Test-netconnection -computername $args[0] -port $args[1]
}
