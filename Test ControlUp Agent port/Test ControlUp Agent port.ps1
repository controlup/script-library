$Machine = $args[0]
$port = $args[1]

if ($args[2] -eq "False")
{
    $Test = test-netconnection -ComputerName $machine -Port $port -WarningAction SilentlyContinue
    if (($Test).TcpTestSucceeded)
    {
        "Port " + $port + " is open."
    }
    else
    {
        if ($test.PingSucceeded -eq $true -and $port -eq "40705")
        {
            try
            {
                $Service = get-service -Computername $machine -name 'cuAgent' -ErrorAction Stop
                "Port " + $port + " is closed.`ncuAgent is installed and it is " + $Service.status + "."
            }
            catch
            {
                try{
                    $test = get-service -ComputerName $machine
                    "cuAgent is not installed."
                }
                Catch
                {
                    write-host "Port is not open. It appears you don't have admin rights to check for the existance of the service on the destination computer."
                }
                                   
            }
        }
        elseif ($test.PingSucceeded -eq $true)
        {
            "Port " + $port + " is not open."
        }
        else
        {
            "Ping failed."
        }
    }
}

if ($args[2] -eq "True")
{
    Test-netconnection -computername $args[0] -port $args[1] -InformationLevel Detailed
}


