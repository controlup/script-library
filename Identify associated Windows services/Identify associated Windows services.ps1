$computer = "."
$InputProcessId = $args[0]
if ($args[0] -eq $null) {
    Write-Host "This script expects the PID as a single argument."
    Exit 1
} else {
    try {
        $ProcObject = Get-WmiObject -Class Win32_Process -ComputerName $computer -Filter "ProcessId=$InputProcessId" -ErrorAction Stop
    } catch {}
    if ($ProcObject -eq $null) {
        Write-Host "The selected process was not found on this system."
        Exit 1
    } else {
        $results = ($ProcObject | ForEach {
            $process = $_
                Get-WmiObject -Class Win32_Service -ComputerName $computer -Filter "ProcessId=$($_.ProcessId)" | % {
                    New-Object PSObject -Property @{ProcessId=$process.ProcessId;
                                                                            ServiceName=$_.Name;
                                                                            State=$_.State;
                                                                            DisplayName=$_.DisplayName;
                                                                            StartMode=$_.StartMode}
                    }
        })
        if ($results -eq $null) {
            Write-Host "The selected process is not associated with any Windows services."
        } else {
            Write-Host "The selected process is associated with the following Windows service/s:`r`n"
            $results | ft -AutoSize ServiceName,DisplayName,State,StartMode
        }
    }
}
