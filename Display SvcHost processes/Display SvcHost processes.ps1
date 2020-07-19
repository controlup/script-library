param (
    [string]$computer = "."
)
$results = (Get-WmiObject -Class Win32_Process -ComputerName $computer -Filter "Name='svchost.exe'" | % {
    $process = $_
    Get-WmiObject -Class Win32_Service -ComputerName $computer -Filter "ProcessId=$($_.ProcessId)" | % {
        New-Object PSObject -Property @{ProcessId=$process.ProcessId;
                                        CommittedMemory=$process.WS;
                                        PageFaults=$process.PageFaults;
                                        CommandLine=$_.PathName;
                                        ServiceName=$_.Name;
                                        State=$_.State;
                                        DisplayName=$_.DisplayName;
                                        StartMode=$_.StartMode}
    	}
       }
)
$results

