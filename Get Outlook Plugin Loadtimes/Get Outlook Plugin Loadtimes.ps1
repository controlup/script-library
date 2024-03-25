[datetime]$outlookStartTime = $args[0]


$Events = Get-EventLog -LogName Application -Source "Outlook" -ErrorAction SilentlyContinue | Where-Object {$_.EventID -eq 45 -and $_.TimeGenerated -gt $outlookStartTime}

Foreach($Event in $Events)
{
    $EventDescription = ($Event.Message) -split "`n"

    $result = New-Object 'System.Collections.Generic.List[string]'
    If($($EventDescription.Length) -gt 9)
    {
        #Check if the event entry is roughly the same as the process start time
        $interval = $outlookStartTime - $Event.TimeGenerated
        if ($interval.Minutes -ge 0 -and $interval.Minutes -lt 5)
        {
            $AddinsNumber = [int][Math]::Floor($EventDescription.Length / 9)
            Write-Host "Number of Addins: $AddinsNumber"
            $counter = 0
            Do
            {
                # Define custom object properties
                $customObjectProperties = @{
                    ComputerName   = $env:COMPUTERNAME
                    #BootTime       = $Event.TimeGenerated
                    PluginName     = ($EventDescription[3+$counter*9] -split ': ')[1]
                    #ProgID         = $($EventDescription[5+$counter*9] -split ': ')[1]
                    #GUID           = $($EventDescription[6+$counter*9] -split ': ')[1]
                    LoadBehavior   = $($EventDescription[7+$counter*9] -split ': ')[1]
                    #HKLM           = $($EventDescription[8+$counter*9] -split ': ')[1]
                    "Load Time (ms)"    = $($EventDescription[10+$counter*9] -split ': ')[1]
                }

                # Create a custom object
                $entry = New-Object PSObject -Property $customObjectProperties
                
                $result += $entry
                $counter++
            }Until($counter -eq $AddinsNumber)
        }
    }
    Write-Output $result | Select-Object PluginName, "Load Time (ms)", LoadBehavior |Format-Table
}

