Function Test-AppVPublishingServer
{
    $Publishing_Servers = Get-AppvPublishingServer
    foreach ($Publishing_Server in $Publishing_Servers)
    {
        if (Test-Connection -ComputerName $Publishing_Server.Name -Count 2 -Quiet)
        {
            $status = "Reachable"
        }
        else
        {
            $status = "Un-Reachable"
        }
        $Publishing_Server | Add-Member -MemberType NoteProperty -Name "Status" -Value $status
    }
    $Publishing_Servers | sort -Property Name -Descending | Select Name,URL,Status
}

$ErrorActionPreference = "Stop"

If ( (Get-Module -Name AppvClient -ErrorAction SilentlyContinue) -eq $null )
{
        # using try/catch can stop the script completely if needed with "Exit with error" - 'Exit 1' (or some other non-zero exit code)
        # and avoid a long string of errors because the first statement was not successful.
        Try {
                Import-Module AppvClient
        } Catch {
                Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
                Exit 1
        }
}

Test-AppVPublishingServer

