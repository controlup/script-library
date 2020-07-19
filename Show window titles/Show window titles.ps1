<#
    Show Window titles for processes. Must run as the user owning the session otherwise window information is not available

    @guyrleech 2019
#>

if( ! $args -or ! $args.Count )
{
    Throw 'Must pass the session id as the only parameter'
}

[int]$thisSessionId = $args[0] -as [int]

[int]$outputWidth = 400

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

[array]$processes = @( Get-Process | Where-Object {  $_.SessionId -eq $thisSessionId -and $_.MainWindowTitle } )

if( $processes -and $processes.Count )
{
    $processes | Format-Table -AutoSize -Property Id,ProcessName,MainWindowTitle,FileVersion,StartTime
}
else
{
    Write-Warning "No processes with window titles found in session id $thisSessionId"
}

