<#
    Get each CPU's usage over a period and show least loaded

    @guyrleech 2018

    Modification History:

    20/11/18   GRL  Always dispplay two decimal places so pad with zeroes if necessary
#>

[int]$samplePeriod = $args[0]

[hashtable]$perCpuStats = @{}
[datetime]$first = 0
[bool]$gotFirst = $false
[datetime]$last = 0

Get-Counter -Counter '\Processor(*)\% Processor Time' -SampleInterval 1 -MaxSamples $samplePeriod | ForEach-Object `
{
    if( ! $gotFirst )
    {
        $first = $_.TimeStamp
        $gotFirst = $true
    }
    $last = $_.TimeStamp
    $_.CounterSamples | ForEach-Object `
    {
        $thisCPU = $perCpuStats[ $_.InstanceName ]
        if( $thisCpu )
        {
            $null = $thisCpu.Add( $_.CookedValue )
        }
        else
        {
            $perCpuStats.Add( $_.InstanceName , [System.Collections.ArrayList]@( $_.CookedValue ) )
        }
    }
}

"Average CPU utilisation over $samplePeriod seconds"

if( $perCpuStats -and $perCpuStats.Count )
{
    [array]$results = @( $perCpuStats.GetEnumerator() | ForEach-Object `
    {
        if( $_.Key -ne '_total' )
        {
            [pscustomobject][ordered]@{
                'CPU' = $_.Key -as [int]
                'Average CPU Utilisation %' = ( $_.Value | Measure-Object -Average | Select -ExpandProperty Average ).ToString("0.00")
                ##'Minimum CPU Utilisation %' = [math]::Round( ( $_.Value | Measure-Object -Minimum | Select -ExpandProperty Minimum ) , 2 )
                ##'Maximum CPU Utilisation %' = [math]::Round( ( $_.Value | Measure-Object -Maximum | Select -ExpandProperty Maximum ) , 2 )
            }
        }
    })
    $results | Sort CPU | Format-Table -AutoSize
}
else
{
    Write-Error "Failed to retrieve any CPU performance data"
}

