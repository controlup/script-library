#requires -version 3
<#
    Retrieve Citrix Studio logs

    @guyrleech 2018
#>

[bool]$studioOnly = $true
[bool]$directorOnly = $false
[bool]$configChange = $true
[bool]$adminAction = $false

[string]$ddc = 'localhost' 
[string]$username = $null
[string]$operation = $null
[int]$outputWidth = 400
$start  = $null
$end = $null
[int]$maxRecordCount = 5000 

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( $studioOnly -and $directorOnly )
{
    Throw "Cannot specify both -studioOnly and -directorOnly"
}

if( $args.Count -ge 1 -and $args[0] )
{   
    $result = New-Object -TypeName DateTime
    if( [datetime]::TryParse( $args[0] , [ref]$result ) )
    {
        $start = $result
    }
    else
    {
        $last = $args[0]
        [long]$multiplier = 0
        switch( $last[-1] )
        {
            "s" { $multiplier = 1 }
            "m" { $multiplier = 60 }
            "h" { $multiplier = 3600 }
            "d" { $multiplier = 86400 }
            "w" { $multiplier = 86400 * 7 }
            "y" { $multiplier = 86400 * 365 }
            default { Throw "Unknown multiplier `"$($last[-1])`"" }
        }
        if( $last.Length -le 1 )
        {
            $start = (Get-Date).AddHours( -$multiplier )
        }
        else
        {
            $start = (Get-Date).AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}
else
{
    $start = (Get-Date).AddDays( -7 )
}

if( $args.Count -ge 2 -and $args[1] )
{   
    $result = New-Object DateTime
    if( [datetime]::TryParse( $args[1] , [ref]$result ) )
    {
        $end = $result
    }
    elseif( $args[1] -eq 'Now' )
    {
        $end = [datetime]::Now
    }
    else
    {
        $last = $args[1]
        [long]$multiplier = 0
        switch( $last[-1] )
        {
            "s" { $multiplier = 1 }
            "m" { $multiplier = 60 }
            "h" { $multiplier = 3600 }
            "d" { $multiplier = 86400 }
            "w" { $multiplier = 86400 * 7 }
            "y" { $multiplier = 86400 * 365 }
            default { Throw "Unknown multiplier `"$($last[-1])`"" }
        }
        if( $last.Length -le 1 )
        {
            $end = $start.AddHours( $multiplier )
        }
        else
        {
            $end = $start.AddSeconds( ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}
else
{
     $end = [datetime]::Now
}

if( $args.Count -ge 3 -and $args[2] )
{
    $username = $args[2]
}

Add-PSSnapin -Name 'Citrix.ConfigurationLogging.Admin.*'

if( ! ( Get-Command -Name 'Get-LogHighLevelOperation' -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find the Citrix Get-LogHighLevelOperation cmdlet required - is $($env:COMPUTERNAME) a Delivery Controller?"
}

Add-PSSnapin -Name 'Citrix.Broker.Admin.*'

$thisDeliveryController = try
{
    Get-BrokerController -ErrorAction SilentlyContinue | Where-Object { $_.MachineName -match "\\$($env:COMPUTERNAME)`$" } 
}
catch{}
    

if( ! $thisDeliveryController )
{
    Write-Warning 'This machine does not appear to be a delivery controller so there may be no results'
}

[hashtable]$queryparams = @{
    'AdminAddress' = $ddc
    'SortBy' = '-StartTime'
    'MaxRecordCount' = $maxRecordCount
    'ReturnTotalRecordCount' = $true
}
if( $configChange -and ! $adminAction )
{
    $queryparams.Add( 'OperationType' , 'ConfigurationChange' )
}
elseif( ! $configChange -and $adminAction )
{
    $queryparams.Add( 'OperationType' , 'AdminActivity' )
}
if( ! [string]::IsNullOrEmpty( $username ) )
{
    if( $username.IndexOf( '\' ) -lt 0 )
    {
        $username = $env:USERDOMAIN + '\' + $username
    }
    $queryparams.Add( 'User' , $username )
}
if( $directorOnly )
{
    $queryparams.Add( 'Source' , 'Citrix Director' )
}
if( $studioOnly )
{
    $queryparams.Add( 'Source' , 'Studio' )
}

$recordCount = $null

[array]$results = @( Get-LogHighLevelOperation -Filter { StartTime -ge $start -and EndTime -le $end }  @queryparams -ErrorAction SilentlyContinue -ErrorVariable RecordCount | ForEach-Object -Process `
{
    if( [string]::IsNullOrEmpty( $operation ) -or $_.Text -match $operation )
    {
        $result = [pscustomobject][ordered]@{
            'Started' = $_.StartTime
            ##'Duration (s)' = [math]::Round( (New-TimeSpan -Start $_.StartTime -End $_.EndTime).TotalSeconds , 2 )
            'User' = $_.User
            'From' = $_.AdminMachineIP
            'Operation' = $_.text
            ##'Source' = $_.Source
            ##'Type' = $_.OperationType
            'Successful' = $_.IsSuccessful
        }
        if( ! $configChange )
        {
            Add-Member -InputObject $result -NotePropertyMembers @{
                'Targets' = $_.TargetTypes -join ','
                'Target Process' = $_.Parameters[ 'ProcessName' ]
                'Target Machine' = $_.Parameters[ 'MachineName' ]
                'Target User' = $_.Parameters[ 'UserName' ]
            }
        }
        $result
    }
} )

if( $recordCount -and $recordCount.Count )
{
    if( $recordCount[0] -match 'Returned (\d+) of (\d+) items' )
    {
        if( [int]$matches[1] -lt [int]$matches[2] )
        {
            Write-Warning "Only retrieved $($matches[1]) of a total of $($matches[2]) items, use -maxRecordCount to return more"
        }
        ## else we got all the records
    }
    else
    {
        Write-Error $recordCount[0]
    }
}

[string]$dateRange = "between $(Get-Date $start -Format G) and $(Get-Date $end -Format G)"
[string]$message = if( ! [string]::IsNullOrEmpty( $username ) ) { " for user $username" }
if( ! $results -or ! $results.Count )
{
    Write-Warning "No events found$message $dateRange"
}
else
{
    Write-Output "Retrieved $($results.Count) events$message $dateRange"
    $results | Format-Table -AutoSize
}

