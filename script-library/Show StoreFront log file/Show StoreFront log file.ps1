#requires -version 3
<#
    Take Citrix (almost) XML format StoreFront logs and convert to csv or grid view

    @guyrleech, 2018

    Modification History:
#>

## Arguments are:
##  0  Output file (mandatory)
##  1  Start date/time or last x (defaults to 1 day)
##  2  End date/time (defaults to now)

[string]$logsFolder = "$($env:ProgramFiles)\Citrix\Receiver StoreFront\Admin\trace" 

if( ! ( Test-Path -Path $logsFolder -PathType Container ) )
{
    [string]$exceptionText = "StoreFront log folder $logsFolder does not exist."
    [string]$product = 'Citrix StoreFront'
    $installDir = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match $product } | Select -First 1 -ExpandProperty InstallLocation
    if( [string]::IsNullOrEmpty( $installDir ) )
    {
        $exceptionText += ' StoreFront installation not found.'
    }
    Throw $exceptionText
}

[string]$outputFile = $null

if( $args.Count -ge 1 -and $args[0] )
{
    $outputFile = $args[0]
    if( Test-Path -Path $outputFile )
    {
        Throw "Output file `"$outputFile`" already exists, cannot overwrite"
    }
    elseif( ! ( Test-Path -Path (Split-Path -Path $outputFile -Parent) -PathType Container ) )
    {
        Throw "Folder for output file `"$outputFile`" does not exist"
    }
}

if( [string]::IsNullOrEmpty( $outputFile ) )
{
    Throw "No output file specified"
}

$startDate = $null
$endDate = $null

if( $args.Count -ge 2 -and $args[1] )
{
    ## This could be specified as a date/time or last days/minutes/hours/etc so figure out which
    $result = New-Object DateTime
    if( [datetime]::TryParse( $args[1] , [ref]$result ) )
    {
        $startDate = $result
    }
    else
    {
        ## see what last character is as will tell us what units to work with
        [string]$last = $args[1]
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
        $endDate = Get-Date
        if( $last.Length -le 1 )
        {
            $startDate = $endDate.AddSeconds( -$multiplier )
        }
        else
        {
            $startDate = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}
else
{
    $startDate = (Get-Date).AddDays( -1 )
    $endDate = Get-Date
}

if( $args.Count -ge 3 -and $args[2] )
{
    ## This could be specified as a date/time or a duration of days/minutes/hours/etc so figure out which
    $result = New-Object DateTime
    if( [datetime]::TryParse( $args[2] , [ref]$result ) )
    {
        $endDate = $result
    }
    else
    { 
        [string]$last = $args[2]
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
            $endDate = $startDate.AddSeconds( $multiplier )
        }
        else
        {
            $endDate = $startDate.AddSeconds( ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
        }
    }
}

Function Process-LogFile
{
    Param( [string]$inputFile , [datetime]$start , [datetime]$end )

    [string]$parent = 'GuyLeech' ## doesn't matter what it is since it only ever lives in memory

    Write-Verbose "Processing $inputFile ..." 
    ## Missing parent context so add
    [xml]$wellformatted = "<$parent>" + ( Get-Content $inputFile ) + "</$parent>"

    <#
    <EventID>0</EventID>
		    <Type>3</Type>
		    <SubType Name="Information">0</SubType>
		    <Level>8</Level>
		    <TimeCreated SystemTime="2016-05-25T14:59:12.8077640Z" />
		    <Source Name="Citrix.DeliveryServices.WebApplication" />
		    <Correlation ActivityID="{00000000-0000-0000-0000-000000000000}" />
		    <Execution ProcessName="w3wp" ProcessID="52168" ThreadID="243" />
    #>

    try
    {
        $wellformatted.$parent.ChildNodes | Where-Object { [datetime]$_.System.TimeCreated.SystemTime -ge $start -and [datetime]$_.System.TimeCreated.SystemTime -le $end } | ForEach-Object `
        {
            $result = [pscustomobject]@{'Date'=[datetime]$_.System.TimeCreated.SystemTime; 
                'File' = (Split-Path $inputFile -Leaf); 
                'Type'=$_.System.Type;
                'Subtype'=$_.System.Subtype.Name;
                'Level'=$_.System.Level;
                'SourceName'=$_.System.Source.Name;
                'Computer'=$_.System.Computer;
                'ApplicationData'=$_.ApplicationData}
            ## Not all objects have all properties
            if( ( Get-Member -InputObject $_.system -Name 'Process' -Membertype Properties -ErrorAction SilentlyContinue ) )
            {
                $result | Add-Member -MemberType NoteProperty -Name 'Process' -Value $_.System.Process.ProcessName
                $result | Add-Member -MemberType NoteProperty -Name 'PID' -Value $_.System.Process.ProcessID;`
                $result | Add-Member -MemberType NoteProperty -Name 'TID' -Value $_.System.Process.ThreadID;`
            }
            $result
        } | Select Date,File,SourceName,Type,SubType,Level,Computer,Process,PID,TID,ApplicationData ## This is the order they will be displayed in
    }
    catch {}
}

[int]$filesCount = 0

[array]$results = @( Get-ChildItem -Path $logsFolder | Where-Object { $_.LastWriteTime -ge $startDate } | ForEach-Object `
{
    Process-LogFile -inputFile $_.FullName -start $startDate -end $endDate
    $filesCount++
})

if( $results )
{
    "$($results.Count) log entries found in $filesCount files between $(Get-Date $startDate -Format U) and $(Get-Date $endDate -Format U), writing to $outputFile"

    $results | Sort Date | Export-Csv -Path $outputFile -NoTypeInformation -NoClobber
}
else
{
    Write-Warning "No entries found between $(Get-Date $earliest -Format G) and $(Get-Date $endDate -Format G) in $filesCount files searched"
}

