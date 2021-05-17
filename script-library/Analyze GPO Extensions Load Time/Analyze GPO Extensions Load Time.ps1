
<#
.SYNOPSIS

Show the durations of GPO Client Side Extension processing by matching events for the specified user in the Group Policy event log

.DETAILS

Find the latest event id 4001 ("Starting user logon Policy processing for username") for the specified user and use its activity id to find the 4016 ("Starting xxx Extension Processing") and 5016 ("") events so we can report on these

.PARAMETER user

The domain qualified name of the user being reported on

.EXAMPLE

& '.\Analyze GPO Extensions Load Time'  contoso\bob

Show a break down of the timings of GPO client side extension processing for the last logon of user bob in domain contoso

.CONTEXT

Session

.NOTES

Uses code used in Analyze Logon Durations script

.MODIFICATION_HISTORY:

    @guyrleech 06/11/2020  Initial release
    @guyrleech 10/11/2020  Changed dates/times output to just times
    @guyrleech 13/11/2020  Changed so could run in PowerShell 2.0
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='domain\username to report on')]
    [string]$user
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputwidth = 400

if( [string]::IsNullOrEmpty( $user ) -or $user.IndexOf( '\' ) -lt 0 )
{
    Throw 'Username must be in domain\username format'
}

if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}
## Code below all from Analyze Logon Durations script

## Find event id 4001 from GP log so we can get activity id to cross ref to 5016 event for finishing of GPO processing

[string]$query = "*[EventData[Data[@Name='PrincipalSamName'] and (Data='$user')]] and *[System[(EventID='4001')]]"
$CSEArray = $null
[hashtable]$CSE2GPO = @{}

if( $startProcessingEvent = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $query -MaxEvents 1 -ErrorAction SilentlyContinue ) ## most recent
{
    $query = "*[System[(EventID='4016' or EventID='5016' or EventID='6016' or EventID='7016') and TimeCreated[@SystemTime>='$($startProcessingEvent.TimeCreated.ToUniversalTime().ToString("s")).$($startProcessingEvent.TimeCreated.ToUniversalTime().ToString("fff"))Z'] and Correlation[@ActivityID='{$($startProcessingEvent.ActivityID.Guid)}']]]"
    if( ! ( $CSEarray = @( Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $query -ErrorAction SilentlyContinue ) ) -or ! $CSEArray.Count )
    {
        Throw "Failed to find any group policy event id 5016 instances for CSE finishes"
    }
    else
    {
        ## build hash table of cse id and GPO names so we can output when we iterate over finish events later
        $CSEArray | Where-Object { $_.Id -eq 4016 } | ForEach-Object `
        {
            $CSE2GPO.Add( $_.Properties[0].Value , $_.Properties[5].Value )
        }
    }
}
else
{
    Throw "Failed to find group policy processing starting event id 4001 for user $user"
}
if( $CSEArray -and $CSEArray.Count )
{
    $lastToFinish = $null
    [hashtable]$GPOTotalTimes = @{}

    [array]$CSEtimings = @( $CSEArray | Where-Object { $_.Id -ne '4016' } | ForEach-Object `
    {
        $CSE = $_
        [double]$duration = $CSE.Properties[0].Value / 1000
        if( ! $lastToFinish -or $CSE.TimeCreated -gt $lastToFinish )
        {
            $lastToFinish = $CSE.TimeCreated
        }

        ## look up the list of GPOs via the CSE extension id from 4016 event we built earlier
        [string[]]$GPOs = @( $CSE2GPO[ $CSE.Properties[3].Value ] -split "`n" )
        ForEach( $GPO in $GPOs )
        {
            try
            {
                if( ! [string]::IsNullOrEmpty( $GPO.Trim() ) )
                {
                    $GPOTotalTimes.Add( $GPO , $duration )
                }
                ## else empty string so ignore
            }
            catch
            {
                ## already have it so we add to the time
                [double]$alreadyGot = $GPOTotalTimes.Get_Item( $GPO )
                $GPOTotalTimes.Set_Item( $GPO , $alreadyGot + $duration )
            }
        }

        New-Object -Typename pscustomobject -Property (@{
            CSE       = $CSE.Properties[2].Value
            StartTime = $CSE.TimeCreated.AddMilliseconds( -$CSE.Properties[0].Value ).ToLongTimeString()
            EndTime   = $CSE.TimeCreated.ToLongTimeString()
            Duration  = $duration 
            GPOs      = ($GPOs -join ', ').Trim( '[, ]') })
    } )
            
    if( $lastToFinish )
    {
        ("Group Policy processing started at {0} , overall duration {1:N2} seconds" -f (Get-Date -Date $startProcessingEvent.TimeCreated -Format G) , ( $lastToFinish - $startProcessingEvent.TimeCreated ).TotalSeconds )
    }

    $CSEtimings | Sort-Object -Property Duration -Descending | Format-Table -AutoSize
            
    if( $GPOTotalTimes -and $GPOTotalTimes.Count )
    {
        "$($GPOTotalTimes.Count) processed GPO CSEs sorted by the highest total processing time"
        $GPOTotalTimes.GetEnumerator() | Where-Object { $_.Name } | Sort-Object -Property Value -Descending | Format-Table -AutoSize -Property @{n='GPO';e={$_.Name}},@{n='Time Spent (s)';e={$_.Value}}
    }
}

