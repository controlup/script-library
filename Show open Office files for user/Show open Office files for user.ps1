
<#
.SYNOPSIS
    Show open Excel workbook details by creating an Excel COM object in the user's session as that user

.PARAMETER officeApplication
    The Office application (process) to interrogate

.NOTES
    Modification History:

    2024/02/29  Guy Leech  Script born
    2024/05/27  Guy Leech  Added mechanism to cope with COM object timeout/hang
    2024/05/30  Guy Leech  Added mapping from Office application parameter to process name. Changed exception handler to avoid hangs if powershell in runspace still running
    2024/05/31  Guy Leech  Added more detail if fails to get Excel COM object as could be hung
#>

[CmdletBinding()]

Param
(
    [ValidateSet('excel','winword','powerpnt','word','powerpoint')]
    [string]$officeApplication = 'excel' ,
    [decimal]$jobTimeoutSeconds = 30
)

#region ControlUp_Standards
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
try
{
    if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
    {
        $WideDimensions.Width = $outputWidth
        $PSWindow.BufferSize = $WideDimensions
    }
}
catch
{
    ## not fatal
}
#endregion ControlUp_Standards

$officeAppInstance = $null

[hashtable]$officeProcessToComObject = @{
    'Excel'     = 'Excel.Application' 
    'Winword'   = 'Word.Application'
    'Powerpnt'  = 'Powerpoint.Application' 
    'Word'      = 'Word.Application'
    'Powerpoint'= 'Powerpoint.Application'
}

[hashtable]$officeProductNameToProcess = @{
    'Word'      = 'winword'
    'Powerpoint'= 'Powerpnt'
}

[string]$officeComObject = $officeProcessToComObject[ $officeApplication ]
$PowerShell = $null
$Runspace = $null

if( [string]::IsNullOrEmpty( $officeComObject ) )
{
    Throw "Unsupported Office process $officeApplication"
}

[int]$thisSessionId = Get-Process -id $pid | Select-Object -ExpandProperty SessionId

[string]$officeProcessName = $officeProductNameToProcess[ $officeApplication ]
if( [string]::IsNullOrEmpty( $officeProcessName ) )
{
    $officeProcessName = $officeApplication
}
Write-Verbose -Message "Looking for process $officeProcessName in session id $thisSessionId"

$officeProcesses = $null
$officeProcesses = Get-Process -Name $officeProcessName -ErrorAction SilentlyContinue | Where-Object SessionId -eq $thisSessionId

if( $null -eq $officeProcesses )
{
    Throw "No $officeProcessName processes found in session $thisSessionId"
}

if( $officeProcesses -is [array] -and $officeProcesses.Count -ne 1 )
{
    Throw "There are multiple ($($officeProcesses.Count)) $officeApplication processes in session $thisSessionId"
}

try
{
    #PowerShell jobs cause Excel to hang and the job times out even though it does seem to get the object, probably as runs in separate process
    ## https://devblogs.microsoft.com/scripting/beginning-use-of-powershell-runspaces-part-1/
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions = 'ReuseThread'
    $PowerShell = [powershell]::Create()
    $PowerShell.runspace = $Runspace
    $Runspace.Open()
    [void]$PowerShell.AddScript({
        Param( $officeComObject )
        
        [System.Diagnostics.Debug]::WriteLine("Pid $pid getting active object for $officeComObject" )
        $object = [Runtime.InteropServices.Marshal]::GetActiveObject( $officeComObject )
        [System.Diagnostics.Debug]::WriteLine("Pid $pid got active object $object for $officeComObject" )
        $object
    })
    [void]$PowerShell.AddParameters( @{
        'officeComObject' = $officeComObject
    } )
    $AsyncObject = $null
    $AsyncObject = $PowerShell.BeginInvoke()
    
    if( $null -eq $AsyncObject )
    {
        Throw "Failed to start runspace to get handle to Office process"
    }

    [bool]$waitResult = $false
    Write-Verbose -Message "$([datetime]::Now.ToString('G')): waiting for up to $jobTimeoutSeconds seconds for Office object"
    $waitResult = $AsyncObject.AsyncWaitHandle.WaitOne( $jobTimeoutSeconds * 1000 )
    Write-Verbose -Message "$([datetime]::Now.ToString('G')): back from wait, result is $waitResult"

    if( -Not $waitResult )
    {
        [string]$extraInfo = $null
        if( $null -ne $officeProcesses )
        {
            $extraInfo = ", $officeApplication (pid $($officeProcesses.Id)) could be hung - $([int]$officeProcesses.CPU) seconds of CPU consumed since started $([math]::Round( ([datetime]::Now - $officeProcesses.StartTime).TotalMinutes , 1)) minutes ago"
        }
        Throw "Failed to wait for job to get handle to Office process$extraInfo"
    }

    $officeAppInstance = $powershell.EndInvoke( $AsyncObject ) | Select-Object -First 1 ## can return a collection so we grab first and hopefully only object

    if( $null -ne $officeAppInstance )
    {
        # List all open Excel workbooks and their paths
        [int]$counter = 0
        if( $officeApplication -ieq 'excel' )
        {
            $collection = $officeAppInstance.Workbooks
            $activeDocumentName = $officeAppInstance.ActiveWorkbook.Name
            $activeSheetName = $officeAppInstance.ActiveSheet.Name
        }
        elseif( $officeApplication -imatch 'word$' )
        {
            $collection = $officeAppInstance.Documents
            $activeDocumentName = $officeAppInstance.ActiveDocument.Name
        }
        elseif( $officeApplication -imatch '^powerp' )
        {
            $collection = $officeAppInstance.Presentations
            $activeDocumentName = $officeAppInstance.ActivePresentation.Name
        }
        else
        {
            Write-Warning -Message "Unknown office application $officeApplication"
        }
        [int]$collectionCount = ($collection | Measure-Object).Count ## doesn't always have a count property
        if( $collectionCount -gt 0 )
        {
            $collection | ForEach-Object `
            {
                $counter++
                [string]$active = ' '
                if( $_.Name -ieq $activeDocumentName )
                {
                    $active = '*'
                }
                Write-Output -InputObject "$($counter)/$($collectionCount): $active $($_.Name) ($($_.FullName))"
            }
            if( $officeApplication -ieq 'excel' )
            {
                Write-Output -InputObject "`r`nActive sheet of `"$activeDocumentName`" is `"$activeSheetName`""
            }
        }
        else
        {
            Write-Output -InputObject "$officeApplication is running but there are no open $officeApplication files"
        }
    }
    else
    {
        Throw "Failed to connect to existing $officeApplication process id $($officeProcesses.Id) in session $thisSessionId"
    }
}
catch
{
    throw
}
finally
{
    if( $null -ne $officeAppInstance )
    {
        if( $waitResult )
        {
            ## no cleanup since we connected to an existing instance
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject( $officeAppInstance )
            $officeAppInstance = $null
        }
        ## else the wait failed, probably timed out so no cleanup required
    }
    if( $null -ne $PowerShell )
    {
        if( $PowerShell.InvocationStateInfo.State -ine 'running' )
        {
            $PowerShell.Dispose()
            $PowerShell = $null
        }
        ## else will hang if running and we try and stop it
    }
    if( $null -ne $Runspace )
    {
        if( $null -eq $PowerShell )
        {
            $runspace.Close()
            $runspace.Dispose()
        }
        else
        {
            $Runspace.CloseAsync()
        }
        $runspace = $null
    }
}
