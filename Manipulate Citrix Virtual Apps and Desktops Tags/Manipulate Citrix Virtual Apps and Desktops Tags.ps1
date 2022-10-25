#requires -version 3.0

<#
.DESCRIPTION

    Set or unset a Citrix tag for the specified machine, optionally creating it.
    Can also delete tags or list current tag assignments

.PARAMETER computerName

    One or more CVAD computers to operate on

.PARAMETER tagName

    The name of the tag to operate with

.PARAMETER operation

    The operation to perform with the tag

.PARAMETER enableMaintenanceMode

    Enable maintenance mode on the specified computers

.PARAMETER disableMaintenanceMode

    Disable maintenance mode on the specified computers

.PARAMETER createTag

    Create the tag if it does not exist. Default behaviour is to error if the tag for a set/unset does not exist

.PARAMETER ddc

    The delivery controller to operate on

.PARAMETER tagDescription

    A description for the tag if it is to be created
    
.PARAMETER maxRecordCount

    Maximum number of records to fetch in one call from Citrix - default is 250 which may not be large enough

.EXAMPLE

    & '.\Manipulate Citrix tag.ps1' -computerName GLXA19PVS* -tagName dummy -operation list -ddc grl-xaddc02

    List the tags, power, registration and maintenance mode states on the specified CVAD machines matching the name GLXA19PVS* by connecting to the delivery controller grl-xaddc02
    The -tagname is required but is ignored

.EXAMPLE

    & '.\Manipulate Citrix tag.ps1' -computerName GLXA19PVS401,GLXA19PVS501,GLXA19PVS502 -tagName excluded -operation set -ddc grl-xaddc02 

    Assign the tag "excluded" to the 3 specified CVAD machines by connecting to the delivery controller grl-xaddc02. 
    
.EXAMPLE

    & '.\Manipulate Citrix tag.ps1' -computerName GLXA19PVS401,GLXA19PVS501,GLXA19PVS502 -tagName excluded -operation unset -ddc grl-xaddc02 -disableMaintenanceMode yes

    Remove the tag "excluded" from the 3 specified CVAD machines by connecting to the delivery controller grl-xaddc02 and take the machines out of maintenance mode
      
.EXAMPLE

    & '.\Manipulate Citrix tag.ps1' -computerName dummy -tagName oops -operation delete -ddc grl-xaddc02

    Remove the tag "oops" from the all CVAD machines and then delete it by connecting to the delivery controller grl-xaddc02
    
.NOTES

    Citrix PowerShell for CVAD must be available on the machine where the script runs

    User running the script must have sufficient privileges in CVAD to perform the required functions

    Modification History:

    2021/07/27  @guyrleech  Initial release
    2022/07/27  @guyrleech  Added -maxrecordcount
#>

<#
Copyright © 2021 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Computer to operate on')]
    [string[]]$computerName ,
    [Parameter(Mandatory=$true,HelpMessage='Name of tag')]
    [string]$tagName ,
    [Parameter(Mandatory=$true,HelpMessage='Operation to perform')]
    [ValidateSet('set','unset','list','delete')]
    [string]$operation ,
    [ValidateSet('yes','no')]
    [string]$enableMaintenanceMode ,
    [ValidateSet('yes','no')]
    [string]$disableMaintenanceMode ,
    [ValidateSet('yes','no')]
    [string]$createTag ,
    [string]$ddc ,
    [string]$tagDescription ,
    [int]$maxRecordCount = 5000
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputWidth = 400
if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

if( $enableMaintenanceMode -eq 'yes' -and $disableMaintenanceMode -eq 'yes' )
{
    Throw "Cannot enable and disable maintenance mode"
}

Add-PSSnapin -Name Citrix.Broker.Commands.* , Citrix.Broker.Admin.*

if( ! (Get-Command -Name Get-BrokerTag -ErrorAction SilentlyContinue ) )
{
    Throw "Get-BrokerTag cmdlet not found - are the Citrix PowerShell cmdlets installed?"
}

## array may have been flattened if called from outside PowerShell, eg scheduled task
if( $computerName -and $computerName.Count -eq 1 -and $computerName[0].IndexOf( ',' ) -ge 0 )
{
    $computerName = @( $computerName -split ',' )
}

$tag = $null

[hashtable]$ddcArgument = @{ }
if( $PSBoundParameters[ 'ddc' ] -and $ddc -ne 'localhost' -and $ddc -ne $env:COMPUTERNAME )
{
    $ddcArgument.Add( 'AdminAddress' , $ddc )
}

if( $operation -ne 'list' )
{
    if( ! ( $tag = Get-BrokerTag -Name $tagName -ErrorAction SilentlyContinue @ddcArgument ) )
    {
        if( $operation -eq 'delete' )
        {
            Throw "Tag `"$tagName`" not found so cannot delete"
        }
        if( [string]::IsNullOrEmpty( $createTag ) -or $createTag[0] -ine 'y' )
        {
            Throw "Tag `"$tagName`" not found and -createTag not set to `"yes`""
        }
        [hashtable]$newTagArguments = $ddcArgument.Clone()
        $newTagArguments.Add( 'Name' , $tagName )

        if( $PSBoundParameters[ 'tagDescription' ] -and ! [string]::IsNullOrEmpty( $tagDescription ) )
        {
            $newTagArguments.Add( 'Description' , $tagDescription )
        }
        Write-Verbose -Message "Creating new tag `"$tagName`""
        if( ! ( $tag = New-BrokerTag @newTagArguments ) )
        {
            Throw "Failed to create new tag `"$tagName`""
        }
    }
    elseif( $tag -is [array] )
    {
        Throw "Tag name `"$tagName`" is not unique, found $($tag.Count) matching tag names"
    }
}

if( $operation -eq 'delete' )
{
    Remove-BrokerTag -AllMachines -Tags $tag.Name
    Remove-BrokerTag -InputObject $tag
    [bool]$status = $?
    if( $status )
    {
        Write-Output -InputObject "Tag `"$($tag.Name)`" deleted ok"
    }
    exit $(if( $status ) { 0 } else { 42 }) ## cannot rely on $LASTEXITCODE
}

## get all broker machines and cache
[hashtable]$brokerMachines = @{}

Get-BrokerMachine @ddcArgument -MaxRecordCount $maxRecordCount | ForEach-Object `
{
    try
    {
        [string]$machineName = ($_.MachineName -split '\\')[-1]
        $brokerMachines.Add( $machineName , $_ )
    }
    catch
    {
        Write-Warning -Message "Duplicate machine name $machineName"
    }
}

Write-Verbose -Message "Got $($brokerMachines.Count) broker machines"

if( $computerName.Count -eq 1 -and $computerName[0].IndexOf( '*' ) -ge 0 )
{
    [string]$pattern = $computerName[0]
    $computerName = @( $brokerMachines.GetEnumerator() | Where-Object Name -Like $pattern | Select-Object -ExpandProperty Name | Sort-Object -Unique)
    Write-Verbose -Message "Found $($computerName.Count) machines matching $pattern"
}

[int]$counter = 0

[array]$results = @( ForEach( $computer in $computerName )
{
    $counter++
    if( $machine = $brokerMachines[ $computer ] )
    {
        Write-Verbose -Message "$counter / $($computerName.Count) : $($machine.Tags.Count) tags (`"$($machine.Tags -join '","')`") maintenance mode $($machine.InMaintenanceMode) registration state $($machine.RegistrationState) power state $($machine.PowerState) users $($machine.SessionCount)"
        [bool]$operationSuccess = $false

        switch( $operation )
        {
            'set'
            {
                if( $machine.Tags -contains $tag.Name )
                {
                    Write-Warning -Message "$computer already has tag $($tag.Name) set"
                }
                Add-BrokerTag -InputObject $tag -Machine $machine @ddcArgument
                $operationSuccess = $?
            }
            'unset'
            {
                if( $machine.Tags -notcontains $tag.Name )
                {
                    Write-Warning -Message "$computer does not already have tag $($tag.Name) set"
                }
                Remove-BrokerTag -InputObject $tag -Machine $machine @ddcArgument
                $operationSuccess = $?
            }
            'list'  { $machine }
        }

        if( $operationSuccess -and $operation -match 'set' )
        {
            if( $disableMaintenanceMode -eq 'yes' )
            {
                if( ! $machine.InMaintenanceMode )
                {
                    Write-Warning -Message "$($machine.MachineName) already not in maintenance mode"
                }
                else
                {
                    Write-Verbose -Message "Disabling maintenance mode on $($machine.MachineName)"
                    Set-BrokerMachine -InputObject $machine -InMaintenanceMode $false @ddcArgument
                }
            }
            elseif( $enableMaintenanceMode -eq 'yes' )
            {
                if( $machine.InMaintenanceMode )
                {
                    Write-Warning -Message "$($machine.MachineName) already in maintenance mode"
                }
                else
                {
                    Write-Verbose -Message "Enabling maintenance mode on $($machine.MachineName)"
                    Set-BrokerMachine -InputObject $machine -InMaintenanceMode $true @ddcArgument
                }
            }
        }
    }
    else
    {
        Write-Warning -Message "Failed to get machine $computer"
    }
})

if( $operation -eq 'list' )
{
    $results | Select-Object -Property @{n='Computer';e={($_.MachineName -split '\\')[-1]}},@{n='TagCount';e={$_.Tags.Count}},InMaintenanceMode,MaintenanceModeReason,RegistrationState,PowerState,SessionCount,@{n='Tags';e={$_.Tags -join ' , '}} | Format-Table -AutoSize
}
else
{
    $results ## not expecting any output
    Write-Output -InputObject "$operation completed"
}

Write-Verbose -Message "Processed $counter machines"

