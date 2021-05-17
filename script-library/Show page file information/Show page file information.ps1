#requires -version 3.0
<#
    Get page file configuration information and usage

    @guyrleech 2018
#>

[string]$regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
[int]$outputWidth = 400

$currentUsage = Get-CimInstance -ClassName Win32_PageFileUsage

$clearPageFileAtShutdown = Get-ItemProperty -Path $regKey -Name 'ClearPageFileAtShutdown' | Select -ExpandProperty 'ClearPageFileAtShutdown'
$disablePagingExecutive = Get-ItemProperty -Path $regKey -Name 'DisablePagingExecutive' | Select -ExpandProperty 'DisablePagingExecutive'

[array]$pagefileConfiguration = @( Get-ItemProperty -Path $regKey -Name 'PagingFiles' | Select -ExpandProperty 'PagingFiles' | ForEach-Object `
{
    ## automatic management will have no sizes after page file name, system managed will have 0 0 and manual will have initial maximum
    [bool]$customSize = $false
    [string]$configuration = $null
    [string[]]$components = $_ -split '\s+'
    if( ! $components -or $components.Count -le 1 )
    {
        $configuration = 'Automatically managed'
    }
    elseif( $components[1] -eq '0' -and $components[2] -eq '0' )
    {
        $configuration = 'System managed size'
    }
    else
    {
        $configuration = 'Custom Size'
        $customSize = $true
    }
    $result = [pscustomobject][ordered]@{
        'Pagefile' = $components[0]
        'Configuration' = $configuration
        'Initial Size (MB)' = $null
        'Maximum Size (MB)' = $null
    }
    if( $customSize )
    {
        $result.'Initial Size (MB)' = $components[1]
        $result.'Maximum Size (MB)' = $components[2]
    }
    $currentState = $currentUsage | Where-Object { $_.Name -eq $components[0] -or $configuration -eq 'Automatically managed' }
    if( $currentState )
    {
        $result.Pagefile = $currentState.Name ## name from registry may not have drive letter
        Add-Member -InputObject $result -NotePropertyMembers @{
            'Current Usage (MB)' = $currentState.CurrentUsage
            'Limit (MB)' = $currentState.AllocatedBaseSize
            'Current Usage (%)' = [math]::Round( ($currentState.CurrentUsage / $currentUsage.AllocatedBaseSize ) * 100 )
            'Peak Usage (%)' = [math]::Round( ($currentState.PeakUsage / $currentUsage.AllocatedBaseSize ) * 100 )
        }
    }
    $result
} )

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( $disablePagingExecutive )
{
    Write-Warning "Paging is disabled"
}

if( $clearPageFileAtShutdown )
{
    Write-Warning "Page file is set to be cleared at shutdown"
}

"Page file configuration:"

$pagefileConfiguration | Format-Table -AutoSize

