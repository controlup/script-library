<#
    Check for any automatic start services which aren't running and start them

    @guyrleech 2019
#>

[int]$outputWidth = 400

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

$DebugPreference = 'SilentlyContinue'

[array]$stoppedServices = @( Get-Service | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' } )

if( ! $stoppedServices -or ! $stoppedServices.Count )
{
    "No stopped, automatic start, services found"
}
else
{
    [string[]]$serviceNames = @( $stoppedServices | Select-Object -ExpandProperty Name )
    ## Look for services that we depend on that are also stopped and mark them since the start of dependent serivce will try and start them first so we don't need to
    [int]$removed = 0
    ForEach( $stoppedService in $stoppedServices )
    {
        $stoppedService.DependentServices | ForEach-Object `
        {
            if( $serviceNames -contains $_.Name )
            {
                Add-Member -InputObject $stoppedService -MemberType NoteProperty -Name 'DoNotStart' -Value $true
                $removed++
            }
        }            
    }

    [int]$errors = 0
    [int]$startedOk = 0
    "Found $($stoppedServices.Count - $removed) stopped services which should start automatically:"
    $stoppedServices | Where-Object { ! $_.DoNotStart } | Select DisplayName,@{n='Depends on';e={$_.ServicesDependedOn|select -ExpandProperty 'DisplayName'}}
    "Restarting ..."
    ForEach( $service in $stoppedServices )
    {
        if( ! $service.PSObject.Properties[ 'DoNotStart' ] -or ! $service.DoNotStart )
        {
            $startError = $null
            $started = Start-Service -InputObject $service -PassThru -ErrorAction SilentlyContinue -ErrorVariable StartError
            if( $startError )
            {
                ## cater for where a service starts but immediately stops, by design
                if( $startError[0] -and $startError[0].Exception -and $startError[0].Exception.InnerException )
                {
                    Write-Error "Failed to start service `"$($service.DisplayName)`" : $($startError[0].Exception.GetBaseException())"
                    $errors++
                }
            }
            else
            {
                $startedOk++
            }
        }
        else
        {
            Write-Debug "Skipping `"$($service.Name)`" as a dependency"
        }
    }
    "Restarting complete with $errors failures"
}

