<#
    Find user logon scripts for user from gpresult

    @guyrleech 2018
#>

[string]$user = $args[0]

if( [string]::IsNullOrEmpty( $user ) )
{
    Throw "Must pass the domain\username as an argument"
}

[string]$xmlFile = Join-Path $env:temp ( 'gpresult.' + ( $user -replace '\\' , '.' ) + '.' + $pid + '.xml' )
[string]$stdoutFile = Join-Path $env:temp ( 'gpresult.' + ( $user -replace '\\' , '.' ) + '.' + $pid + '.output.log' )
[int]$outputWidth = 400
[string]$localPolicyFolder = 'GroupPolicy\User\Scripts\Logon'

try
{
    # Altering the size of the PS Buffer
    $PSWindow = (Get-Host).UI.RawUI
    $WideDimensions = $PSWindow.BufferSize
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions

    if( Test-Path -Path $stdoutFile ) 
    {
        Remove-Item -Path $stdoutFile -Force -ErrorAction SilentlyContinue
    }

    $gpresult = Start-Process -FilePath 'gpresult.exe' -ArgumentList "/scope user /user $($args[0]) /f /x `"$xmlFile`"" -PassThru -Wait -NoNewWindow -RedirectStandardOutput $stdoutFile 

    if( ! $gpresult )
    {
        Throw "Failed to launch gpresult.exe process for user $user"
    }

    if( ! ( Test-Path -Path $xmlFile -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
    
        [string]$exceptionText = "gpresult failed to produce file `"$xmlFile`" for user $user"
        if( Test-Path -Path $stdoutFile -PathType Leaf -ErrorAction SilentlyContinue )
        {
            $exceptionText += ", error: $(Get-Content -Path $stdoutFile)"
        }
        Throw $exceptionText
    }

    [xml]$gpo = Get-Content -Path $xmlFile -ErrorAction Stop

    if( ! $gpo )
    {
        Throw "Failed to read gpresult XML file `"$xmlFile`""
    }

    $scripts = $gpo.rsop.UserResults.ExtensionData|? { $_.Name.'#text'-eq 'Scripts' }

    if( $scripts )
    {
        ## Each script has the GUID of the GPO from whence it came so we build a hash table of the GUIDS and names for the output
        [hashtable]$groupPolicies = @{}
        [hashtable]$domains = @{}
        $gpo.rsop.UserResults.GPO | Where-Object { $_.PSObject.properties[ 'Enabled' ] -and $_.Enabled -eq 'true' -and $_.PSObject.properties[ 'AccessDenied' ] -and $_.AccessDenied -ne 'true' } | ForEach-Object `
        {
            if( $_.Path.PSObject.properties[ 'Identifier' ] )
            {
                [string]$guid = $_.Path.Identifier|select -ExpandProperty '#text'
                $groupPolicies.Add( $guid , $_.Name )
            }
            [string]$domain = $null
            if( $_.Path.PSObject.properties[ 'Domain' ] )
            {
                $domain = $_.Path.Domain|select -ExpandProperty '#text'
                if( ! [string]::IsNullOrEmpty( $domain ) ) ## will be empty for local policies
                {
                    try
                    {
                        ## Get a list of domains so we can tell users the paths where scripts are located
                        $domains.Add( $domain , $true )
                    }
                    catch {}
                }
            }
        }
        [array]$output = @( $scripts | select -ExpandProperty Extension | select -ExpandProperty Script | Where-Object { $_.Type -eq 'Logon' }  | Sort -Property Order | ForEach-Object `
        {
            [string]$guid = $_.GPO.Identifier.'#text'
            [string]$scriptFolder = $null
            if( $guid -and $guid -eq 'LocalGPO' )
            {
                $scriptFolder = (Join-Path ([environment]::getfolderpath('system')) $localPolicyFolder )
            }
            if( ! [string]::IsNullOrEmpty( $guid ) )
            {
                [string]$size = 'Folder not found'
                [string]$domain = $null
                if( $_.GPO.PSObject.properties[ 'Domain' ] )
                {
                    $domain = $_.GPO.Domain.'#text'
                }

                [string]$lastModified = $null
                if( [string]::IsNullOrEmpty( $scriptFolder ) -and ! [string]::IsNullOrEmpty( $domain ) )
                {
                    $scriptFolder = "\\$domain\SYSVOL\$domain\Policies\$guid\User\Scripts\Logon"
                }
                if( Test-Path -Path $scriptFolder -PathType Container -ErrorAction SilentlyContinue )
                {
                    [string]$ScriptPath = Join-Path $scriptFolder $_.Command
                    if( Test-Path -Path $ScriptPath -PathType Leaf -ErrorAction SilentlyContinue )
                    {
                        $size = Get-ItemProperty -Path $ScriptPath -ErrorAction SilentlyContinue -Name Length|select -ExpandProperty Length
                        $lastModified = Get-Date -Date ( Get-ItemProperty -Path $ScriptPath -ErrorAction SilentlyContinue -Name LastWriteTime|select -ExpandProperty LastWriteTime ) -Format G
                    }
                    else
                    {
                        $size = 'File not found'
                    }
                }
                $object = [pscustomobject][ordered]@{
                    'GPO' = $groupPolicies[ $guid ]
                    'Script' = $_.Command
                    'Parameters' = $_.Parameters
                    'Size' = $size
                    'Last Modified' = $lastModified
                    'GUID' = $guid
                }
                if( $domains.Count -gt 1 )
                {
                    if( ! [string]::IsNullOrEmpty( $domain ) )
                    {
                        $object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $domain
                    }
                }
                $object
            }
        })
        
        $output | Format-Table -AutoSize

        if( $domains -and $domains.Count -gt 1 )
        {
            $domains.GetEnumerator()|ForEach-Object `
            {
                Write-Output "GPO logon scripts for domain $($_.Key) can be found in \\$($_.Key)\SYSVOL\$($_.Key)\Policies\<GUID>\User\Scripts\Logon"
            }
        }
        else
        {
            [string]$singleDomain = $domains.GetEnumerator() | Select -ExpandProperty Key -First 1
            Write-Output "GPO logon scripts are in \\$singleDomain\SYSVOL\$singleDomain\Policies\<GUID>\User\Scripts\Logon"
        }

        Write-Output "`nLocal logon scripts are in $(Join-Path ([environment]::getfolderpath('system')) $localPolicyFolder)"
    }
    else
    {
        Write-Output "No group policy logon scripts found"
    }
}
Catch
{
    Throw $_
}
Finally
{
    Remove-Item -Path $xmlFile -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $stdoutFile -Force -ErrorAction SilentlyContinue
}

