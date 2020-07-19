#required -version 3.0
<#
    Find all user profiles across the local machine selected and allow removal

    @GuyRLeech, 2018

    Modification history:

    09/06/2020 @guyrleech  Cater for OneDrive Files on Demand by showing space consumed and potentially
#>

$VerbosePreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
$DebugPreference = 'SilentlyContinue'

[int]$outputWidth = 400

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

if( $args.Count -ne 3 )
{
    Throw "Takes three arguments - last used (days), over size (MB) and delete (true/false)"
}

[int]$lastUsedDays = $args[0]
[int]$overSizeMB = $args[1]
[bool]$delete = $args[2] -eq 'true'
[int]$FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000
[long]$INVALID_FILE_SIZE = 0xFFFFFFFF
[string]$sortBy = 'Last Used'
[string[]]$excludeUsers = $null
[string[]]$includeUsers = $null #@( 'g[aeiou]' )
[switch]$excludeLocal = $false

[string]$logName = 'Microsoft-Windows-User Profile Service/Operational'

$columns = [System.Collections.Generic.List[String]]@('User Name','Full Name','Profile Path','Profile Size (MB)','Loaded','Roaming','Last Used','Last Local Login (days)','Last Local Logoff (days)','Last AD Login (days)','Account Disabled','Account Locked','Password Expired','Password Last Set','Last Bad Password' ) 

$nativeDefinitions = @'
    [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
    public static extern uint GetCompressedFileSizeW( string pFileName , ref uint lpFileSizeHigh );
'@

Add-Type -MemberDefinition $nativeDefinitions -Name 'kernel32' -Namespace 'win32' -UsingNamespace System.Text -Debug:$false

Function Calculate-FolderSize( [string]$folderName , [string]$sid , [string[]]$excludeUsers , [string[]]$includeUsers , [ref]$lastUsed , [ref]$potentialSize )
{
    ## can't do a Get-ChildItem -Recurse as can't seem to stop junction point traversal so do it manually
    [string]$username = if( $sid )
    {
        try
        {
            ([System.Security.Principal.SecurityIdentifier]($sid)).Translate([System.Security.Principal.NTAccount]).Value
        }
        catch
        {
            $null
        }
    }
    else
    {
        $null
    }

    ForEach( $includedUser in $includeUsers )
    {
        [bool]$found = $false
        if( $username -match $includedUser )
        {
            $found = $false
            break
        }
        if( ! $found )
        {
            return $null,$null
        }
    }
    ForEach( $excludedUser in $excludeUsers )
    {
        if( [string]::IsNullOrEmpty( $username ) -or $username -match $excludedUser )
        {
            return $null,$null
        }
    }
    $items = @( $folderName )
    [array]$files = @( While( $items )
    {
        $newitems = @( $items | Get-ChildItem -Force -ErrorAction SilentlyContinue ) ## | Where-Object { ! ( $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint ) }
        $newitems
        $items = @( $newitems.Where( { $_.Attributes -band [System.IO.FileAttributes]::Directory } ) ) ## -and ! ( $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint ) } ))
    })
    if( $files -and $files.Count )
    {
        $lastUsed.Value = (Get-Date).AddYears( -20 )
        [string]$ntuserdotdat = Join-Path -Path $folderName -ChildPath 'ntuser.dat'
        [uint64]$size = 0
        $potentialSize.Value = 0
        $files | . { Process `
        {
            [uint64]$expandedSize = 0
            if( $_.FullName -ne $ntuserdotdat -and $_.LastWriteTime -gt $lastUsed.Value )
            {
                $lastUsed.Value = $_.LastWriteTime
            }
            if( $_.PSObject.Properties -and $_.PSObject.Properties[ 'Length' ] )
            {
                $expandedSize = $_.Length
            }
            if( ($_.Attributes.value__ -band $FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) -eq $FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS ) ## e.g. OneDrive Files On Demand
            {
                [uint64]$filesizeHigh = 0
                [uint64]$fileSizeLow = [win32.kernel32]::GetCompressedFileSizeW( $_.FullName , [ref] $filesizeHigh )
                if( $fileSizeLow -ne $INVALID_FILE_SIZE )
                {
                    [uint64]$actualsize = ( $filesizeHigh -shl 32 ) + $fileSizeLow
                    $size += $actualsize
                }
            }
            else
            {
                $size += $expandedSize
            }
            $potentialSize.Value += $expandedSize
        }}
        $size
    }
    else
    {
        [long]0
    }
    $username
}

[hashtable]$profileObjects = @{}
[long]$totalSize = 0
[datetime]$overAge = (Get-Date).AddDays( -$lastUsedDays )
[bool]$notFullyDownloaded = $false
[uint64]$totalPotentialSize = 0

[array]$userProfiles = @( Get-WmiObject -Class win32_userprofile -ErrorAction SilentlyContinue | Where-Object Special -ne $true | . { Process `
{
    $profile = $_
    [datetime]$lastUsed = Get-Date
    
    try
    {
        $lastUsed = [Management.ManagementDateTimeConverter]::ToDateTime( $profile.LastUseTime )
    }
    catch{}

    ## TODO Rule out on age before we calculate size??

    ## Get size of profile, last used - translate SID on remote machine in case a local account
    [uint64]$potentialSize = 0
    [uint64]$spaceUsed,[string]$username = Calculate-FolderSize -folderName $profile.LocalPath -sid $profile.sid -excludeUsers $excludeUsers -includeUsers $includeUsers -lastUsed ([ref]$lastUsed) -potentialSize ([ref]$potentialSize)
    [uint64]$sizeMB = 0
    if( $spaceUsed )
    {
        $sizeMB = [math]::Round( $spaceUsed / 1MB ) -as [int]
    }
    if( ! [string]::IsNullOrEmpty( $username ) -and ( $sizeMB -ge $overSizeMB -or $lastUsed -lt $overAge ) ) ## username could be null if excluded by called function
    {
        [string]$domainname,[string]$unqualifiedUserName = ( $username -split '\\' )
        ## Look in user profile log for last logon
        $lastLocalLogon = '-'
        $lastLogonTime = Get-WinEvent -FilterHashtable @{ LogName = $logName ; id = 1 ; UserId = $profile.SID } -ErrorAction SilentlyContinue | Select -First 1 -ExpandProperty TimeCreated
        if( $lastLogonTime )
        {
            $lastLocalLogon = [math]::round( (New-TimeSpan -End ([datetime]::Now)  -Start $lastLogonTime).TotalDays , 1 )
        }
        $lastLocalLogoff = '-'
        $lastLogoffTime = Get-WinEvent -FilterHashtable @{ LogName = $logName ; id = 4 ; UserId = $profile.SID } -ErrorAction SilentlyContinue | Select -First 1 -ExpandProperty TimeCreated
        if( $lastLogoffTime )
        {
            $lastLocalLogoff = [math]::round( (New-TimeSpan -End ([datetime]::Now)  -Start $lastLogoffTime).TotalDays , 1 )
        }
        if( $spaceUsed -lt $potentialSize )
        {
            $notFullyDownloaded = $true ## we will add an extra column showing the fully downloaded size
        }
        $totalPotentialSize += $potentialSize
        [hashtable]$properties = @{ 'User Name' = $username ; 'Profile Path' = $profile.LocalPath ; 'Profile Size (MB)' = $sizeMB ; 'Fully Downloaded Size (MB)' = [int]($potentialSize / 1MB)
            'Last Used' = $lastUsed ; 'Roaming' = $profile.RoamingConfigured ; 'Loaded' = $profile.Loaded ; 'Last Local Login (days)' = $lastLocalLogon ; 'Last Local Logoff (days)' = $lastLocalLogoff }
        $totalSize += $properties[ 'Profile Size (MB)' ]
        ## if $unqualifiedUserNmame is null then not a domain\username so won't be in AD
        if( ! [string]::IsNullOrEmpty( $unqualifiedUserName ) -and ! [string]::IsNullOrEmpty( $domainname ) )
        {
            ## we stuff the profile object into a separate hash table so we can call its delete method later if required. 
            $profileObjects.Add( $username , $profile )
            $user = [ADSI]"WinNT://$domainname/$unqualifiedUserName,user"
            if( $user -and $user.Path )
            {
                $lastLogin = 'Never'
                try
                {
                    $lastLogin = [math]::round( (New-TimeSpan -End ([datetime]::Now)  -Start $user.LastLogin.Value).TotalDays , 1 )
                }
                catch{}

                $properties += @{
                    'Full Name' = $user.FullName.Value
                    ##'Description' = $user.Description.Value
                    'Last AD Login (days)' = $lastLogin
                    ##'Password Last Changed' = (Get-Date).AddSeconds( -($user.PasswordAge.Value) )
                    'Password Expired' = if( $user.PasswordExpired )  { 'Yes' } else { 'No' }
                    'Account Disabled' = if( ( $user.UserFlags.Value -band 0x02 ) )  { 'Yes' } else { 'No' }
                    'Account Locked' = if( ( $user.UserFlags.Value -band 0x10 ) ) { 'Yes' } else { 'No' }
                    ##'Bad Passwords' = $user.BadPasswordAttempts.Value
                }
            }
        }
        [pscustomobject]$properties
    }
}})

if( $userProfiles -and $userProfiles.Count )
{
    [string]$header = "Found $($userProfiles.Count) user profiles either not used in the last $lastUsedDays days or in excess of $($overSizeMB)MB which in total are consuming $($totalSize)MB"
    if( $notFullyDownloaded )
    {
        $header += ", fully downloaded size would be $([int]($totalPotentialSize / 1MB))MB"
    }
    Write-Output -InputObject $header

    [datetime]$oldestEvent = Get-WinEvent -FilterHashtable @{ LogName = $logName ; id = 1,4 } -ErrorAction SilentlyContinue -MaxEvents 1 -Oldest | Select -ExpandProperty TimeCreated
    if( $oldestEvent )
    {
        "Oldest recorded logon/logoff event is from $(Get-Date $oldestEvent -Format G) ($([int](New-TimeSpan -Start $oldestEvent -End (Get-Date)).TotalDays) days ago)"
    }
    if( $notFullyDownloaded )
    {
        For( [int]$index = 0 ; $index -lt $columns.Count ; $index++ )
        {
            if( $columns[ $index ] -eq 'Profile Size (MB)' )
            {
                $columns.Insert( $index + 1 , 'Fully Downloaded Size (MB)' )
                break
            }
        }
    }
    $userProfiles | Select-Object -Property $columns | Sort-Object -Property $sortBy -Descending |Format-Table -AutoSize

    if( $delete )
    {
        [long]$totalSizeDeleted = 0
        [int]$deleted = 0
        $userProfiles | Where-Object { ! $_.Loaded } | ForEach-Object `
        {
            $profile = $_
            $profileObject = $profileObjects[ $profile.'User Name' ]
            if( $profileObject )
            {
                Write-Verbose "Deleting profile for $($profile.'User Name')"
                $profileObject.Delete()
                if( $? )
                {
                    $deleted++
                    $totalSizeDeleted += $profile.'Profile Size (MB)'
                }
            }
            else
            {
                Write-Warning "Failed to retrieve cached profile object for user $($profile.'User Name')"
            }
        }
        Write-Output "Deleted $deleted profiles occupying $totalSizeDeleted MB"
    }
}
else
{
    Write-Warning "No local user profiles found either not used in the last $lastUsedDays days or in excess of $($overSizeMB)MB"
}

