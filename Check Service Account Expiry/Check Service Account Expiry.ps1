#requires -runasadministrator

<#
.SYNOPSIS
    Check Service Accounts for Account and Password expiry

.DESCRIPTION
    Check service accounts for services and scheduled tasks
    The use case is where services are run as domain accounts (rather than as managed service accounts)
    so local system and network service accounts will be recognised, but ignored

.NOTES
    Version:        1.0
    Author:         Bill Powell, based on an idea from Balint Oberrauch
    Creation Date:  2024-01-09  Bill Powell  Script created
    Updated:        2024-01-23  Bill Powell  Added cross-referencing to SCM registry. Removed unused code
#>


[CmdletBinding()]
param (
)


# get AD domain info
$ComputerSystem = Get-CimInstance Win32_ComputerSystem
if ($ComputerSystem.PartOfDomain -ne $true) {
    Write-Error "Computer $($ComputerSystem.Name) is not part of a domain"
    exit 0
}

Function FixOutBufferSize {
    # Altering the size of the PS Buffer
    Write-Debug "In Function FixOutBufferSize"
    [int]$outputWidth = 800
    $PSWindow = (Get-Host).UI.RawUI
    $WideDimensions = $PSWindow.BufferSize
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

FixOutBufferSize

#region ActiveDirectory functions using ADSI

$script:SearcherLookup = @{}

function New-DomainSearcher {
    [CmdletBinding()]
    param (
        [parameter (mandatory = $true)][string]$DomainDN
    )
    $UserDomain = ([ADSI]"LDAP://$DomainDN")

    #
    # create a searcher
    $ADRootEntry = New-Object System.DirectoryServices.DirectoryEntry($UserDomain.Path)
    $DirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher
    $DirectorySearcher.SearchRoot = $ADRootEntry
    $DirectorySearcher.PageSize = 1000
    $DirectorySearcher.SearchScope = "Subtree"
    $script:SearcherLookup[$DomainDN] = $DirectorySearcher
}

#
# values from https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/useraccountcontrol-manipulate-account-properties
$ACCOUNTDISABLE       = 0x000002
$DONT_EXPIRE_PASSWORD = 0x010000
$PASSWORD_EXPIRED     = 0x800000

function Get-ADItemField {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        $Item, 
        [string]$FieldName
    )
    $propertyCollection = $Item.Properties
    $field = $propertyCollection.Item($FieldName)
    [string]$field
}

$script:DomainLookup = @{}

function Get-FullyQualifiedDomainDN {
    [CmdletBinding()]
    param (
        [parameter (mandatory = $true)][string]$Domain
    )
    $DomainLC = $Domain.ToLower()
    $FullyQualifiedDomainDN = $script:DomainLookup[$DomainLC]
    if ($null -eq $FullyQualifiedDomainDN) {
        $DomainInfo = [adsi]"LDAP://${DomainLC}"
        [string]$FullyQualifiedDomainDN = $DomainInfo.distinguishedName[0]
        $script:DomainLookup[$DomainLC] = $FullyQualifiedDomainDN
    }
    $FullyQualifiedDomainDN
}

$script:AccountCache = @{}

function Get-CachedAdAccount {
    [CmdletBinding()]
    param (
        [parameter (mandatory = $true)][string]$AccountName,
        [parameter (mandatory = $true)][string]$Domain
    )
    $FullyQualifiedDomainDN = Get-FullyQualifiedDomainDN -Domain $Domain
    $AccountKey = $AccountName + '|' + $FullyQualifiedDomainDN
    $SearchResult = $script:AccountCache[$AccountKey]
    if ($null -eq $SearchResult) {
        $DirectorySearcher = $script:SearcherLookup[$FullyQualifiedDomainDN]
        if ($DirectorySearcher -eq $null) {
            New-DomainSearcher -Domain $FullyQualifiedDomainDN
            $DirectorySearcher = $script:SearcherLookup[$FullyQualifiedDomainDN]
        }
        $userFilter = "(&(|(objectCategory=User)(objectClass=msDS-GroupManagedServiceAccount)(objectClass=msDS-ManagedServiceAccount))(|(sAMAccountName=$($AccountName))(userprincipalname=$($AccountName))))"               # users are specified by sAMAccountName
        $DirectorySearcher.Filter = $userFilter 
        $SearchResult = $DirectorySearcher.FindOne()   # should be 0 or 1 result
        $script:AccountCache[$AccountKey] = $SearchResult
    }
    $SearchResult
}

#endregion

function Check-PasswordExpiry {
    [CmdletBinding()]
    param (
        [parameter(mandatory = $true)][string]$SAMAccountName
    )
    $Info = [pscustomobject]@{
            'PasswordLastSet'        = 'Unset'
            'PasswordExpires'        = 'Unset'
            'PasswordChangeableFrom' = 'Unset'
            'PasswordRequired'       = 'Unset'
            'PasswordMayBeChanged'   = 'Unset'
    }
    $output = & net 'user' $SAMAccountName '/domain'
    $output | Where-Object {$_ -like "*password*"} | ForEach-Object {
        $Line = $_
        if ($Line -match "^Password last set\s+(?<time>\S.*\S)\s*$") {
            $FieldName = 'PasswordLastSet'
        }
        elseif ($Line -match "^Password expires\s+(?<time>\S.*\S)\s*$") {
            $FieldName = 'PasswordExpires'
        }
        elseif ($Line -match "^Password changeable\s+(?<time>\S.*\S)\s*$") {
            $FieldName = 'PasswordChangeableFrom'
        }
        elseif ($Line -match "^Password required\s+(?<flag>\S.*\S)\s*$") {
            $FieldName = 'PasswordRequired'
        }
        elseif ($Line -match "^User may change password\s+(?<flag>\S.*\S)\s*$") {
            $FieldName = 'PasswordMayBeChanged'
        }
        if ($Matches.time) {
            if ([string]$Matches.time -as [DateTime]) {
                $Info.$FieldName = [datetime]::Parse($Matches.time)
            }
            else {
                $Info.$FieldName = $Matches.time
            }
        }
        elseif ($Matches.flag) {
            $Info.$FieldName = $Matches.flag
        }
    }
    $Info
}

$CommonServiceFields = "ServiceType,Identifier,ServiceGroup,State,Path,Enabled,SID,Authority,AccountName,AccountType,AccountRaw,DistinguishedName,Description,AccountExpires,PasswordExpires" -split ','

function New-CommonService {
    [CmdletBinding()]
    param (
        [parameter(mandatory = $true)][ValidateSet('ScheduledTask', 'Service')][String]$ServiceType
    )
    New-Object psobject | Select-Object -Property $CommonServiceFields | ForEach-Object {
        $_.ServiceType = $ServiceType
        $_
    }
}

$script:SIDInfoCache = @{}

function Get-AccountDetails {
    [CmdletBinding()]
    param (
        [parameter(mandatory = $true)][string]$SID
    )
    $UserInfo = $script:SIDInfoCache[$SID]
    if ($null -eq $UserInfo) {
        $UserInfo = [pscustomobject]@{
            SID = $SID
            Authority = $null
            AccountName = $null
            AccountType = 'Unknown'
            DistinguishedName = ''
            FullName = ''
            ExpirationDate = 'Unknown'
        }
        Write-Verbose "Look up SID $SID"
        $AccountName = ([System.Security.Principal.SecurityIdentifier]($SID)).Translate([System.Security.Principal.NTAccount]).Value
        if ($AccountName -match "^(?<auth>[^\\]*)\\(?<acc>[^\\]*)") {
            $UserInfo.Authority = $Matches.auth
            $UserInfo.AccountName = $Matches.acc
        }
        if ($SID -notmatch "^S-1-5-\d\d$") {
            #
            # not a built-in user
            try {
                $DomainUser = [adsi]"LDAP://<SID=$SID>"
                $DomainUser | Out-Null
                $ObjectClass = $DomainUser.objectClass | Select-Object -Last 1
                $UserInfo.FullName = [string]$DomainUser.DisplayName
                if ($ObjectClass -like "*ManagedServiceAccount*") {
                    $UserInfo.AccountType = $ObjectClass
                }
                else {
                    $UserInfo.AccountType = 'ActiveDirectory'
                }
                $UserInfo.DistinguishedName = [string]$DomainUser.distinguishedName
            }
            catch {
                $exception = $_
                $exception | Out-Null
            }
            if ([string]::IsNullOrWhiteSpace($UserInfo.DistinguishedName)) {
                try {
                    $LocalUser = [adsi]"WinNT://./$($UserInfo.AccountName),user"
                    $UserInfo.FullName = [string]$LocalUser.Description
                    $UserInfo.AccountType = 'Local'
                }
                catch {
                    $exception = $_
                    $exception | Out-Null
                }
            }
        }
        else {
            $UserInfo.AccountType = 'BuiltIn'
            $UserInfo.ExpirationDate = 'Never'
        }
        $script:SIDInfoCache[$SID] = $UserInfo
    }
    $UserInfo
}

#
# get scheduled tasks
function Get-AllTaskSubFolders ([Parameter(Mandatory=$true)]$FolderRef) {
    $FolderRef
    $FolderRef.GetFolders(1) | ForEach-Object {
        $ChildFolder = $_
        Get-AllTaskSubFolders -FolderRef $ChildFolder
    }
}

$AllTasksAndServices = New-Object System.Collections.Generic.List[PSObject]

function Get-AllScheduledTasks {
    [CmdletBinding()]
    param ($computername = "localhost",
           [switch]$RootFolder
    ) 
    try {
	    $script:Schedule = New-Object -ComObject ("Schedule.Service")
    } catch {
	    Write-Warning "Schedule.Service COM Object not found, this script requires this object"
	    return
    }
    if ($RootFolder) {
        $script:RecurseTree = $false
    }

    $script:Schedule.Connect($ComputerName)
    $Root = $script:Schedule.GetFolder('\')
    $AllFolders = Get-AllTaskSubFolders -FolderRef $Root

    foreach ($Folder in $AllFolders) {
        if (($Tasks = $Folder.GetTasks(1))) {
            $Tasks | Foreach-Object {
                $Task = $_
                $TaskXml = [xml]($Task.Xml)
                $State = 'Unknown'
                $Triggers = $null
                try {
                    $State = switch ($Task.State) {
                        0 {'Unknown'}
                        1 {'Disabled'}
                        2 {'Queued'}
                        3 {'Ready'}
                        4 {'Running'}
                        Default {'Unknown'}
                    }
                }
                catch {
                    $exception = $_
                }
                $CSTask = New-CommonService -ServiceType ScheduledTask
                $SID = $TaskXml.Task.Principals.Principal.UserID
                if ([string]::IsNullOrWhiteSpace($SID)) {
                    $CSTask.SID = 'UNKNOWN'
                    $CSTask.AccountName = 'UNKNOWN'
                }
                else {
                    $CSTask.SID = $SID
                    $UserInfo = Get-AccountDetails -SID $SID
                    $CSTask.AccountName = $UserInfo.AccountName
                    $CSTask.AccountType = $UserInfo.AccountType
                    $CSTask.Authority   = $UserInfo.Authority
                }
                $CSTask.Identifier     = $Task.name
                $CSTask.Path           = $Task.path
                $CSTask.State          = '' 
                $CSTask.Enabled        = '' 
                $CSTask.Description    = $TaskXml.Task.RegistrationInfo.Description 
                $CSTask.AccountExpires = ''
                $AllTasksAndServices.Add($CSTask) 

            }
        }
    }
}

Get-AllScheduledTasks | Out-Null

#region Get Service info from registry

$RequiredProperties = 'Description,DisplayName,ImagePath,ObjectName,Start,Type,Owners,PSPath,PSParentPath,PSChildName,PSProvider' -split ','

$ServicesByDescription = @{}
$ServicesByServiceName = @{}

# Fetching and filtering service accounts from SCM (Service Control Manager)
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Services"
Get-ChildItem -Path $regPath | 
  ForEach-Object {
    $RegItem = $_
    $RegItemProperties = Get-ItemProperty $RegItem.PSPath -Name $RequiredProperties -ErrorAction SilentlyContinue
    if ($RegItemProperties -ne $null) {
        $RegItemProperties | Out-Null
        if (-not [string]::IsNullOrWhiteSpace($RegItemProperties.ObjectName) -or $true) {
            if ($RegItem.PSChildName -ne $RegItemProperties.PSChildName) {
                $RegItem | Out-Null
            }
            $ServiceInfoSplat = [pscustomobject]@{
                ServiceName = $RegItemProperties.PSChildName
                AccountName = $RegItemProperties.ObjectName
                Origin =      'Registry'
                Description = $RegItemProperties.Description
                DisplayName = $RegItemProperties.DisplayName
                ImagePath   = $RegItemProperties.ImagePath
                ObjectName  = $RegItemProperties.ObjectName
            }
            $ServicesByServiceName[$ServiceInfoSplat.ServiceName] = $ServiceInfoSplat
            if (-not [string]::IsNullOrWhiteSpace($ServiceInfoSplat.Description)) {
                $ServicesSharingDescription = $ServicesByDescription[$ServiceInfoSplat.Description]
                if ($ServicesSharingDescription -eq $null) {
                    $ServicesSharingDescription = @($ServiceInfoSplat)
                }
                else {
                    $ServicesSharingDescription += $ServiceInfoSplat
                    #
                    # now is a good moment to copy across the ObjectName
                    [string[]]$ObjectNames = $ServicesSharingDescription.ObjectName | Where-Object {-not [string]::IsNullOrWhiteSpace($_)} | Sort-Object -Unique
                    switch ($ObjectNames.Count) {
                        0 {}
                        1 {
                                foreach ($ServiceInfoSplat in $ServicesSharingDescription) {
                                    if ([string]::IsNullOrWhiteSpace($ServiceInfoSplat.AccountName)) {
                                        $ServiceInfoSplat.AccountName  = $ObjectNames[0]
                                    }
                                    if ([string]::IsNullOrWhiteSpace($ServiceInfoSplat.ObjectName)) {
                                        $ServiceInfoSplat.ObjectName  = $ObjectNames[0]
                                    }
                                }
                            }
                        default {
                                Write-Error "multiple accounts ($($ObjectNames -join ',')) specified for $($ServiceInfoSplat.ServiceName)"
                            }
                    }
                }
                $ServicesByDescription[$ServiceInfoSplat.Description] = $ServicesSharingDescription
            }
        }
    }
}

#endregion

$AllServicesWmi = Get-CimInstance -ClassName win32_service

$script:DomainAuthorities = @{}

$AllServicesWmi | ForEach-Object {
    $ServiceWMI = $_
    if ([string]::IsNullOrWhiteSpace($ServiceWMI.StartName)) {
        $ServiceInfoSplat = $ServicesByServiceName[$ServiceWMI.Name]
        $ServiceAccountName = $ServiceInfoSplat.AccountName
        if ([string]::IsNullOrWhiteSpace($ServiceAccountName)) {
            Write-Warning "Empty AccountName for $($ServiceInfoSplat.ServiceName)"
        }
    }
    else {
        $ServiceAccountName = $ServiceWMI.StartName
    }
    $CSTask = New-CommonService -ServiceType Service
    if ($ServiceAccountName -eq 'localSystem') {
        $CSTask.AccountName = $ServiceAccountName
        $CSTask.AccountType = 'Service Control Manager'
        $CSTask.Authority = $ServiceWMI.SystemName
    }
    elseif ($ServiceAccountName -match "^(?<auth>[^\\]*)\\(?<acc>[^\\]*)") {
        $CSTask.Authority = $Matches.auth
        $CSTask.AccountName = $Matches.acc
        $CSTask.AccountRaw = $ServiceAccountName
        $ntAccount = New-Object System.Security.Principal.NTAccount($CSTask.Authority, $CSTask.AccountName)
        $SID = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        $CSTask.SID = $SID
        if (-not [string]::IsNullOrWhiteSpace($SID)) {
            $UserInfo = Get-AccountDetails -SID $SID
            $CSTask.AccountName = $UserInfo.AccountName
            $CSTask.AccountType = $UserInfo.AccountType
            $CSTask.Authority   = $UserInfo.Authority
        }
        else {
            $ServiceAccountName | Out-Null
        }
        if ($CSTask.AccountType -ne 'BuiltIn') {
            $script:DomainAuthorities[$CSTask.Authority] = $true
        }
    }
    elseif ($ServiceAccountName -match "^(?<acc>[^@]*)\@(?<dom>[^@]*)") {
        $CSTask.Authority = $Matches.dom
        $CSTask.AccountName = $Matches.acc
        $CSTask.AccountRaw = $ServiceAccountName
        $ntAccount = New-Object System.Security.Principal.NTAccount($CSTask.Authority, $CSTask.AccountName)
        $SID = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        $CSTask.SID = $SID
        if (-not [string]::IsNullOrWhiteSpace($SID)) {
            $UserInfo = Get-AccountDetails -SID $SID
            $CSTask.AccountName = $UserInfo.AccountName
            $CSTask.AccountType = $UserInfo.AccountType
            $CSTask.Authority   = $UserInfo.Authority
        }
        else {
            $ServiceAccountName | Out-Null
        }
        if ($CSTask.AccountType -ne 'BuiltIn') {
            $script:DomainAuthorities[$CSTask.Authority] = $true
        }
    }
    elseif ($ServiceAccountName -match "^(?<acc>[^$]*)\$") {
        # Managed service account
        $CSTask.Authority = $Matches.dom
        $CSTask.AccountName = $Matches.acc
        $CSTask.AccountRaw = $ServiceAccountName
        $ntAccount = New-Object System.Security.Principal.NTAccount($CSTask.Authority, $CSTask.AccountName)
        $SID = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        $CSTask.SID = $SID
        if (-not [string]::IsNullOrWhiteSpace($SID)) {
            $UserInfo = Get-AccountDetails -SID $SID
            $CSTask.AccountName = $UserInfo.AccountName
            $CSTask.AccountType = $UserInfo.AccountType
            $CSTask.Authority   = $UserInfo.Authority
        }
        else {
            $ServiceAccountName | Out-Null
        }
    }
    else {
        $CSTask.Path = $ServiceWMI.PathName
        if ($ServiceWMI.PathName -match "^\S+svchost.exe\s*-k (?<sg>\S+)( -[a-jl-z].*){0,1}$") {
            $CSTask.ServiceGroup   = $Matches.sg
        }
        $ServiceInfoSplat = $ServicesByServiceName[$ServiceWMI.Name]
        $ServiceAccountName | Out-Null
    }
    $CSTask.Description = $ServiceWMI.Description
    $CSTask.Identifier = $ServiceWMI.Name
    $CSTask.Enabled = $ServiceWMI.Status
    $CSTask.State = $ServiceWMI.State
    $AllTasksAndServices.Add($CSTask) 
    $CSTask | Out-Null
}

#
# need to get account expiry for only those entries that use domain accounts

$DomainName = [string]$UserDomain.name

#
# get the MaxPwdAge from the domain object
# see https://serverfault.com/questions/58720/powershell-how-do-i-query-pwdlastset-and-have-it-make-sense
#     https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/323750
$MaxPwdAgeVal = $UserDomain.ConvertLargeIntegerToInt64($UserDomain.maxPwdAge.value)
if ($MaxPwdAgeDays -eq [long]::MinValue) {
    $MaxPwdAgeDays = 42   # no authoritative source for this
}
else {
    $MaxPwdAgeDays = $MaxPwdAgeVal / -864000000000
}

function ConvertFrom-FileTime {
    [CmdletBinding()]
    [OutputType([string])]
    param (        
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true)][string]$FileTime,
        [switch]$AsDateTime
    )
    if ($FileTime -eq $null) {
        "Null"
    }
    elseif ($FileTime -eq [long]::MaxValue) {
        "Never"
    }
    elseif ($FileTime -eq 0) {
        "Unset"
    }
    else {
        $DateTime = ([datetime]::FromFileTime($FileTime))
        if ($AsDateTime) {
            $DateTime
        }
        else {
            $DateTime.ToString("yyyy-MM-dd HH:mm:ss")
        }
    }
}

#
# generate list of valid (domain) authorities discovered
[string[]]$script:DomainAuthorityList = $script:DomainAuthorities.GetEnumerator() | ForEach-Object {
    $_.Key
}

[object[]]$DomainAccountServices = $AllTasksAndServices | Where-Object {$_.Authority -in $script:DomainAuthorityList} | ForEach-Object {
    $CSTask = $_
    [object[]]$SearchResult = Get-CachedAdAccount -AccountName $CSTask.AccountName -Domain $CSTask.Authority
    $SearchResult | ForEach-Object {
        $Item = $_
        if ($Item -ne $null) {
            [int]$userAccountControl = Get-ADItemField -Item $Item -FieldName "userAccountControl"
            $PasswordExpired         = [bool]($userAccountControl -band $PASSWORD_EXPIRED)
            $PasswordNeverExpires    = [bool]($userAccountControl -band $DONT_EXPIRE_PASSWORD)
            $AccountExpires          = Get-ADItemField -Item $Item -FieldName "accountexpires"     | ConvertFrom-FileTime
           # $badpasswordtime         = Get-ADItemField -Item $Item -FieldName "badpasswordtime"    | ConvertFrom-FileTime
           # [datetime]$pwdlastset    = Get-ADItemField -Item $Item -FieldName "pwdlastset"         | ConvertFrom-FileTime -AsDateTime
           # $lastlogontimestamp      = Get-ADItemField -Item $Item -FieldName "lastlogontimestamp" | ConvertFrom-FileTime
           # $lastlogon               = Get-ADItemField -Item $Item -FieldName "lastlogon"          | ConvertFrom-FileTime
            if ($CSTask.AccountType -notlike "*ManagedServiceAccount*") {
                $CSTask.AccountExpires   = $AccountExpires
                $PasswordInfo = Check-PasswordExpiry -SAMAccountName $CSTask.AccountName
                $CSTask.PasswordExpires  = $PasswordInfo.PasswordExpires
            }
            else {
                $CSTask.AccountExpires   = 'N/A'
                $CSTask.PasswordExpires  = 'N/A'
            }
        }
        else {
            $Item | Out-Null
        }
    }
    $CSTask
}

$DomainAccountServices | Format-Table -Property Authority,AccountName,AccountType,AccountExpires,PasswordExpires,ServiceType,Identifier,Path,Description


Write-Output "$($DomainAccountServices.Count) services / scheduled tasks found running under domain accounts"

