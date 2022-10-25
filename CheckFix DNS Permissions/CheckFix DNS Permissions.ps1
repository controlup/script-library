<#
    Find Active Directory DNS records with unresolved SIDs and replace with the machine account

    @guyrleech 2019

    Based on code from https://www.heelpbook.net/2018/powershell-find-and-add-dns-record-permissions/
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Enter computer name or pattern to check DNS record(s) for')]
    [string]$computername ,    
    [Parameter(Mandatory=$false,HelpMessage='Whether to fix issues found (True or False)')]
    [ValidateSet('Yes','No')]
    [string]$fixParameter = 'No'
)

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

[bool]$fix = $( if( $fixParameter -eq 'Yes' ) { $true } else { $false } )

Import-Module -Name ActiveDirectory -ErrorAction Stop -Verbose:$false

[string[]]$reservedNames = @( 'DomainDnsZones' , 'ForestDnsZones' , '@' )
[int]$fixedCount = 0
$thisDomain = Get-ADDomain
[string]$DomainName = $thisDomain.DNSroot
[string]$AdIntegrationType = 'Domain'
[string]$DomainDn = $thisDomain.DistinguishedName

## get domain controllers so we can avoid lest we break anything big style
[hashtable]$domainControllers = @{}
Get-ADGroupMember -Identity 'Domain Controllers' | ForEach-Object `
{
     $dc = $_
     $domainControllers.Add( $dc.DistinguishedName , $dc.SID )
}

[string]$adpath = "AD:\DC=$DomainName,CN=MicrosoftDNS,DC=$AdIntegrationType`DnsZones,$DomainDn"

Write-Verbose "AD path is `"$adpath`""

## exclude _ldap _kerberos _dc etc
[array]$results = @( Get-ChildItem $adpath | Where-Object {  $_.Name -like $computername -and $_.Name[0] -ne '_' -and $_.Name.IndexOf('.') -lt 0 -and $reservedNames -notcontains $_.Name } | ForEach-Object `
{
    $record = $_
    Write-Verbose -Message "Checking $($record.Name)"
    try
    {
        $computer = Get-ADComputer -Identity $record.Name -ErrorAction SilentlyContinue -ErrorVariable ComputerError
    }
    catch
    {
        $computer = $null ## may be a manual entry so don't report as missing
    }
    if( $computer )
    {
        if( $domainControllers[ $computer.DistinguishedName ] )
        {
            Write-Warning "$($computer.Name) is a domain controller so skipping due to risk"
        }
        else
        {
            [string]$ADPath = "ActiveDirectory:://RootDSE/$($record.DistinguishedName)"
            $ACL = Get-Acl -Path $ADPath -ErrorAction SilentlyContinue
            if( $ACL )
            {
                $result = New-Object -TypeName PSObject
                Add-Member -InputObject $result -MemberType NoteProperty -Name 'Computer' -Value $computer.Name
                [string]$machineAccount = $((($computer.DNSHostName -split '\.')[1..0]) -join '\') + '$'
                [int]$foundSelf = 0
                [bool]$aclChanged = $false
                [int]$unresolvedSIDs = 0

                if( $ACL.Owner -ne $machineAccount )
                {
                    Write-Verbose -Message "Found incorrect owner $($ACL.Owner) on $($computer.Name)"
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Incorrect Owner' -Value $ACL.Owner
                    if( $fix )
                    {
                        $ACL.SetOwner( [System.Security.Principal.NTAccount]$machineAccount )
                        $aclChanged = $true
                    }
                }
                ForEach( $ACE in $ACL.Access )
                {
                    if( $ACE.IdentityReference -eq $machineAccount )
                    {
                        $foundSelf++
                    }
                    elseif( $ACE.IdentityReference -match '^S\-1\-5\-21\-\d{10}-\d{10}-\d{10}-\d+$' )
                    {
                        Write-Verbose -Message "Found unresolved SID $($ACE.IdentityReference) in ACL for $($computer.Name)"
                        $unresolvedSIDs++
                        if( $fix )
                        {
                            $removal = $ACL.RemoveAccessRule( $ACE )
                            if( $removal )
                            {
                                $aclChanged = $true
                            }
                            else
                            {
                                Write-Warning "Failed to remove $($ACE.IdentityReference) from ACE on $($computer.Name)"
                            }
                        }
                    }
                }
                if( $unresolvedSIDs )
                {
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Unresolved SIDs' -Value $unresolvedSIDs
                }
                if( ! $foundSelf )
                {
                    Write-Verbose -Message "Failed to find $machineAccount in ACL for $($computer.Name)"
                    Add-Member -InputObject $result -MemberType NoteProperty -Name 'Missing Machine Account' -Value 'Yes'
                    if( $fix )
                    {
                        $Acl.AddAccessRule( ( New-Object System.DirectoryServices.ActiveDirectoryAccessRule( $Computer.Sid , 'GenericAll', 'Allow' ) ) )
                        $aclChanged = $true
                    }
                }
                if( ( $result.PSObject.Properties.GetEnumerator()|Measure-Object|Select-Object -ExpandProperty Count ) -gt 1 )
                {
                    $result
                }
                if( $aclChanged -and $fix )
                {
                    $newACL = Set-Acl -Path $ADPath -AclObject $Acl -Passthru
                    if( $newACL )
                    {
                        $fixedCount++
                    }
                }
            }
            else
            {
                Write-Warning "Failed to get ACL for $($record.distinguishedName)"
            }
        }
    }
})

if( $results -and $results.Count )
{
    [string]$status = "Found $($results.Count)"
    if( $fixedCount )
    {
        $status += " and fixed $fixedCount"
    }
    Write-Output -InputObject "$status DNS record permission issues:"
    $results | Format-Table -AutoSize
}
else
{
    Write-Output -InputObject 'Found no DNS record permission issues'
}

