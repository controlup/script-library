#requires -version 3.0

<#
    Retrieve all user accounts which have Citrix Studio admin access

    Guy Leech

    Modification History:
    
    31/10/18  GRL  Add unidentified accounts to the output
#>

[string]$userName = $null
if( $args.Count -and $args[0] )
{
    $userName = $args[0]
}

[int]$outputWidth = 400
[hashtable]$global:groups = @{}

Function Get-GroupMembers( $admin , $group )
{
    $group.psbase.invoke('members') | ForEach-Object `
    {
        $adspath = $_.GetType().InvokeMember( 'ADSPath' ,  'GetProperty',  $null,  $_, $null)
        $class = $_.GetType().InvokeMember( 'Class' ,  'GetProperty',  $null,  $_, $null)

        if( $class -eq 'group' )
        {
            try
            {
                $groups.Add( $adspath , $global:groups )
                Get-GroupMembers -group ([ADSI]"$adspath,$class") -admin $admin
            }
            catch
            {
                Write-Warning "Group $adspath is nested recursively"
            }
        }
        elseif( $class -eq 'user' )
        {
            $accountbits = $adspath -split '/'
            Get-UserProperties -name "$($accountbits[-2])\$($accountbits[-1])" -user ([ADSI]"$adspath,$class") -viaGroup $group.Name.Value -admin $admin
        }
    }
}

Function Get-UserProperties( $name , $admin , $user , $viaGroup )
{
    $lastLogin = if( $user )
    {
        try
        {
            [math]::round( (New-TimeSpan -End ([datetime]::Now)  -Start $user.LastLogin.Value).TotalDays , 1 )
        }
        catch
        {
            'Never'
        }
    }
    else
    {
        '?'
    }

    [pscustomobject][ordered]@{ 
        'Name' = $name 
        'Full Name' = $user.FullName.Value
        'Admin Group' = $viaGroup 
        'Admin Enabled' = if( $admin.Enabled ) { 'Yes' } else { 'No' }
        'Rights' = $admin.Rights -join ','
        ##'Description' = $user.Description.Value
        'Last AD Login (days)' = $lastLogin
        ##'Password Last Changed' = (Get-Date).AddSeconds( -($user.PasswordAge.Value) )
        'Password Expired' = if( $user) { if( $user.PasswordExpired )  { 'Yes' } else { 'No' } } else { '?' }
        'Account Disabled' = if( $user) { if( ( $user.UserFlags.Value -band 0x02 ) )  { 'Yes' } else { 'No' }} else { '?' }
        'Account Locked' =   if( $user) { if( ( $user.UserFlags.Value -band 0x10 ) ) { 'Yes' } else { 'No' }} else { '?' }
        ##'Bad Passwords' = $user.BadPasswordAttempts.Value
    }
}

Add-PSSnapin 'Citrix.DelegatedAdmin.Admin.*' -ErrorAction Stop

## can't search at this level for account as will only work if assigned at the to level, not via a group
[array]$admins = @( Get-AdminAdministrator -ErrorAction Stop )

$results = @( ForEach( $admin in $admins )
{
    $user = $null
    $group = $null

    [string]$domain,[string]$account = $admin.Name -split '\\'
    
    $group = [ADSI]"WinNT://$domain/$account,group"
    if( ! $group -or ! $group.PSObject.properties[ 'Path' ] )
    {
        $group = $null
        $user = [ADSI]"WinNT://$domain/$account,user"
    }

    if( $group )
    {
        Get-GroupMembers -admin $admin -group $group 
    }
    elseif( $user -and $user.PSObject.properties[ 'Path' ]  )
    {
        Get-UserProperties -admin $admin -user $user -name $admin.Name
    }
    else
    {
        Write-Warning "Unable to find account entity `"$($admin.Name)`""
        Get-UserProperties -admin $admin -user $null -name $admin.Name
    }
}) | Where-Object { $_.Name -match $userName } 

# Altering the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = $outputWidth
$PSWindow.BufferSize = $WideDimensions

[string]$summary = "Got $($results.count) individual admins from $($admins.Count) entries"
if( ! [string]::IsNullOrEmpty( $userName ) )
{
    $summary += " matching `"$username`""
}

$summary
$results | Sort 'Rights' | Format-Table -AutoSize

