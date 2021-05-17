$hostname = $args[0]

function get-sid
{
    Param ( $DSIdentity )
    $ID = new-object System.Security.Principal.NTAccount($DSIdentity)
    return $ID.Translate( [System.Security.Principal.SecurityIdentifier] ).toString()
}
$admin = get-sid "Administrator"

Write-Output "Computer SID = $($admin.SubString(0, $admin.Length - 4))"
Write-Output "Domain SID = $(get-sid $hostname$)"
