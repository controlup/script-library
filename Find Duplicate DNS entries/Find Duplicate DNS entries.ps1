##
## build-dns-objects.ps1
##

##
## Michael B. Smith
## michael (at) TheEssentialExchange.com
## April, 2012
##
## Patched 2016-02-03
## If a computer had multiple IP addresses, the scanner would get confused and
## not properly populate the $objects array.
##

##
## Primary functionality:
##
## Based on either an input file or the output of a default command:
##
## 	dnscmd ( $env:LogonServer ).SubString( 2 ) /enumrecords $env:UserDnsDomain "@"
##
## Create an array containing all of the DNS objects describing the input.
##
## ----
##
## Secondary functionality:
##
## Find all the duplicate IP addresses and the duplicate names 
## contained within either the file or the command output.
##
## By specifying the -skipRoot option, all records for the root of
## the domain are ignored.
##

##
## General record format returned by DNScmd.exe:
##
##	name
##	[aging:xxxxxxxx] 
##	TTL
##	resource-record-type
##	value
##	[optional additional values]
##
## Fields may be separated by one-or-more spaces or one-or-more tabs
## [aging:xxxxxxxx] fields are optional
##

[CmdletBinding(SupportsShouldProcess=$false, ConfirmImpact='None')]

Param(
	[string]$filename,
	[switch]$skipRoot=$true
)

Set-StrictMode -Version 2.0

function new-dns-object
{
	return ( "" | Select Name, Aging, TTL, RRtype, Value )
}

function tmpFileName
{
	[string] $strFile = ( Join-Path $Env:Temp ( Get-Random ) ) + ".txt"
	Write-Verbose "tmpFileName $strFile"

	if( ( Test-Path -Path $strFile -PathType Leaf ) )
	{
		rm $strNetworkFile -EA 0
		if( $? )
		{
		##	write-output "...file was deleted"
		}
		else
		{
		##	write-output "...couldn't delete file, error: $($error[0].ToString())"
		}
	}

	return $strFile
}

# verify dnscmd.exe exists locally before running the script
If (! (Test-Path (Join-Path (Split-Path $env:comspec) "dnscmd.exe") )) {
    Write-Host "dnscmd.exe does not exist on this computer. The script cannot continue."
    Exit 1
} Else {

    if( $filename -and ( $filename.Length -gt 0 ) )
    {
    	$tmp = $filename
    }
    else
    {
    	$tmp = tmpFileName
    	dnscmd ( $env:LogonServer ).SubString( 2 ) /enumrecords $env:UserDnsDomain "@" >$tmp
    }

    $objects = @()
    $records = gc $tmp
    Write-Verbose "records = $( $records.Count )"

    $priorName = ''

    ## Primary functionality:

    foreach( $record in $records )
    {
    	## Write-Debug "Processing: $record"

    	if( !$record )
    	{
    		continue
    	}
    	if( $record -eq "Returned records:" )
    	{
    		continue
    	}
    	if( $record -eq "Command completed successfully." )
    	{
    		continue
    	}

    	if( $record -match "SOA" )
    	{
    		continue
    	}

    	$firstChar = $record.SubString( 0, 1 )
    	$record = $record.Trim()

    	if( $record.Length -eq 0 )
    	{
    		continue
    	}

    	$object = new-dns-object

    	$index = 0

    	$record = $record.Replace( "`t", " " )
    	$record = $record.Replace( "  ", " " )
    	$record = $record.Replace( "  ", " " )
    	$record = $record.Replace( "  ", " " )
    	Write-Debug "'$record'"

    	$array = $record.Split( ' ' )
    	Write-Debug "array contains $( $array.Count ) elements"
    	if( $array.Count -gt 5 )
    	{
    		Write-Warning "This record has been parsed incorrectly: '$record'"
    	}

    	if( ( $firstchar -eq " " ) -or ( $firstchar -eq "`t" ) )
    	{
    		$object.Name = $priorName
    		Write-Debug "Assigned priorName '$priorName'"
    	}
    	else
    	{
    		$object.Name = $array[ 0 ]
    		$priorName   = $array[ 0 ]
    		$index++
    	}

    	if(($array[$index].Length -ge 3) -and ($array[ $index ].SubString( 0, 3 ) -eq "[Ag"))  ## [Aging:3604987]
    	{
    		$object.Aging = $array[ $index ]
    		$index++
    	}
     
    	$object.TTL    = $array[ $index ]
    	$object.RRType = $array[ $index + 1 ]
    	$object.Value  = $array[ $index + 2 ]

    	$objects += $object

    	Write-Debug $object
    }

    ## Secondary functionality:

    ## There are more efficient ways to do this, but this is easy.

    ## search for duplicate names

    Write-Host "Duplicates for $env:UserDnsDomain :"
    $hash = @{}
    $duplicates = 0
    foreach( $o in $objects )
    {
    	if( $o.RRtype -eq "A" )
    	{
    		$name = $o.Name
    		if( $skipRoot -and ( $name -eq "@" ) )
    		{
    			continue
    		}
    		if( $hash.ContainsKey( $name ) )
    		{
    			"Duplicate name: $name, IP: $($o.Value), original IP: $($hash[ $name ])"
                $duplicates++
    		}
    		else
    		{
    			$hash[ $name ] = $o.Value
    		}
    	}
    }
    $hash = $null

    ## search for duplicate IP addresses

    $hash = @{}
    foreach( $o in $objects )
    {
    	if( $o.RRtype -eq "A" )
    	{
    		if( $skipRoot -and ( $o.Name -eq "@" ) )
    		{
    			continue
    		}

    		$ip = $o.Value
    		if( $hash.ContainsKey( $ip ) )
    		{
    			"Duplicate IP: $ip, name: $($o.Name), original name: $($hash[ $ip ])"
                $duplicates++
    		}
    		else
    		{
    			$hash[ $ip ] = $o.Name
    		}
    	}
    }
    $hash = $null

    #$global:DNSobjects = $objects
    
    If ($duplicates -eq "0") {
        Write-Host "No duplicates found."
    }
    " "
    "Done"
}

