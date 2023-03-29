<#
 	.SYNOPSIS
        Capture information using process monitor.
		
    .DESCRIPTION
        This script borrows heavily from the great work done by: 
        Nick Atkins - @Nik_41tkins - http://nomanualrequired.blogspot.com/

        The license he included with his original script:

            This program is free software: you can redistribute it and/or modify
            it under the terms of the GNU General Public License as published by
            the Free Software Foundation, either version 3 of the License, or
            (at your option) any later version.

            This program is distributed in the hope that it will be useful,
            but WITHOUT ANY WARRANTY; without even the implied warranty of
            MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
            GNU General Public License for more details.

            You should have received a copy of the GNU General Public License
            along with this program.  If not, see <http://www.gnu.org/licenses/>.

        This script will create a dynamic process monitor (procmon) filter using the session
        ID and start a capture lasting a duration you specify.
		
	.PARAMETER	<Duration <string[]>
		The duration (in seconds) to run the capture. Requires Procmon 3.2+ as versions from this point have the duration switch

	.PARAMETER  <ProcMonFolder <string[]>
		Specifies folder that contains procmon.exe.

    .PARAMETER  <SaveLogTo <string[]>
		Specifies folder the process monitor log file will be saved.

    .PARAMETER  <DropFilteredEvents <string>>
		When False DropFilteredEvents will keep all the events in the capture. If True then procmon will 'drop filtered events'. Dropping filtered events means faster procmon operation
        but you can't recover dropped events.

    .PARAMETER  <Filter <string>>
		Specify a filter to include with the capture. Best used with "$DropFilteredEvents = $True" to capture events surgically.
        Examples of filters:
        -Filter "Operation,is,WriteFile,include"
            -- Any process that does a 'WriteFile' is added to the capture
        -Filter "Process Name,contains,procmon,exclude;Process Name,is,Procexp.exe,exclude"
            -- applies two filters. First filter is excluding procmon from the capture, the second filter is excluding procexp from the capture.

    .PARAMETER  <DownloadProcmon <bool>>
		Download Procmon if it's not present on in the ProcMonFolder
		
    .LINK
        For more information refer to:
            http://github.com/trentent

    .LINK
        Stay in touch:
        http://twitter.com/trententtye

    .NOTES
        Warning!  Procmon logs may require disks that operate in the 100MB/sec to properly operate.
        TODO: TTYE - Implement Procmon altitude adjustment

    .EXAMPLE
        C:\PS> Remote-Procmon -Duration 15 -ProcMonFolder C:\sysinternals -SaveLogTo D:\procmonlogs
		
		Starts process monitor for 15 seconds.  Procmon.exe is located in the path
        C:\sysinternals\procmon.exe and the logs will be saved to the D:\procmonlogs folder.

        C:\PS> . .\TraceSystemActivity.ps1 15 C:\sysinternals D:\procmonlogs

    .EXAMPLE
        C:\PS> Remote-Procmon -Duration 15 -ProcMonFolder C:\sysinternals -SaveLogTo D:\procmonlogs -DropFilteredEvents:$True -Filter "Process Name,contains,procmon,exclude"
		
		Starts process monitor for 15 seconds, excluding all events with the process name "procmon".  Procmon.exe is located in the path
        C:\sysinternals\procmon.exe and the logs will be saved to the D:\procmonlogs folder.

        C:\PS> . .\TraceSystemActivity.ps1 15 C:\sysinternals D:\procmonlogs -DropFilteredEvents:$True -Filter "Process Name,contains,procmon,exclude"
#>

[CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [int]$Duration = 10,

        [Parameter(Mandatory=$true)]
        [String]$ProcMonFolder,

        [Parameter(Mandatory=$true)]
        [String]$SaveLogTo,

        [Parameter(Mandatory=$false)]
        [String]$DropFilteredEvents="false",

        [Parameter(Mandatory=$false)]
        [String]$DownloadProcmon="true",

        [Parameter(Mandatory=$false)]
        [String]$Filter = "Process Name,contains,procmon.exe,exclude;Process Name,contains,procexp.exe,exclude;Process Name,contains,Autoruns.exe,exclude;Process Name,contains,Procmon64.exe,exclude;Process Name,contains,Procexp64.exe,exclude;Process Name,contains,System,exclude"
)

function Remote-Procmon {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [int]$Duration = 10,

        [Parameter(Mandatory=$true)]
        [String]$ProcMonFolder,

        [Parameter(Mandatory=$true)]
        [String]$SaveLogTo,

        [Parameter(Mandatory=$false)]
        [boolean]$DropFilteredEvents=$false,

        [Parameter(Mandatory=$false)]
        [boolean]$DownloadProcmon=$true,

        [Parameter(Mandatory=$false)]
        [String]$Filter = "Process Name,contains,procmon.exe,exclude;Process Name,contains,procexp.exe,exclude;Process Name,contains,Autoruns.exe,exclude;Process Name,contains,Procmon64.exe,exclude;Process Name,contains,Procexp64.exe,exclude;Process Name,contains,System,exclude"
    )

    begin {
        function Convert-FilterToObj {
            [CmdletBinding()]Param(
                [ValidateNotNullOrEmpty()][String]$Filters
            )

            #Include profiling events only if paramater switch set
            if( $Filters.length -ne 0)
            {
                $Filters += ";EventClass,is,profiling,exclude"
            }

            $FilterObj = @()  #Init blank array
            #Filters must be seperated by ;
            $Filter = $filters.split(';')
            $Filter | ForEach-Object {
                $filterOptions =@()
                #Use builtin split method to convert comma seperated string into array
                foreach ($option in $_.split(',')) {
                    $filterOptions += $option
                }
                #Checking our array for exactly four objects
                if ($filterOptions.count -ne 4) {
                    Write-Error "Error: Filter `"$filterOptions`" is not in correct format."
                } else {
                    Write-Verbose "Filter `"$filterOptions`"."
                    #Convert filter into an object, remove spaces from Column/Relation
                    $CurrentFilter = New-Object system.Object
                    $CurrentFilter | Add-Member noteproperty -Name "Column" -Value $filterOptions[0].replace(' ','')
                    $CurrentFilter | Add-Member noteproperty -Name "Relation" -Value $filterOptions[1].replace(' ','')
                    $CurrentFilter | Add-Member noteproperty -Name "Value" -Value $filterOptions[2]
                    $CurrentFilter | Add-Member noteproperty -Name "Action" -Value $filterOptions[3]
                    #Add current filter to output object
                    $FilterObj += $CurrentFilter
                }
            }
            Write-Output $FilterObj
        }

	    function Convert-LargeValues {
		    [CmdletBinding()]param($Value)
		    if ($Value.Length -gt 2) {
			    $FirstByte =  [string]::Join("",$Value[1..2])
			    $SecondByte = $Value[0]
		    } else {
			    $FirstByte = $Value
			    $SecondByte = 0
		    }
		    Write-Output "0x$FirstByte","0x$SecondByte"
	    }

        function Convert-LargerValues {
		    [CmdletBinding()]param($Value)
            if ($Value.Length -gt 4 -and $Value.Length -le 6) {
			    $FirstByte =  [string]::Join("",$Value[4..5])
			    $SecondByte = [string]::Join("",$Value[2..3])
                $ThirdByte = [string]::Join("",$Value[0..1])
                Write-Output "0x$FirstByte","0x$SecondByte","0x$ThirdByte"
		    }
		    if ($Value.Length -gt 2 -and $Value.Length -le 4) {
			    $FirstByte =  [string]::Join("",$Value[2..3])
			    $SecondByte = [string]::Join("",$Value[0..1])
                Write-Output "0x$FirstByte","0x$SecondByte"
		    } else {
			    $FirstByte = $Value
			    $SecondByte = 0
                Write-Output "0x$FirstByte","0x$SecondByte"
		    }
		    
	    }

	    function Write-ProcmonFilterValue {
            [CmdletBinding()]param(
                [ValidateNotNullOrEmpty()][System.Object]$FilterObj
            )
            #Start of filter is 1 declare type as an array of bytes
            [Byte[]]$FilterRegkey = "0x1"
            #Followed by number of filters
            $NumFilters = [convert]::tostring((($FilterObj | measure).count),"16")
            #Multiple registry keys can overflow when using largeValues
            #A function was created to split these large values into two bytes
            $FilterRegKey += (Convert-LargeValues -value $NumFilters)
            #Two padding bytes
            $FilterRegkey += "0x0","0x0"
            #Header is written, build filters from friendly strings
            $FilterObj | %	{
                #Check for syntax errors
                if ($FilterRegkey -match "Error") {
                    Write-Error "Error in Write-ProcmonFilterValue: $FilterRegkey"
                } 
                #First write column code, and 9c divider
                switch($_.Column)
                {
                    "ProcessName"      {$FilterRegKey += "0x75","0x9c"}
                    "PID"              {$FilterRegKey += "0x76","0x9c"}
                    "Result"           {$FilterRegKey += "0x78","0x9c"}
                    "Detail"           {$FilterRegkey += "0x79","0x9c"}
                    "Duration"         {$FilterRegKey += "0x8d","0x9c"}
                    "ImagePath"        {$FilterRegKey += "0x84","0x9c"}
                    "RelativeTime"     {$FilterRegKey += "0x8c","0x9c"}
                    "CommandLine"      {$FilterRegKey += "0x82","0x9c"}
                    "User"             {$FilterRegKey += "0x83","0x9c"}
                    "Operation"        {$FilterRegKey += "0x77","0x9c"}
                    "ImagePath"        {$FilterRegKey += "0x84","0x9c"}
                    "Session"          {$FilterRegKey += "0x85","0x9c"}
                    "Path"             {$FilterRegKey += "0x87","0x9c"}
                    "TID"              {$FilterRegKey += "0x88","0x9c"}
                    "Duration"         {$FilterRegKey += "0x8D","0x9c"}
                    "TimeOfDay"        {$FilterRegKey += "0x8E","0x9c"}
                    "Version"          {$FilterRegKey += "0x91","0x9c"}
                    "EventClass"       {$FilterRegKey += "0x92","0x9c"}
                    "AuthenticationID" {$FilterRegKey += "0x93","0x9c"}
                    "Virtualized"      {$FilterRegKey += "0x94","0x9c"}
                    "Integrity"        {$FilterRegKey += "0x95","0x9c"}
                    "Category"         {$FilterRegKey += "0x96","0x9c"}
                    "Parent PID"       {$FilterRegKey += "0x97","0x9c"}
                    "Architecture"     {$FilterRegKey += "0x98","0x9c"}
                    "Sequence"         {$FilterRegKey += "0x7A","0x9c"}	
                    "Company"          {$FilterRegKey += "0x80","0x9c"}
                    "Description"      {$FilterRegkey += "0x81","0x9c"}
                    default            {
					                    [string]$FilterRegKey = "Error: Check Column values."
					                    Write-Error "$FilterRegKey"
					                   }
		        }
		   
                #Add two zero bytes padding before comparison
                $FilterRegkey += "0x0","0x0"
                #Now add Relation byte
                switch($_.Relation)
                {
                    "is"         {$FilterRegKey += "0x0"}
                    "isNot"      {$FilterRegkey += "0x1"}
                    "lessThan"   {$filterregkey += "0x2"}
                    "moreThan"   {$FilterRegkey += "0x3"}
                    "endsWith"   {$FilterRegkey += "0x5"}
                    "BeginsWith" {$FilterRegkey += "0x4"}
                    "Contains"   {$FilterRegKEy += "0x6"}
                    "excludes"   {$FilterRegkey += "0x7"}
                    default      {
				                   [string]$FilterRegKey = "Error: Check Relation values."
				                   Write-Error "$FilterRegKey"
                                 }
                }
		    
                #Add three zero bytes before Action (Include/Exclude)
                $FilterRegKey += "0x0","0x0","0x0"

                #Now Include/Exclude
                if ($_.Action -match "incl")     { $FilterRegkey += "0x1" }
                elseif ($_.Action -match "excl") { $FilterRegKey += "0x0" }
                else {
                    [string]$FilterRegkey = "Error: Check Action Values."
                    Write-Error "$FilterRegKey"
                }

                #Add length of <Value> string.
                #Length is hex value of (characters * 2(account for nulls) + 2)(account for spacer bytes)
                $NumPathChars = [Convert]::tostring(((($_.value.toCharArray() | measure).count *  2) + 2),"16")
                $FilterRegKey += (convert-LargeValues -value $NumPathChars)

                #Two zero bytes padding
                $FilterRegkey += "0x0","0x0"
                #Convert string "Value" to binary Ascii array (ie. A = 0x41)
                $_.Value.toCharArray() | % {
                    $FilterRegkey += (convert-largeValues -value ([Convert]::ToString(([char]$_ -as [int]),"16")))
                }
                #Current Filter calculated, pad with 10 zero bytes TTYE -- need session number at the 3rd octect or includes do not take effect for session filter
                if ($_.Column -like "PID") {
                    $hexPID = '{0:x}' -f [int]$_.Value
                    #pad the hex with leading zeros if we're an odd number
                    if ($hexPID.Length % 2 -eq 1 ) { $hexPID = $hexPID.PadLeft($hexPID.Length+1, "0") }
                    Write-Verbose "PID in hex: $HexPID"
                    $PIDEnablementValues = Convert-LargerValues -Value $hexPID
                    Write-Verbose "PID in bytes: $PIDEnablementValues"

                    #pad 2 more zero byte values...
                    $FilterRegkey += "0x0","0x0"
                    
                    #this leaves us 8 bytes to fill in with "filter enabled" values  so we need to subtract however many bytes we use from 8 and fill
                    #the rest with zeros

                    foreach ($byte in $PIDEnablementValues) {
                        $FilterRegkey += $byte
                    }
                    Write-Verbose "Filter Rules So far: $FilterRegkey"
                    $renamingZeroByte = 8 - $PIDEnablementValues.count

                    #pad with zeros
                    For ($i=0; $i -lt $renamingZeroByte; $i++) {
                        $FilterRegkey += "0x00"
                    }
                } else {
                    $FilterRegkey += "0x0","0x0","0x0","0x0","0x0","0x0","0x0","0x0","0x0","0x0"
                }
                #Check for syntax errors
                if($FilterRegkey -match "Error") {
                    Write-Error ($FilterRegkey | sort | get-unique)
                }
		
                #Set filter
                if ($env:Username -like "*$env:COMPUTERNAME*") {
                    if (-not($RunningAsSystemAccount)) {
                        Write-Verbose "It seems we are running under the SYSTEM account"
                    }
                    $RunningAsSystemAccount = $true
                    if (-not(Test-Path "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor")) {
                        New-Item "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
                    }
                    New-ItemProperty "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" "FilterRules" -Value $FilterRegKey -PropertyType Binary -Force -ErrorVariable SetRegKeyErr | Out-Null
                } else {
                    if (-not(Test-Path "HKCU:\Software\Sysinternals\Process Monitor")) {
                        New-Item "HKCU:\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
                    }
                    New-ItemProperty "HKCU:\Software\Sysinternals\Process Monitor" "FilterRules" -Value $FilterRegKey -PropertyType Binary -Force -ErrorVariable SetRegKeyErr | Out-Null
                }
                if (($setRegKeyErr | measure).count -ne 0) {
                    Write-Error "Error: Writing registry failed: $SetRegKeyErr"
                }
            }
        }
    }
    process {
        #cleanUp sysinternals keys
        if ($env:Username -like "*$env:COMPUTERNAME*") {
            if (Test-Path "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor") {
                Remove-Item "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" -Recurse -Force -ErrorVariable SetRegKeyErr | Out-Null
            } 
        } else {
            if (Test-Path "HKCU:\Software\Sysinternals\Process Monitor") {
                Remove-Item "HKCU:\Software\Sysinternals\Process Monitor" -Recurse -Force  -ErrorVariable SetRegKeyErr | Out-Null
            }
        }


        $ErrorActionPreference = "Stop"
        #needed to run under the SYSTEM account
        if (-not(Test-Path HKU:\)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
        }

        #Check if running as admin
        if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
            Write-Error "Error: Script must run with Admin rights."}

        #If ProcMonFolder ends with a "\", remove
        if (($ProcMonFolder[$ProcMonFolder.Length - 1]) -eq '\') {
            $ProcMonFolder = $ProcMonFolder.Remove((($ProcMonFolder.Length) - 1),1)
        }

        #If ProcMonFolder has a double quote '"' remove them
        if ($ProcMonFolder.Contains("`"")) {
            $ProcMonFolder = $ProcMonFolder.Replace("`"","")
        }

        #If ProcMonFolder includes the procmon.exe in it's path, we'll remove it.
        if ($ProcMonFolder.ToLower().EndsWith("procmon.exe") -or $ProcMonFolder.ToLower().EndsWith("procmon64.exe")) {
            $ProcMonFolder = ([System.IO.Path]::GetFullPath($ProcMonFolder) | Split-Path -Parent)
        }

        Write-Verbose "ProcMonFolder path: $ProcMonFolder"

        ##Main

        ##Setup running enviroment based on specified paramaters
        #Default duration to 60 seconds if not set.
        if ($Duration -eq 0) {
            $Duration = 60
        }

        #running as a ControlUp SBA under the system account?
        if ($env:Username -like "*$env:COMPUTERNAME*") {
            $RegPath = "HKU:\.DEFAULT"
        } else {
            $RegPath = "HKCU:"
        }

        #should we keep all events?
        #Include profiling events only if paramater switch set
        if($DropFilteredEvents -eq $true -and ($Filter.length) -ne 0) {
            $Filter += ";EventClass,is,profiling,exclude"
            #Set "Drop Filtered Events"
            Write-Verbose "Dropping Filtered Events.`nFilter: $($Filter | Out-String))"
            if (-not(Test-Path "$RegPath\Software\Sysinternals\Process Monitor")) {
                New-Item "$RegPath\Software\Sysinternals\Process Monitor" -Force -ErrorVariable SetRegKeyErr | Out-Null
            }
            New-ItemProperty "$RegPath\Software\Sysinternals\Process Monitor" "DestructiveFilter" -Value ("0x1","0x0","0x0","0x0") -PropertyType Binary -Force | out-null
        }

        if($DropFilteredEvents -eq $false) {
            #Keep all events
            Write-Verbose "Keeping all events"
            if (-not(Test-Path "$RegPath\Software\Sysinternals\Process Monitor")) {
                New-Item "$RegPath\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
            }
            New-ItemProperty "$RegPath\Software\Sysinternals\Process Monitor" "DestructiveFilter" -Value ("0x0") -PropertyType Dword -Force -ErrorVariable SetRegKeyErr | out-null
        }

        #Convert user input into proper object
        if(($Filter.length) -ne 0) {
	        Write-Verbose "Converting user supplied filter into object format."
	        $FilterObj = convert-FiltertoObj -Filters $Filter
        } else {
	        Write-Verbose "No filter specified, writing one that will never trigger."
	        $Filter = "ProcessName,is,AAAAAAA,exclude"
	        $FilterObj = convert-FiltertoObj -Filters $Filter
        }

        #Attempt to write filter
        Write-Verbose "Attemping to write filter registry value."
        Write-ProcmonFilterValue -FilterObj $FilterObj
				
        #If no directory specified try current folder
        if(($ProcMonFolder.Length) -eq 0) {
            $ProcMonFolder = "."
        }

        #ensure the target procmonfolder directory exists, and create it if it does not exist
        if (-not(Test-Path $ProcMonFolder -PathType Container)) {
            Write-Verbose -Message "Path: `"$ProcMonFolder`" was not found. Attempting to create directoy structure."
            New-Item -Path "$ProcMonFolder" -ItemType Directory -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Set-Location $ProcMonFolder
            if ((Get-Location).Path -notlike $ProcMonFolder) {
                Throw "$ProcMonFolder path does not exist and could not be created."
            }
        } else {
            Write-Verbose -Message "ProcMonFolder variable is set to a directory: $ProcMonFolder"
            Set-Location $ProcMonFolder
        }

        #Test for executable
        $FoundProcmon = $true
        Write-Verbose "Testing for Procmon.exe."
        if (-not(Test-Path "$ProcMonfolder\procmon.exe")) {
            Write-Output "Unable to find procmon.exe in this path: $ProcMonfolder\procmon.exe`". Procmon will be downloaded automatically."
            $FoundProcmon = $false
        }

        #if FoundProcmon is false and DownloadProcmon is true, then download procmon.
        if (($FoundProcmon -eq $false) -and ($DownloadProcmon -eq $true)) {
            #borrowed from the great @guyrleech
            Write-Output "Downloading procmon ..."
            [Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            (New-Object System.Net.WebClient).DownloadFile( 'https://live.sysinternals.com/procmon.exe' , "$ProcMonfolder\procmon.exe" )
            if( ! ( Test-Path "$ProcMonfolder\procmon.exe" -ErrorAction SilentlyContinue -PathType Leaf ) ) {
                Throw "Failed to download procmon. Please place procmon in the target directory and try again."
            }
            Unblock-File -Path "$ProcMonfolder\procmon.exe" -Confirm:$false
            $downloadedProcmon = $true
            $signing = Get-AuthenticodeSignature -FilePath "$ProcMonfolder\procmon.exe" -ErrorAction SilentlyContinue
            if( ! $signing ) {
                Throw "Could not get signing information from `"$ProcMonfolder\procmon.exe`""
            }
            if( ! $signing.Status -ne 'Valid' ) {
                Throw "Certificate status for `"$ProcMonfolder\procmon.exe`" is $($signing.Status), not `"Valid`""
            }
            if( $signing.SignerCertificate.Subject -notmatch '^CN=Microsoft Corporation,' ) {
                Throw "`"$ProcMonfolder\procmon.exe`" is not signed by Microsoft Corporation, found $($signing.SignerCertificate.Subject)"
            } elseif( ! ( Test-Path $ProcMonfolder\procmon.exe -ErrorAction SilentlyContinue -PathType Leaf ) ) {
                Throw "Procmon not found at $procmon"
            }
        }

        #Try and unblock procmon executable.
        Unblock-File -Path "$ProcMonfolder\procmon.exe" -Confirm:$false
       
        #test the SaveLogTo directory
        if (-not(Test-Path $SaveLogTo)) {
            Write-Host "Path $SaveLogTo not found.  Attempting to create it"
            try {
                New-Item $SaveLogTo -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            catch 
            {
                Write-Error "Unable to create directory `"$SaveLogTo`""
            }
        }

        #If SaveLogTo ends with a "\", remove
        if (($SaveLogTo[$SaveLogTo.Length - 1]) -eq '\') {
            $SaveLogTo = $SaveLogTo.Remove((($SaveLogTo.Length) - 1),1)
        }
        Write-Verbose "Log output directory: $SaveLogTo"


        #Cheap way to get a pseudorandom tempfile name with benifit of timestamp
        $FileDate=((get-date).TimeOfDay.ToString().Replace('.','_')).replace(':','_')
        $TempPML = "$SaveLogTo\Temp_$FileDate.pml"

        

        #Run procmon backed to file, suppress prompts
        if (([bool]([System.Diagnostics.Process]::GetProcessesByName("procmon"))) -or ([bool]([System.Diagnostics.Process]::GetProcessesByName("procmon64")))) {
            Write-Verbose "Procmon is already running. Attempting to terminate existing processes"
            start-process -filepath ".\Procmon.exe" -argument "/terminate /quiet /accepteula" -Wait -NoNewWindow
            #if procmon is still running it will be forcibly terminated.
            if ([bool]([System.Diagnostics.Process]::GetProcessesByName("procmon"))) {
                Write-Verbose "Attemping to forcibly terminate an existing procmon process"
                [System.Diagnostics.Process]::GetProcessesByName("procmon").Kill()
            }
            if ([bool]([System.Diagnostics.Process]::GetProcessesByName("procmon64"))) {
                Write-Verbose "Attemping to forcibly terminate an existing procmon64 process"
                [System.Diagnostics.Process]::GetProcessesByName("procmon64").Kill()
            }
        }
        
        Write-Verbose "Attempting to start Process Monitor."
        start-process -filepath ".\Procmon.exe" -argument "/runtime $Duration /backingfile $TempPML /quiet /accepteula" -WindowStyle Hidden

        Write-Host "Process Monitor running for $($Duration + 5) seconds."
        Write-Host "Saved the file to: `"\\$env:computername`\$($($TempPML).TrimStart(".").Replace(':','$'))`""
        Write-Host "Expected completion time: $((Get-date).AddSeconds($duration).ToLongTimeString())."

    }
}

#convert the String parameter of "DropFilteredEvents" into a boolean
if ($DropFilteredEvents -like "false" -or $DropFilteredEvents -eq "0") {
    Remove-Variable DropFilteredEvents
    $DropFilteredEvents = $false
} else {
    Remove-Variable DropFilteredEvents
    $DropFilteredEvents = $true
}

if ($DownloadProcmon -like "false" -or $DownloadProcmon -eq "0") {
    Remove-Variable DownloadProcmon
    $DownloadProcmon = $false
} else {
    Remove-Variable DownloadProcmon
    $DownloadProcmon = $true
}

[hashtable]$params = @{
    'Duration'           = $duration
    'ProcMonFolder'      = $ProcMonFolder
    'SaveLogTo'          = $SaveLogTo
    'DropFilteredEvents' = $DropFilteredEvents
    'Filter'             = $Filter
    'DownloadProcmon'    = $DownloadProcmon
}

Remote-Procmon @params
