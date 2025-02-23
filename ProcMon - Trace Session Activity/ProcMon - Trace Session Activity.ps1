<#
 	.SYNOPSIS
        Capture session information using process monitor.
		
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
		
	.PARAMATER <SessionId <string[]>
		The session ID to set in the filter for the process monitor capture.
		
	.PARAMETER	<Duration <string[]>
		The duration (in seconds) to run the capture.

	.PARAMETER  <ProcMonFolder <string[]>
		Specifies folder that contains procmon.exe.

    .PARAMETER  <SaveLogTo <string[]>
		Specifies folder the process monitor log file will be saved.
		
    .LINK
        For more information refer to:
            http://theorypc.ca

    .LINK
        Stay in touch:
        http://twitter.com/trententtye

    .EXAMPLE
        C:\PS> Remote-Procmon -SessionId 3 -Duration 15 -ProcMonFolder C:\sysinternals -SaveLogTo D:\procmonlogs
		
		Starts process monitor for 15 seconds, filtering for session 3, keeping all events except exclusions.  Procmon.exe is located in the path
        C:\sysinternals\procmon.exe and the logs will be saved to the D:\procmonlogs folder.

        C:\PS> . .\TraceSessionActivity 3 15 C:\sysinternals D:\procmonlogs
#>

function Remote-Procmon {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [String]$SessionId,

        [Parameter(Mandatory=$false)]
        [int]$Duration = 10,

        [Parameter(Mandatory=$true)]
        [String]$ProcMonFolder,

        [Parameter(Mandatory=$true)]
        [String]$SaveLogTo
    )

    begin {
        function Convert-FilterToObj {
            [CmdletBinding()]Param(
                [ValidateNotNullOrEmpty()][String]$Filters
            )

            #Include profiling events only if paramater switch set
            if((-not($KeepAll)) -and (($Filters.length) -ne 0))
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

	    function Write-ProcmonFilterValue {
            [CmdletBinding()]param(
                [ValidateNotNullOrEmpty()][System.Object]$FilterObj
            )
            #Start of filter is 1 declare type as an array of bytes
            [Byte[]]$FilterRegkey = "0x1"
            #Followed by number of filters
            $NumFilters = [convert]::tostring((($FilterObj | Measure-Object).count),"16")
            #Multiple registry keys can overflow when using largeValues
            #A function was created to split these large values into two bytes
            $FilterRegKey += (Convert-LargeValues -value $NumFilters)
            #Two padding bytes
            $FilterRegkey += "0x0","0x0"
            #Header is written, build filters from friendly strings
            $FilterObj | ForEach-Object	{
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
                $NumPathChars = [Convert]::tostring(((($_.value.toCharArray() | Measure-Object).count *  2) + 2),"16")
                $FilterRegKey += (convert-LargeValues -value $NumPathChars)

                #Two zero bytes padding
                $FilterRegkey += "0x0","0x0"
                #Convert string "Value" to binary Ascii array (ie. A = 0x41)
                $_.Value.toCharArray() | ForEach-Object {
                    $FilterRegkey += (convert-largeValues -value ([Convert]::ToString(([char]$_ -as [int]),"16")))
                }
                #Current Filter calculated, pad with 10 zero bytes TTYE -- need session number at the 3rd octect or includes do not take effect for session filter
                if ($_.Column -like "Session") {
                    $hexSessionID = '{0:x}' -f $_.Value
                    $FilterRegkey += "0x0","0x0","0x$hexSessionID","0x0","0x0","0x0","0x0","0x0","0x0","0x0"
                } else {
                    $FilterRegkey += "0x0","0x0","0x0","0x0","0x0","0x0","0x0","0x0","0x0","0x0"
                    }
                }
                #Check for syntax errors
                if($FilterRegkey -match "Error") {
                    Write-Error ($FilterRegkey | Sort-Object | get-unique)
                }
		
                #Set filter
                if ($env:Username -like "*$env:COMPUTERNAME*") {
                    Write-Verbose "It seems we are running under the SYSTEM account"
                    if (-not(Test-Path "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor")) {
                        New-Item "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
                    }
                    New-ItemProperty "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" "FilterRules" -Value $FilterRegKey -PropertyType Binary -Force -ErrorVariable SetRegKeyErr | Out-Null
                } else {
                    if (-not(Test-Path "HKCU:\Software\Sysinternals\Process Monitor")) {
                        New-Item "HKCU:\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
                    }
                    New-ItemProperty "HKCU:\Software\Sysinternals\Process Monitor" "FilterRules" -Value $FilterRegKey -PropertyType Binary -Force -ErrorVariable SetRegKeyErr | Out-Null
                if (($setRegKeyErr | Measure-Object).count -ne 0) {
                    Write-Error "Error: Writing registry failed: $SetRegKeyErr"
                }
            }
        }
    }
    process {
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
            $ProcMonFolder.Remove((($ProcMonFolder.Length) - 1),1)
        }


	
        ##Main
        $Filters = "Session,is,$SessionId,include"

        ##Setup running enviroment based on specified paramaters
        #Default duration to 60 seconds if not set.
        if ($Duration -eq 0) {
            $Duration = 60
        }




        if (-not($KeepAll)) {
            #Set "Drop Filtered Events"
            if ($env:Username -like "*$env:COMPUTERNAME*") {
                if (-not(Test-Path "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor")) {
                    New-Item "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" -Force -ErrorVariable SetRegKeyErr | Out-Null
                }
                    New-ItemProperty "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" "DestructiveFilter" -Value ("0x1") -PropertyType Dword -Force -ErrorVariable SetRegKeyErr | out-null
                } else {
                    if (-not(Test-Path "HKCU:\Software\Sysinternals\Process Monitor")) {
                        New-Item "HKCU:\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
                    }
                    New-ItemProperty "HKCU:\Software\Sysinternals\Process Monitor" "DestructiveFilter" -Value ("0x1") -PropertyType Dword -Force -ErrorVariable SetRegKeyErr | out-null
                if (($setRegKeyErr | Measure-Object).count -ne 0) {
                    Write-Error "Error: Writing registry failed: $SetRegKeyErr"
                }
            }
        } else {
            #keep all events
            if ($env:Username -like "*$env:COMPUTERNAME*") {
                if (-not(Test-Path "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor")) {
                    New-Item "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" -Force -ErrorVariable SetRegKeyErr | Out-Null
                }
                    New-ItemProperty "HKU:\.DEFAULT\Software\Sysinternals\Process Monitor" "DestructiveFilter" -Value ("0x0") -PropertyType Dword -Force -ErrorVariable SetRegKeyErr | out-null
                } else {
                    if (-not(Test-Path "HKCU:\Software\Sysinternals\Process Monitor")) {
                        New-Item "HKCU:\Software\Sysinternals\Process Monitor" -Force  -ErrorVariable SetRegKeyErr | Out-Null
                    }
                    New-ItemProperty "HKCU:\Software\Sysinternals\Process Monitor" "DestructiveFilter" -Value ("0x0") -PropertyType Dword -Force -ErrorVariable SetRegKeyErr | out-null
                if (($setRegKeyErr | Measure-Object).count -ne 0) {
                    Write-Error "Error: Writing registry failed: $SetRegKeyErr"
                }
            }
        }
        #Convert user input into proper object
        if(($Filters.length) -ne 0) {
	        Write-Verbose "Converting user supplied filter into object format."
	        $FilterObj = convert-FiltertoObj -Filters $Filters
        } else {
	        Write-Verbose "No filter specified, writing one that will never trigger."
	        $Filters = "pid,is,AAAAAAA,exclude"
	        $FilterObj = convert-FiltertoObj -Filters $Filters
        }

        #Attempt to write filter
        Write-Verbose "Attemping to write filter registry value."
        Write-ProcmonFilterValue -FilterObj $FilterObj
				
        #If no directory specified try current folder
        if(($ProcMonFolder.Length) -eq 0) {
            $ProcMonFolder = "."
        }
        Set-Location $ProcMonFolder

       
        #test the SaveLogTo directory
        if (-not(Test-Path $SaveLogTo)) {
            Write-Output "Path $SaveLogTo not found.  Attempting to create it"
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
        $TempCSV = ".\Temp_$FileDate.csv"

        #Test for executable
        Write-Verbose "Testing for Procmon.exe."
        if (-not(Test-Path "$ProcMonfolder\procmon.exe")) {
            Write-Error "Tested: `"$ProcMonfolder\procmon.exe`""
        }

        #Run procmon backed to file, supress prompts
        Write-Verbose "Attempting to start Process Monitor."
        start-process -filepath ".\Procmon.exe" -argument "/runtime $Duration /backingfile $TempPML /quiet /accepteula" -WindowStyle Hidden

        #Sleep for a number of seconds procmon should run
        Write-Output "Process Monitor running for $($Duration + 5) seconds."
        Write-Output "Saved the file to: `"\\$env:computername`\$($($TempPML).TrimStart(".").Replace(':','$'))`""
        Write-Output "Expected completion time: $((Get-date).AddSeconds($duration).ToLongTimeString())."
    }
}

Remote-Procmon -SessionId $args[0] -Duration $args[1] -ProcMonFolder $args[2] -SaveLogTo $args[3] -Verbose

