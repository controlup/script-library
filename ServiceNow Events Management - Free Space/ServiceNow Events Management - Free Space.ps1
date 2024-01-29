#requires -Version 5
<#
    .SYNOPSIS
        Sends a payload to ServiceNow in the Events Management database

    .DESCRIPTION
        Sends a payload to ServiceNow in the Events Management database
		
	.PARAMETER	<LogicalDisk <String>>
		The Logical Disk of the resource affected

	.PARAMETER	<Machine <String>>
		The Name of the machine

	.PARAMETER	<Severity <String>>
		Severity Level of the event

    .PARAMETER	<ServiceNowEndPoint <String>>
		URL of the ServiceNow Events Management API

    .PARAMETER	<AuthenticationHeader <String>>
		Base64 Encoded authentication header string

    .EXAMPLE
        . .\ServiceNowLogicalDisk.ps1 -LogicalDisk "C:\" -Machine "NPCVAD2022" -Severity "4" -ServiceNowEndPoint "https://dev142496.service-now.com/api/global/em/jsonv2" -AuthenticationHeader "YWO9dn4k6K8vTynfkSTTTy9mcjdERlg4"

    .NOTES

    .CONTEXT
        LogicalDisk

    .MODIFICATION_HISTORY
        Created TTYE : 2023-12-21


    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Logical Disk Drive Letter')][ValidateNotNullOrEmpty()]                          [string]$LogicalDisk,
    [Parameter(Mandatory=$true,HelpMessage='Name of the machine')][ValidateNotNullOrEmpty()]                                [string]$Machine,
    [Parameter(Mandatory=$true,HelpMessage='Severity Value')][ValidateNotNullOrEmpty()]                                     [string]$Severity,
    [Parameter(Mandatory=$true,HelpMessage='ServiceNow Events Management API URL')][ValidateNotNullOrEmpty()]               [string]$ServiceNowEndpoint,
    [Parameter(Mandatory=$true,HelpMessage='Base64 Authentication Header')][ValidateNotNullOrEmpty()]                       [string]$AuthenticationHeader

)

function Get-Size {
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='String representing size')][ValidateNotNullOrEmpty()]          [string]$Size
)

    #$Size should come in as a string like "3.9 (GB)"
    #expectataion is the value in the parenthesis is the size abbreviation and the float value is the multiplier

    $Size -match "[0-9.]*" | Out-Null
    [float]$Number = $Matches[0]

    $ReturnObj = [System.Collections.Generic.List[object]]::new()
    $ReturnObj.Add(@{ OriginalValue     = $number })
        
    switch -regex ($Size)
    {
        'TB' { $ReturnObj.Add(@{ OriginalSize     = "TB" }) ; $ReturnObj.Add(@{ SizeInBytes     = ($Number*1099511627776) }) }
        'GB' { $ReturnObj.Add(@{ OriginalSize     = "GB" }) ; $ReturnObj.Add(@{ SizeInBytes     = ($Number*1073741824) }) }
        'MB' { $ReturnObj.Add(@{ OriginalSize     = "MB" }) ; $ReturnObj.Add(@{ SizeInBytes     = ($Number*1048576) }) }
        'KB' { $ReturnObj.Add(@{ OriginalSize     = "KB" }) ; $ReturnObj.Add(@{ SizeInBytes     = ($Number*1024) }) }
    }
    Return $ReturnObj
}

#Start-Transcript -Path "D:\Log.txt" -Force
Write-Output "LogialDisk:$LogicalDisk"
Write-Output "Machine:$Machine"
Write-Output "ServiceNowEndpoint:$ServiceNowEndpoint"
Write-Output "AuthenticationHeader:$($AuthenticationHeader.Substring(0,3))..."

#ServiceNow Events Management API Endpoint
$emEndpoint = $ServiceNowEndpoint


$headers = @{
    "Content-Type"="application/json"
    "Accept"="application/json"
    }

<#  ### Use the following code snippet to get your Base64 Authentication string
# Set the credentials
$User = 'admin'
$Pass = 'Password123!'
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $User, $Pass)))
#>

# Build & set authentication header
$headers.Add('Authorization', ('Basic {0}' -f $AuthenticationHeader))

$Source = "ControlUp"
$resource = ($LogicalDisk | ConvertTo-Json).Replace("`"","")
$node = $Machine
$type = "Disk"
$MetricName = "DiskSpace"
$messageKey = "$($source)$($node)$($type)$($resource)$($MetricName)"

#Change the message description and severity level based on the current state. Eg, if the free space is less than threshold then we set the severity via the passed in level.
#If the freespace is greater than the threshold then we make severity "0"

$CurrentSize = Get-Size -Size $($CUTriggerObject.Columns.ColumnValueAfter)
$ThresholdSize = Get-Size -Size $($CUTriggerObject.Columns.ColumnCrossedThreshold)

if ($CurrentSize.SizeInBytes -le $ThresholdSize.SizeInBytes) {
    $DescriptionMessage = "The disk $resource on computer $Machine is running out of disk space. Free space is at $($CUTriggerObject.Columns.ColumnValueAfter), below the threshold of $($CUTriggerObject.Columns.ColumnCrossedThreshold)."
} else {
    $DescriptionMessage = "The disk $resource on computer $Machine now has more free space ($($CUTriggerObject.Columns.ColumnValueAfter)) then the configured threshold of $($CUTriggerObject.Columns.ColumnCrossedThreshold)."
    $Severity=0
}



$payLoad = @"
{ "records":	
 [
  {
   "source":"$source",
   "event_class":"",
   "resource":"$resource",
   "node":"$node",
   "metric_name":"$MetricName",
   "message_key":"$messageKey",
   "type":"$type",
   "severity":$Severity,
   "description":"$DescriptionMessage",
   "additional_info": '{"custom_field1":"value1","custom_field2":"value2"}'
  }     
 ]
}
"@


## If we get "The remote server returned an error: (500) Internal Server Error." Then it's been observed the payload has text issues. Specifically, ensuring that special characters like
## backslash is properly escaped

try {
    $webResult = Invoke-WebRequest -Uri $emEndpoint -Method Post -Headers $headers -Body $payLoad -UseBasicParsing
} catch {
    Write-Error "$($failure.Exception | select-object *)"
}
#Stop-Transcript

