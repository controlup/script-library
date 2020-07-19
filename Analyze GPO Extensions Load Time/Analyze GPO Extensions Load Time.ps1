function Get-GPUserCSE
{
<#  
.SYNOPSIS
        Lists every Group Policy Client Side Extension
        and their associated load time in milliseconds.
.DESCRIPTION
        This script looks under the 'Group Policy Event Log'
        and lists every applied Group Policy Client Side Extensions.  

.PARAMETER Username
        Type in the Positional argument of the Down-Level Logon Name (Domain\User)

.EXAMPLE
        Get-GPUserCSE -Username MyDomain\MyUser

        CSE Name                  Time(ms) GPOs                                  
        --------                  -------- ----                                  
        Group Policy Registry          531 VSI User-V4, XenApp 6.5 User Env      
        Registry                       296 Local Group Policy, Local Group Policy
        Citrix Group Policy            281 Local Group Policy, Local Group Policy
        Scripts                         93 VSI User-V4, VSI System-V4            
        Folder Redirection              78 None                                  
        Citrix Profile Management       16 None                                  
        
        
        Group Policy Client Side Extenstions with an error
        
        CSE Name                   Time(ms) ErrorCode GPOs                      
        --------                   -------- --------- ----                      
        Internet Explorer Branding       16       127 VSI User-V4, VSI System-V4

.LINK
        See http://www.controlup.com
#>

[CmdletBinding()] 
 Param( 
    [Parameter(Mandatory=$false, 
    ValueFromPipelineByPropertyName=$true)] 
    [String]
    $Username
    ) 

begin{ 
    $ErrorActionPreference = "Stop"

    # XPath query used to get evend id 4001.
    $Query = "*[EventData[Data[@Name='PrincipalSamName'] and (Data='$Username')]] and *[System[(EventID='4001')]]"
}

Process{ 
    try {
        [array]$Events = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath "$Query"
        $ActivityId = $Events[0].ActivityId.Guid
    }

    catch {
        Write-Host "Could not find relevant events in the Microsoft-Windows-GroupPolicy/Operational log. `n`
        The default log size (4MB) only supports user sessions that logged on a few hours ago. `
        Please increase the log size to support older sessions."
        Exit 1
    }

    try {

        # Gets all events that match event id 4016,5016,6016 and 7016 and correlated with the activity id of event id 4001.
        [array]$CSEarray = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath @"
        *[System[(EventID='4016' or EventID='5016' or EventID='6016' or EventID='7016') and Correlation[@ActivityID='{$ActivityID}']]]
"@
    }

    catch {
        Write-Host "Could not find relevant events in the Microsoft-Windows-GroupPolicy/Operational log. `n`
        It's seems like there are no Client Side Extensions applied to your session."
        Exit 1
    }

    try {

        [array]$GPEnd = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath "*[EventData[Data[@Name='PrincipalSamName'] and (Data='$Username')]] and *[System[(EventID='8001')]]"
    }

    catch {
        Write-Host "Could not find relevant events in the Microsoft-Windows-GroupPolicy/Operational log. `n`
        It's seems like there are no Client Side Extensions applied to your session."
        Exit 1
    }


    $Output = @()

    # Run only for for event id 4016 records.
    foreach ($i in ($CSEarray | Where-Object {$_.Id -eq '4016'})) {
        $obj = New-Object -TypeName psobject
        $obj | Add-Member -MemberType NoteProperty -Name Name -Value ($i.Properties[1] | Select-Object -ExpandProperty Value)
        $obj | Add-Member -MemberType NoteProperty -Name String -Value (($i.Properties[5] `
        | Select-Object -ExpandProperty Value).trimend("`n") -replace "`n",", ")

        # Every object in output has CSE Name and String of all the GPO Names.
        $Output += $obj
    }
    # Run only for for event id 5016,6016 and 7016 records.
    foreach ($i in ($CSEarray | Where-Object {$_.Id -ne '4016'})) {

        # Add the duration of the CSE to the object.
        $Output | Where-Object {$_.Name -eq ($i.Properties[2] | Select-Object -ExpandProperty Value)} `
        | Add-Member -MemberType NoteProperty -Name Time -Value ($i.Properties[0] | Select-Object -ExpandProperty Value)

        # Add the ErrorCode to the object
        $Output | Where-Object {$_.Name -eq ($i.Properties[2] | Select-Object -ExpandProperty Value)} `
        | Add-Member -MemberType NoteProperty -Name ErrorCode -Value ($i.Properties[1] | Select-Object -ExpandProperty Value)

    }

}

End{
    $TableFormat = @{Expression={$_.Name};Label="CSE Name"}, `
    @{Expression={$_.Time};Label="Time(ms)"}, `
    @{Expression={$_.String};Label="GPOs"}
    $TableFormatWithError = @{Expression={$_.Name};Label="CSE Name"}, `
    @{Expression={$_.Time};Label="Time(ms)"}, `
    @{Expression={$_.ErrorCode};Label="ErrorCode"}, `
    @{Expression={$_.String};Label="GPOs"}

    $GPTotalDuration = $($GPEnd[0].TimeCreated - $Events[0].TimeCreated).TotalSeconds

    if ($GPTotalDuration -gt 0) {
        "Overall Group Policy Processing Duration:`t" + "{0:N2}" -f $GPTotalDuration + " Seconds"
        } 
    
    $Output | Where-Object {$_.ErrorCode -eq 0} | Sort-Object Time -Descending | Format-Table $TableFormat -AutoSize -Wrap

        if (($Output.ErrorCode | Measure-Object -Sum).Sum -ne 0) {
            Write-Host "Group Policy Client Side Extenstions with an error"
            $Output | Where-Object {$_.ErrorCode -ne 0} | Sort-Object Time -Descending | Format-Table $TableFormatWithError -AutoSize -Wrap
        }
        $TotalSeconds = (($Output | %{$_.Time} | Measure-Object -Sum | Select-Object -ExpandProperty Sum)/1000)
        "GP Extensions Processing Duration:`t" + "{0:N2}" -f $TotalSeconds + " Seconds"
        if ($GPTotalDuration -gt 0 -and $GPTotalDuration -gt $TotalSeconds) {
        "GP Processing Duration (not attributed to specific extensions):`t" + "{0:N2}" -f $($GPTotalDuration-$TotalSeconds) + " Seconds"
        }
    }
}

Get-GPUserCSE -Username $args[0]
