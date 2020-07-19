<#
ControlUp-friendly version of Carl Webster's 'Get-FRErrorsV2' script (Get Folder Redirection Errors)
http://carlwebster.com/downloads/download-info/get-gpo-folder-redirection-errors-xenapp-6-5/
#>

[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-30))
[Datetime]$EndDate = (Get-Date -displayhint date)
try {
    $Errors = Get-EventLog -logname application `
        -source "Microsoft-Windows-Folder Redirection" `
        -entrytype "Error" `
        -after $StartDate `
        -before $EndDate `
        -EA 0
    } catch {
    	Write-Host "Error querying event log"
		Continue
	}
If($? -and $Null -ne $Errors) {
    If (!($Errors -is [Array])) {
        #force the singleton to an array
        [array]$Errors = $Errors
    }

    $Users = $Errors | Select UserName | Sort UserName -Unique	
    If (!($Users -is [Array])) {
        #force the singleton to an array
        [array]$Users = $Users
    }
    $ErrorPaths = @()
    Write-Host "$($Errors.Count) Folder Redirection errors found for $($Users.Count) users"

    Foreach ($Err in $Errors) {
        $m = $null
        $ErrorPath = $null
        $m = $Err.Message
        try {
            $ErrorPath = $m.substring($m.indexof("\\"),$m.indexof('".')-$m.indexof("\\"))
        } catch {
            # Meh
        }
        if ($ErrorPath -ne $null) {
            $ErrorPaths += $ErrorPath
        }

    }
    $ErrorPaths = $ErrorPaths | Sort -Unique
    Write-Host "$($ErrorPaths.Count) erroneous paths found:"
    $ErrorPaths

}
Else
{
    Write-Host "No Folder Redirection errors found"
}
