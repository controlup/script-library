<#
 .NAME:     Addition_to_Local_Admins.ps1

 .CREDIT:   https://security.stackexchange.com/questions/149519/how-to-find-who-granted-local-admin-privileges-to-a-user  
            https://girl-germs.com/?p=363     which GPO corresponds with which Event ID
                Need to verify the computer has 'Audit Security Group Management' in Accoutn MAnagement enabled
#>

$ErrorActionPreference = "ignore"

# Check if the Audit policy for recording the event 4732 is enabled for Success
$checkPol = (auditpol /get /subcategory:"Security Group Management" | findstr "Success")

if([string]::IsNullOrEmpty($checkPol))
    { Write-Output 'Audit policy not properly configued
       run:
       auditpol /set /subcategory:"Security Group Management" /success:enable';
     exit
    }

### Create a filter query to search for additions to BUILTIN\Administrators
### Security log event ID 4732
### Adding specifically to the Administrators SID
$xmlFilter = @"
<QueryList>
<Query Id="0" Path="Security">
<Select Path="Security">
*[System[(EventID=4732)]] 
and 
*[EventData[Data[@Name='TargetSid'] and Data='S-1-5-32-544']]
</Select>
</Query>
</QueryList>
"@

# Query and get the events
try {$adm_inclusion = Get-WinEvent -FilterXml $xmlFilter}
    Catch {Write-Output "No events found (and auditpol was properly configured)"; exit }

$adm_inclusion | Format-List -Property TimeCreated,Id,Message | findstr /C:"TimeCreated" /C:"Subject:" /C:"Security ID" /C:"Account" /C:"Member" /C:"Group"

#$adm_inclusion.Message | findstr /C:"Subject:" /C:"Security ID" /C:"Account" /C:"Member" /C:"Group"