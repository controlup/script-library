<#

.DESCRIPTION
    
     This script displays a users Name, Company, Title, Office, Office Phone Number, Mobile Phone Number
     and E-Mail Address, as well as those same details for their Manager. Created by Rory Monaghan. 

.PARAMETER User
    
     Specifies an Active Directory user object provided through a ControlUp Argument

#>

$user=$args[0].split("\")[1]
write-host $user

$ErrorActionPreference = "Stop"

If ( (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null )
{
    Try {
        Import-Module ActiveDirectory
    } Catch {
        Write-Error "Unable to load the module" -ErrorAction Continue
        Write-Error $Error[1] -ErrorAction Continue
        Exit 1
    }
}

Try {
    $managerfirstname=(get-aduser (get-aduser $user -Properties manager).manager).GivenName
    $managerlastname=(get-aduser (get-aduser $user -Properties manager).manager).SurName
    $manageraccname=(get-aduser (get-aduser $user -Properties manager).manager).samaccountName

    Get-AdUser -Identity $user -Properties * | Select GivenName,Surname,Department,Title,Company,mail,telephoneNumber,MobilePhone,Office
    write-host "Reporing Manager is $managerfirstname $managerlastname : "
    Get-AdUser -Identity $manageraccname -Properties * | Select GivenName,Surname,Department,Title,Company,mail,telephoneNumber,MobilePhone,Office

} Catch {
    Write-Error $Error[0] -ErrorAction Continue
    Exit 1
}


