<#
 	.SYNOPSIS
        This script that restarts the BrokerAgent service on a Citrix VDA.
		
    .DESCRIPTION
        This is a simple script that restarts the BrokerAgent service on a Citrix VDA. This is useful when VDAs become unregistered. 
        It forces them to try to re-register to a Delivery Controller. In some instances a reboot of the VDA may be required, 
        for everything else there's this SBA!
		   
    .NOTES
        The script checks to see if the service exists. If it does not, it will fail gracefully. 
        It will also check to see if the service is running or not. If it is has been found, but it is not running it will return a
        message to inform you that the service is not currently running. If the service exists and is currently running, the script
        will restart the service.
		
    .LINK
        For more information refer to:
            http://www.controlup.com

    .LINK
        Stay in touch:
        http://twitter.com/rorymon

#>

$serviceName = "BrokerAgent"

If (Get-Service $serviceName -ErrorAction SilentlyContinue) {

    If ((Get-Service $serviceName).Status -eq 'Running') {

        Restart-Service $serviceName -Force

    } Else {

        write-host "$serviceName has been found, but it is not running."

    }

} Else {

    write-host "$serviceName not found"

}
