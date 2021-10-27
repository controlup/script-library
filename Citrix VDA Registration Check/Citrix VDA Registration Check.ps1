#requires -Version 3.0
<#  
.SYNOPSIS     Check Citrix VDA Registration Status [ALPHA]
.DESCRIPTION  [ALPHA] Citrix VDAs need to register with a DDC.
              Troubleshoot and give information on possible causes of unregistered VDAs.
              

.SOURCES      

.EXAMPLE:     \\util01\share\research\Citrix_VDA_Registration.ps1
.CONTEXT      Machine
.TAGS         $HDX, $Citrix, $VDA
.HISTORY      Chris Rogers     - 2021-04-19 - ALPHA Release 
#>

#Set-StrictMode -Version Latest
[string]$ErrorActionPreference = 'Stop'
# $VerbosePreference = 'Continue'      # Remove the comment in the begining to enable Verbose output
# $DebugPreference = 'Continue'      # Remove the comment in the begining to enable Debug output


###############################################################################################################
# VDA Checks
# expand checks to include both conditions where applicable

# Verify domain membership - $domainTest

    $name=$env:COMPUTERNAME
    $domain=$env:USERDNSDOMAIN

    # PartOfDomain (boolean Property)
    $domainMember=(Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain

    $domainTest = $domainMember

    Write-Debug "Name: $name"
    Write-Debug "Domain: $domain"
    
    if($domainMember){
        Write-Debug "Verified: $domain"
    } 
    else{
        Write-Debug "Domain NOT Verified"
    }
    

# look for the VDA ListOfDDCs in the registry - $ddcTest

    $DDCs=((Get-ItemProperty -Path HKLM:\SOFTWARE\Citrix\VirtualDesktopAgent).ListOfDDCs -split" ")
    $ddc1=$DDCs[0]
    $ddc2=$DDCs[1]

    $ddcTest = $DDCs
            
    if($DDCs){
        Write-Debug "Verified DDC1: $ddc1"
        Write-Debug "Verified DDC2: $ddc2"
    } 
    else{
        Write-Debug "DDCs NOT Verified"
    }

# Resolve DNS for DDCs - $dnsTest
    
    $ddcip1=(Resolve-DnsName -Name $ddc1 -Type A -ErrorAction Ignore | select -exp IPAddress)
    $ddcip2=(Resolve-DnsName -Name $ddc2 -Type A -ErrorAction Ignore | select -exp IPAddress)

    $DDCIPs=@($ddcip1, $ddcip2)

    $dnsTest = $DDCIPs

    if($DDCIPs){
        Write-Debug "Verified DDC1 IP: $ddcip1"
        Write-Debug "Verified DDC2 IP: $ddcip2"
    } 
    else{
        Write-Debug "DDCs DNS NOT Verified"
    }

# Reverse DNS for DDCs - $revDnsTest

    $ddcrev1=(Resolve-DnsName $ddcip1 -Type PTR -ErrorAction Ignore | select -exp NameHost)
    $ddcrev2=(Resolve-DnsName $ddcip2 -Type PTR -ErrorAction Ignore | select -exp NameHost)

    
    $DDCRevIPs=@($ddcrev1, $ddcrev2)

    $revDnsTest = $DDCRevIPs

    if($DDCRevIPs){
        Write-Debug "Verified DDC1 Reverse DNS: $ddcrev1"
        Write-Debug "Verified DDC2 Reverse DNS: $ddcrev2"
    } 
    else{
        Write-Debug "DDCs Reverse DNS NOT Verified"
    }

# verify connectivity to DDCS - $pingTest

    $ddcping1=(test-connection -computername "$ddc1" -quiet -count 1)
    $ddcping2=(test-connection -computername "$ddc2" -quiet -count 1)

    $DDCpings=@($ddcping1, $ddcping2)

    $pingTest = $DDCpings

    if($DDCpings){
        Write-Debug "Verified DDC1 Ping: $ddcping1"
        Write-Debug "Verified DDC2 Ping: $ddcping2"
    } 
    else{
        Write-Debug "DDCs Ping NOT Verified"
    }

    #port 80/443 check?

# verify VDA time sync - $timeTest

if ( $env:logonserver -match "\\\\(?<server>(.+))")
{
    $timeServer = $matches['server']
    Write-Debug "NTP Server: $timeServer"    
}
else
{
    $timeServer = 'pool.ntp.org'
    Write-Debug "NTP Server: $timeServer"    
}


### Credit:  https://chrisjwarwick.wordpress.com/2012/08/26/getting-ntpsntp-network-time-with-powershell/
    # Construct client NTP time packet to send to specified server
    # (Request Header: [00=No Leap Warning; 011=Version 3; 011=Client Mode]; 00011011 = 0x1B)
    
    [Byte[]]$NtpData = ,0 * 48
    $NtpData[0] = 0x1B  
    
    $Socket = New-Object Net.Sockets.Socket([Net.Sockets.AddressFamily]::InterNetwork,
    [Net.Sockets.SocketType]::Dgram,
    [Net.Sockets.ProtocolType]::Udp)
    
    $Socket.Connect($timeServer,123)
    [Void]$Socket.Send($NtpData)
    [Void]$Socket.Receive($NtpData)    # Returns length – should be 48…
    
    $Socket.Close()
    
    # Decode the received NTP time packet
    
    # We now have the 64-bit NTP time in the last 8 bytes of the received data.
    # The NTP time is the number of seconds since 1/1/1900 and is split into an
    # integer part (top 32 bits) and a fractional part, multipled by 2^32, in the
    # bottom 32 bits.
    
    # Convert Integer and Fractional parts of 64-bit NTP time from byte array
    $IntPart=0;  Foreach ($Byte in $NtpData[40..43]) {$IntPart  = $IntPart  * 256 + $Byte}
    $FracPart=0; Foreach ($Byte in $NtpData[44..47]) {$FracPart = $FracPart * 256 + $Byte}
    
    # Convert to Millseconds (convert fractional part by dividing value by 2^32)
    [UInt64]$Milliseconds = $IntPart * 1000 + ($FracPart * 1000 / 0x100000000)
    
    # Create UTC date of 1 Jan 1900, add the NTP offset and convert result to local time
    $ntptime = (New-Object DateTime(1900,1,1,0,0,0,[DateTimeKind]::Utc)).AddMilliseconds($Milliseconds).ToLocalTime()
    $pcTime = Get-date
    
    Write-Debug "     PC Time : $pcTime"
    Write-Debug "    NTP Time : $ntpTime"
    
    $timeDiff = (New-TimeSpan -Start $pcTime -End $ntpTime)
    $minDiff = [math]::abs([math]::Round($timeDiff.TotalMinutes,2))
    
    Write-Debug "  Difference : $minDiff minutes"

if($minDiff -lt 4){
    Write-Debug "Verified VDA time difference: $minDiff minutes"
    $timeTest = 'True'
    } 
else{
    Write-Debug "VDA time out of sync: $minDiff"
    $timeTest = 'False'
}

###############################################################################################################
# Report findings - some tests return $true , other return a value

Write-Output "====================== Citrix VDA Check ======================"
if ($domainTest)      {Write-Output "PASS:       VDA is a member of $domain" }
         else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       VDA is not a domain member"
                       Write-Output "      TRY:  Join domain"
                      }
				  
if ($ddcTest)    {Write-Output "PASS:       DDCs are set: $ddc1 $ddc2" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       DDCs are not set "
				       Write-Output "      TRY:  Manually set DDCs in HKLM:\SOFTWARE\Citrix\VirtualDesktopAgent\ListOfDDCs"
				      }

if ($dnsTest)    {Write-Output "PASS:       DDC forward DNS resolves: $ddcip1 $ddcip2" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       DDCs do not have forward DNS "
				       Write-Output "      TRY:  Update DNS records"
				      }

if ($revDnsTest)    {Write-Output "PASS:       DDC reverse DNS resolves: $ddcrev1 $ddcrev2" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       DDCs do not have reverse DNS "
				       Write-Output "      TRY:  Update DNS records"
				      }

if ($pingTest)    {Write-Output "PASS:       VDA can ping DDC: $ddcip1 $ddcip2" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       DDCs are not reachable "
				       Write-Output "      TRY:  Verify network routes and firewall rules"
				      }

if ($timeTest)    {Write-Output "PASS:       VDA time is synced. Deviation: $minDiff Minutes" }
		 else 	      {Write-Output "`n======!!======!!"
					   Write-Output "FAIL:       VDA time is NOT synced"
				       Write-Output "      TRY:  Sync time with domain controller"
				      }


############## REMEDIATIONS ##############

Write-Output "====================== VDA Remediation ======================"
Write-Output "Automatic remediation not enabled."

# Force sync system time - $timeSync

#  sleep -seconds 20

#    w32tm /config /syncfromflags:DOMHIER
#    w32tm /resync /nowait
#    net stop w32time
#    net start w32time

# Force a Group Policy update if things have changed - $gpUpdate

# gpupdate /force

    
#Purge Kerberos - is this needed? - $kerberosPurge

  #  klist purge      
  #  klist purge_bind 
  #  klist -lh 0 -li 0x3e7 purge


# Restart the BrokerAgent service - $brokerRestart

 #   Stop-Service -Name BrokerAgent
 #   sleep 1
 #   Start-Service -Name BrokerAgent

 # which is better?
 #    net stop brokeragent
 #    net start brokeragent



 
