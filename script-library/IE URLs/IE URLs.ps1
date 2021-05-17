<#
.SYNOPSIS
    Retrieves IE process IDs and the URLs associated with them..
.DESCRIPTION
    Retrieves IE process IDs and the URLs associated with them..
    This script does not fully support published applications. (Published Apps will not show the actual URL.) 
    You may also get unusual or no results if IE is at a site which uses Unicode characters.
.PARAMETER Identity
   The name of the server being taken out of maintenance mode - automatically supplied by CU
#>

<#
Credits to: Tome Tanasovski
http://powertoe.wordpress.com/2010/11/10/finding-the-thread-pid-that-belongs-to-a-tab-in-ie-8-with-powershell/
Credits to: Tobias
http://powershell.com/cs/forums/t/8982.aspx
#>

If (!(Get-Process iexplore -ErrorAction SilentlyContinue)) {
    Write-Error "IE is not running in this session." -ErrorAction Continue
    Exit 1
}

$IETabList = @()
$IEPIDList = @()
$PIDTabList = @()

Function Get-PID()
{
$sig = @"
[DllImport("user32.dll", SetLastError=true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
 
[DllImport("user32.dll")]
public static extern IntPtr GetTopWindow(IntPtr hWnd);
 
[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

public enum GetWindow_Cmd : uint {
    GW_HWNDFIRST = 0,
    GW_HWNDLAST = 1,
    GW_HWNDNEXT = 2,
    GW_HWNDPREV = 3,
    GW_OWNER = 4,
    GW_CHILD = 5,
    GW_ENABLEDPOPUP = 6
}
 
[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
 
[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern int GetWindowTextLength(IntPtr hWnd);
"@

    Add-Type -MemberDefinition $sig -Namespace User32 -Name Util -UsingNamespace System.Text

    $iethreads = get-process iexplore | Where-Object {!$_.MainWindowTitle} | ForEach-Object {$_.ID}

    $p=0
    $window = [User32.Util]::GetTopWindow(0)

    while ($window -ne 0) {
        [User32.util]::GetWindowThreadProcessId($window, [ref]$p) | Out-Null
        if ($iethreads -contains $p) {
            $length = [User32.Util]::GetWindowTextLength($window)
            if ($length -gt 0) {
                $string = New-Object System.Text.Stringbuilder 1024
                [User32.Util]::GetWindowText($window,$string,($length+1)) | Out-Null
                if ($string.tostring() -notmatch '^MSCTFIME UI$|^Default IME$|^SysFader$|^MCI command handling window$|^Msg$|^ToolTip$|^DDE Server Window') {
                    $script:IEPIDList += $(new-object psobject -Property @{PID = $p;Title = $string.tostring()})
                }
            }
        }
    $window = [User32.Util]::GetWindow($window, 2)
    }
}

Function Get-URL()
{
    Try {
        $shell = New-Object -ComObject Shell.Application
        $shell.Windows() | Where-Object { $_.Type -eq 'HTML Document' } |
            Select-Object LocationName, LocationURL | ForEach-Object { $script:IETabList += $_ }
    } Catch {
        Write-Host "This is likely a published application and not fully supported in this script. Here is all the information we could gather:"
        $IEPIDList | fl
        Exit 1
    }
}


# Main

Get-PID
Get-URL

Foreach ($Tab in $IETabList)
{
    Foreach ($Process in $IEPIDList)
    {
        If ($Process.Title -match [regex]::escape($Tab.LocationName))
        {
            $PIDTabList += $(new-object psobject -Property @{PID = $Process.PID;URL = $Tab.LocationURL;Title = $Tab.LocationName})
        }
    }
}

If ($PIDTabList) {
    $PIDTabList | fl
    If ($IETabList.Count -ne $PIDTabList.Count) {
        Write-Host "There is at least one tab for which we are not able to determine the PID."
        Write-host "Here is the complete tab list for your inspection."
        $IETabList | fl
    }
}
Else {
    Write-Host "There is no match by title. This may be due to Unicode characters in the tab title."
    Write-Host "The complete output of the tables is below for inspection:"
    Write-Host "IETabList: "
    $IETabList | fl
    Write-Host "IEPIDList: "
    $IEPIDList | fl
    Exit 1
}

