# GDI objects
# Number of GDI handles per process
$ErrorActionPreference = "Stop"

$sig = @'
[DllImport("User32.dll")]
public static extern int GetGuiResources(IntPtr hProcess, int uiFlags);
'@

Add-Type -MemberDefinition $sig -name NativeMethods -namespace Win32

$processes = [System.Diagnostics.Process]::GetProcesses()
[int]$gdiHandleCount = 0
$AllDetails = @()
ForEach ($p in $processes) {
    Try {
        $gdiHandles = [Win32.NativeMethods]::GetGuiResources($p.Handle, 0)
        If ($gdiHandles -eq 0) {
          continue
        }
        $gdiHandleCount += $gdiHandles
        $HandleList = New-Object PSObject
        $HandleList| add-member -MemberType NoteProperty -Name "Process Name" -Value $p.Name
        $HandleList| Add-Member -MemberType NoteProperty -Name "PID" -Value ($p.Id)
        $HandleList| add-member -MemberType NoteProperty -Name "Handles" -Value $gdiHandles
        $AllDetails += $HandleList
    }
    Catch {
        # Write-Host "Error accessing $p.Name"
    }
}

Write-Host "Processes with GDI handles"
$AllDetails | Sort-Object "Process Name" | ft -auto
Write-Host "Total number of GDI handles: $gdiHandleCount"

