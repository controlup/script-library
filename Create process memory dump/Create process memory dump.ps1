#requires -version 3
<#
    Create a dump file for a process

    Modification History:

    25/08/19  GRL   Dump parent process if process is werfault

    @guyrleech 2018
#>

[int]$thePid = $args[0]
[string]$dumpFilePath = $args[1]

## http://sysadminconcombre.blogspot.com/2015/09/powershell-dbghelp-reflexion-to-write.html
$MethodDefinition = @'
[DllImport("DbgHelp.dll", CharSet = CharSet.Unicode)]
public static extern bool MiniDumpWriteDump(
    IntPtr hProcess,
    uint processId,
    IntPtr hFile,
    uint dumpType,
    IntPtr expParam,
    IntPtr userStreamParam,
    IntPtr callbackParam
    );
'@

$dbghelp = Add-Type -MemberDefinition $MethodDefinition -Name 'dbghelp' -Namespace 'Win32' -PassThru

[uint32]$miniDumpWithFullMemory = 2 ## https://docs.microsoft.com/en-gb/windows/desktop/api/minidumpapiset/ne-minidumpapiset-_minidump_type
if( ! ( $process = Get-Process -Id $thePid ) )
{
    Throw "Unable to locate process PID $thePid"
}

if( $process.Name -eq 'werfault' )
{
    if( $parentProcess = Get-CimInstance -ClassName win32_process -Filter "ProcessId = '$thePid'" | Select-Object -ExpandProperty ParentProcessId )
    {
        if( ! ( $process = Get-Process -Id $parentProcess ) )
        {
            Throw "Failed to get parent process id $parentProcess for werfault pid $thePid"
        }
        else
        {
            Write-Output -InputObject "Process $thePid is werfault.exe so dumping parent which is $($process.Name) (pid $($process.Id))"
        }
    }
    else
    {
        Throw "Failed to get parent process id for werfault pid $thePid"
    }
}

if( ! ( $processDumpPath = Join-Path -Path $dumpFilePath -ChildPath "$($process.Name).$thePid.$(Get-Date -Format 'yyMMdd.HHmmss').dmp" -ErrorAction SilentlyContinue ) )
{
    Throw "Path `"$dumpFilePath`" for dump file is invalid"
}

$props = $null

try
{
    $props = Get-ItemProperty -Path $processDumpPath -ErrorAction SilentlyContinue
}
catch
{
    $props = $null
}

if( $props )
{
    Write-Error ("Cannot write dump as dump file `"{0}`" already exists - created {1}, size {2} MB" -f $processDumpPath , (Get-Date $props.CreationTime -Format G) , [math]::Round( $props.Length / 1MB ) )
    exit 1
}

if( ! ( $fileStream = New-Object IO.FileStream( $processDumpPath , [IO.FileMode]::Create) ))
{
    Throw "Failed to open dump file `"$processDumpPath`" for writing"
}

try
{
    if( ! ( $result = $dbghelp::MiniDumpWriteDump( $process.Handle , $process.Id , $fileStream.SafeFileHandle.DangerousGetHandle() , $miniDumpWithFullMemory ,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero) ) )
    {
        Write-Error "Error : cannot dump the process PID $thePid"
    }
    else
    {
        if( $props = Get-ItemProperty -Path $processDumpPath -ErrorAction SilentlyContinue )
        {
            Write-Output ("Dump file `"{0}`" created ok, size {1} MB" -f $processDumpPath , [math]::Round( $props.Length / 1MB ) )
        }
        else
        {
            Write-Error "MiniDumpWriteDump() appeared to succeed but dump file `"$processDumpPath`" does not exist"
        }
    }
}
catch
{
    Write-Error "Error : cannot dump the process PID $thePid - $($_.Exception.Message)"
}
finally
{
    $fileStream.Close()
}

