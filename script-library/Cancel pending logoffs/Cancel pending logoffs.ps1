#requires -version 3.0
<#
    Find running PowerShell processes hosting our logoff SBA and kill it.
    Run if the original script timed out due to a long grace period specified but you no longer need to log users off

    @guyrleech 2018
#>

[string]$message = $args[0]
[string]$owner = Get-Process -Id $pid -IncludeUserName | Select -ExpandProperty Username
[int]$killed = 0

Get-Item -Path ( Join-Path -Path $env:temp -ChildPath "ControlUp.Logoffs.*" ) | ForEach-Object `
{
    $thisPid = ($_.Name -split '\.')[-1]
    $process = Get-Process -Id $thisPid -ErrorAction SilentlyContinue -IncludeUserName
    if( $process )
    {
        if( $process.Username -eq $owner -and $process.Name -eq 'powershell' ) ## can't match arguments as variable
        {
            $parent = Get-Process -Id (Get-CimInstance -ClassName Win32_Process -Filter "processid = $($process.id)"|select -ExpandProperty ParentProcessId) -ErrorAction SilentlyContinue
            if( $parent -and $parent.Name -eq 'cuAgent' )
            {
                $terminated = Stop-Process -InputObject $process -Confirm:$false -PassThru
                if( $terminated -and $terminated.HasExited )
                {
                    $killed++
                    Remove-Item -Path $_.FullName -Force
                }
                else
                {
                    Write-Warning "Failed to terminate process id $($process.Id)"
                }
            }
        }
    }
}

if( ! $killed )
{
    Write-Warning "Found no processes to kill"
}
else
{
    Write-Output "Killed $killed processes"
}

if( $killed -and ! [string]::IsNullOrEmpty( $message ) )
{
    $msgProcess = Start-Process -FilePath 'msg.exe' -ArgumentList "* $($message -replace '\\n' , "`n")" -PassThru -Wait -ErrorAction Stop -WindowStyle Hidden
    if( ! $msgProcess )
    {
        Throw $error[0]
    }
}
