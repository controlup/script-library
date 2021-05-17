function Stop-ProcessByForce {
[CmdletBinding()]
Param(
[String]$Source = "https://live.sysinternals.com/pskill.exe",

[string]$Destination = "$env:TEMP\pskill.exe",

[Parameter(Mandatory=$true)]
[int]$ProcessID
)

try {
    Invoke-WebRequest -Uri $Source -OutFile $Destination
}
catch {
    Throw "Could not download pskill.exe from $Source"
}

Start-Process -FilePath $Destination -ArgumentList ($ProcessID, "/accepteula")

}

if ((Get-Command "pskill.exe" -ErrorAction SilentlyContinue) -eq $null) 
{ 
    Stop-ProcessByForce -ProcessID $args[0]
}
else {
    Start-Process -FilePath "pskill.exe" -ArgumentList ($args[0], "/accepteula")
}
