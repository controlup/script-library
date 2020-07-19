REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientId /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientIdValidation /f
net stop "Windows Update"
move %windir%\windowsupdate.log %windir%\windowsupdate.old.log
move %windir%\SoftwareDistribution %windir%\SoftwareDistributionold
regsvr32 /s atl.dll
regsvr32 /s wucltui.dll
regsvr32 /s wups.dll
regsvr32 /s wuaueng.dll
regsvr32 /s wuapi.dll
regsvr32 /s msxml3.dll
regsvr32 /s mssip32.dll
regsvr32 /s initpki.dll
regsvr32 /s softpub.dll
net start "Windows Update"
gpupdate /force
wuauclt /resetauthorization /detectnow
wuauclt /reportnow
echo DONE

