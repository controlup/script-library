@if exist "%ProgramFiles(x86)%\Citrix\system32\auditlog.exe" (
	cd "%ProgramFiles(x86)%\Citrix\system32" && auditlog.exe /before:%1 /after:%2 /time
	) else If exist "%ProgramFiles%\Citrix\system32\auditlog.exe" (
	cd "%ProgramFiles%\Citrix\system32" && auditlog.exe /before:%1 /after:%2 /time
	) else (echo This is not a Citrix XenApp server. 1>&2 && Exit /B 1 )

