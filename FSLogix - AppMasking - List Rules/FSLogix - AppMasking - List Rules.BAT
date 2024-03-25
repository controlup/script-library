<#
    .SYNOPSIS
        Reports the FSLogix currently-loaded App Masking rules

    .DESCRIPTION
        Reports the FSLogix currently-loaded App Masking rules, implemented by calling 'frx.exe list-rules'

    .NOTES
        If no rules are defined, the script returns:
            No rules

    .CONTEXT
        Session/Machine

    .MODIFICATION_HISTORY
        Created TTYE : 2023-08-16

    AUTHOR: Trentent Tye
#>

$ProgramFiles =[Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)

[string]$FSLogixRulePath = Join-Path -Path $ProgramFiles -ChildPath "FSLogix\Apps\Rules"
[string]$FRXPath = Join-Path -Path $ProgramFiles -ChildPath 'FSLogix\Apps\frx.exe'

& $FRXPath list-rules -verbose

