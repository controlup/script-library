<#
    .SYNOPSIS
        Adds or removes Windows Features from Windows Server Operating Systems

    .DESCRIPTION
        Adds or removes Windows Features from Windows Server Operating Systems

    .EXAMPLE
        . .\Add-Or-RemoveWindowsFeatures.ps1 -Workstation False -ComputerName W2019-001 -Source \\DS1813.bottheory.local\fileshare\OS\W2K19\sources\sxs
        Opens a dialog prompting you to select features to add or remove.

    .PARAMETER  <Workstation <string>>
        True if this is a workstation OS or False if this is a Server operating system.

    .PARAMETER  <ComputerName <string>>
        The name of the computer to run this script against

    .PARAMETER  <Source <string>>
        Specify an optional 'source' path for adding new Windows Features. Not required for most features, but it is required when adding .NetFramework 3.5, for example.

    .CONTEXT
        Console

    .MODIFICATION_HISTORY
        Created TTYE : 2020-10-01


    AUTHOR: Trentent Tye
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$true,HelpMessage='Is this a workstation operating system?')][ValidateNotNullOrEmpty()]              [string]$Workstation = "False",
    [Parameter(Mandatory=$true,HelpMessage='Name of the target machine to add or remove features')][ValidateNotNullOrEmpty()] [string]$ComputerName,
    [Parameter(Mandatory=$false,HelpMessage='Source path for Windows features')]                                              [string]$Source
)


$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'
#Set-StrictMode -Version Latest


if (-not(Test-Path env:source)) {
    $source = $null
}

Write-Verbose  -Message "Parameters: Workstation = $workstation `n                     ComputerName = $computerName `n                     Source = $source"


function Display-ToastNotification ([string]$title, [string]$progressTitle, [string]$progressValue, [string]$progressStatus, [string]$progressValueStringOverride, [int]$sequenceID) {
    $app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] ##need these two nulls to preload the assemblies
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $Template = [Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText01
    [xml]$ToastTemplate = @"
    <toast launch="app-defined-string" duration="long">
        <visual>
        <binding template="ToastGeneric">
            <text>{title}</text>
            <progress title="{progressTitle}" value="{progressValue}" status="{progressStatus}" valueStringOverride="{progressValueStringOverride}"/>
        </binding>
        </visual>
        <audio silent="true" />
        <actions>
        <action activationType="background" content="OK" arguments="later"/>
        </actions>
    </toast>
"@
    $ToastXml = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument
    $ToastXml.LoadXml($ToastTemplate.OuterXml)
    $toastNotification = [Windows.UI.Notifications.ToastNotification]::new($ToastXml)
    $toastNotification.Tag   = "ControlUp-SBA-AddRemovePrograms"
    $toastNotification.Group = "AppRemoveProgramsGroup"

    $DataDictionary = New-Object 'system.collections.generic.dictionary[string,string]'
    $DataDictionary.Add("title", "$title")
    $DataDictionary.Add("progressTitle", "$progressTitle")
    $DataDictionary.Add("progressValue", "$progressValue")
    $DataDictionary.Add("progressStatus", "$progressStatus")
    $DataDictionary.Add("progressValueStringOverride", "$progressValueStringOverride")

    $toastNotification.Data = [Windows.UI.Notifications.NotificationData]::new($DataDictionary, $sequenceID)
    $Global:notify = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app)
    $notify.Show($toastNotification)
}

function Update-ToastNotification ([string]$title, [string]$progressTitle, [string]$progressValue, [string]$progressStatus, [string]$progressValueStringOverride, [int]$sequenceID) {
    $app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] ##need these two nulls to preload the assemblies
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $Template = [Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText01
    [xml]$ToastTemplate = @"
    <toast launch="app-defined-string" duration="long">
        <visual>
        <binding template="ToastGeneric">
            <text>{title}</text>
            <progress title="{progressTitle}" value="{progressValue}" status="{progressStatus}" valueStringOverride="{progressValueStringOverride}"/>
        </binding>
        </visual>
        <audio silent="true" />
        <actions>
        <action activationType="background" content="OK" arguments="later"/>
        </actions>
    </toast>
"@
    $ToastXml = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument
    $ToastXml.LoadXml($ToastTemplate.OuterXml)
    $toastNotification = [Windows.UI.Notifications.ToastNotification]::new($ToastXml)
    $toastNotification.Tag   = "ControlUp-SBA-AddRemovePrograms"
    $toastNotification.Group = "AppRemoveProgramsGroup"
 
    $DataDictionary = New-Object 'system.collections.generic.dictionary[string,string]'
    $DataDictionary.Add("title", "$title")
    $DataDictionary.Add("progressTitle", "$progressTitle")
    $DataDictionary.Add("progressValue", "$progressValue")
    $DataDictionary.Add("progressStatus", "$progressStatus")
    $DataDictionary.Add("progressValueStringOverride", "$progressValueStringOverride")

    $ToastData = [Windows.UI.Notifications.NotificationData]::new($DataDictionary, $sequenceID)
    $null = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Update($ToastData, $toastNotification.Tag,$toastNotification.Group)
}



#Check if it's is a workstation OS and error/exit if true
if ($Workstation -eq "True") {
    Write-Verbose -Message "Desktop operating system detected."
    $features = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-WindowsOptionalFeature -Online } | Sort-Object -Property FeatureName
    $selections = $features | Select-Object -Property "FeatureName","State" | Out-GridView -Title "$ComputerName : Add or Remove Windows Features" -PassThru
} else {
    Write-Verbose -Message "Server operating system detected."
    $features = Get-WindowsFeature -ComputerName $ComputerName

    #going to create a custom object so features will sort correctly in out-gridview
    $FeaturesObj = New-Object System.Collections.Generic.List[PSObject]
    $count = 0
    foreach ($feature in $features) {
        $count++

        switch ($feature.Depth) {
            1  { $spaces = "" }
            2  { $spaces = "    " }
            3  { $spaces = "        " }
            4  { $spaces = "            " }
            5  { $spaces = "                " }
        }

        if ($feature.Installed -eq $true) {
            $displayName = "$spaces[X] $($feature.DisplayName)"
        } else {
            $displayName = "$spaces[ ] $($feature.DisplayName)"
        }

        $FeaturesObj.Add(
            [PSCustomObject]@{
                Item = $count
                DisplayName = $displayName
                OriginalName = $feature.DisplayName
                Name = $feature.Name
                InstallState = $feature.InstallState
            }
        )
    }
    $featureGrid = $FeaturesObj | Select-Object -Property Item,DisplayName,Name,InstallState | Out-GridView -Title "$ComputerName : Add or Remove Windows Features" -PassThru
    $selections = New-Object System.Collections.Generic.List[PSObject]
    foreach ($featureSelection in $featureGrid) {
        $selections.Add( $FeaturesObj.Where{$_.Item -eq $featureSelection.Item} )
    }
}

Write-Verbose "$($selections | Format-Table -AutoSize | Out-String)"


$Results = New-Object System.Collections.Generic.List[PSObject]
$Count = 0
$SelectionsCount = ($selections | Measure-Object).Count
$NotificationDisplayed = $false

If ($SelectionsCount -eq 0) {
    Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "No Features Selected" -progressValue "0.0" -progressStatus "" -progressValueStringOverride "" -sequenceID 0
    Exit 0
}

if ($Workstation -eq "False") {
    ## Server OS Feature Results Object
    foreach ($selection in $selections) {
        Write-Verbose "Feature: $($selection.name)"
        $Count++
        switch -Wildcard ($selection.InstallState) {
            {($_ -eq "Available") -or ($_ -eq "Removed")}{
                if ([string]::IsNullOrEmpty($source)) {
                    if (-not($NotificationDisplayed)) {
                        Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                        $NotificationDisplayed = $true
                    } else {
                        Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID $count
                    }
                    $Result = Add-WindowsFeature -Name $selection.name -ComputerName $ComputerName -ErrorAction SilentlyContinue -ErrorVariable errmsg
                    if ($Result.FeatureResult.Length -eq 0) {
                        $Success = ""
                    } else {
                        $Success = $result.FeatureResult.success
                    }
                    if ($errmsg -ne $null) {
                        Write-Verbose -Message "Error found for $($selection.Name)"
                        $Results.Add(
                            [PSCustomObject]@{
                                DisplayName = $selection.OriginalName
                                FeatureName = $selection.Name
                                RestartNeeded = $Result.RestartNeeded
                                State = $result.exitcode
                                Success = $Success
                            }
                        )
                    } else {
                        Write-Verbose -Message "Adding Result: $($selection.Name) - $($Result.RestartNeeded) - Enabled"
                        $Results.Add(
                            [PSCustomObject]@{
                                DisplayName = $selection.OriginalName
                                FeatureName = $selection.Name
                                RestartNeeded = $Result.RestartNeeded
                                State = $result.exitcode
                                Success = $Success
                            }
                        )
                    }
                } else {
                    if (-not($NotificationDisplayed)) {
                        Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                        $NotificationDisplayed = $true
                    } else {
                        Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID $count
                    }
                    $Result = Add-WindowsFeature -Name $selection.name -Source $Source -ComputerName $ComputerName -ErrorAction SilentlyContinue -ErrorVariable errmsg
                    if ($Result.FeatureResult.Length -eq 0) {
                        $Success = ""
                    } else {
                        $Success = $result.FeatureResult.success
                    }
                    $Results.Add(
                        [PSCustomObject]@{
                            DisplayName = $selection.OriginalName
                            FeatureName = $selection.Name
                            RestartNeeded = $Result.RestartNeeded
                            State = $result.exitcode
                            Success = $Success
                        }
                    )
                }
            }
            "Installed" {
                if (-not($NotificationDisplayed)) {
                    Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                    $NotificationDisplayed = $true
                } else {
                    Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID $count
                }
                $Result = Remove-WindowsFeature -Name $selection.name -ComputerName $ComputerName -Remove -ErrorAction SilentlyContinue -ErrorVariable errmsg
                Write-Verbose "$($result.FeatureResult)"
                if ($Result.FeatureResult.Length -eq 0) {
                    $Success = ""
                } else {
                    $Success = $result.FeatureResult.success
                }
                $Results.Add(
                    [PSCustomObject]@{
                        DisplayName = $selection.OriginalName
                        FeatureName = $selection.Name
                        RestartNeeded = $Result.RestartNeeded
                        State = $result.exitcode
                        Success = $Success
                    }
                )
            }
            "*Pending" {
                if (-not($NotificationDisplayed)) {
                    Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                    $NotificationDisplayed = $true
                } else {
                    Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName-$($selection.OriginalName)" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.name)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID $Count
                }
                $Results.Add(
                    [PSCustomObject]@{
                        DisplayName = $selection.OriginalName
                        FeatureName = $selection.Name
                        RestartNeeded = "Yes"
                        State = $selection.InstallState
                        Success = ""
                    }
                )
            }
        }
    }
} else {
    ## Client OS Feature Results Object
    foreach ($selection in $selections) {
    Write-Verbose "Feature: $($selection.FeatureName)"
    $count++
        switch -Wildcard ($selection.State) {
            "Disabled" {
                if ([string]::IsNullOrEmpty($source)) {
                    if (-not($NotificationDisplayed)) {
                        Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                        $NotificationDisplayed = $true
                    } else {
                        Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                    }
                    $Result = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Enable-WindowsOptionalFeature -Online -FeatureName $args[0] -NoRestart } -ArgumentList "$($selection.FeatureName)" -ErrorAction SilentlyContinue -ErrorVariable errmsg 2>$null
                    if ($errmsg -ne $null) {
                        Write-Verbose -Message "Error found for $($selection.FeatureName)"
                        $Results.Add(
                            [PSCustomObject]@{
                                FeatureName = $selection.FeatureName
                                RestartNeeded = ""
                                State = "$errmsg"
                            }
                        )
                    } else {
                        Write-Verbose -Message "Adding Result: $($selection.FeatureName) - $($Result.RestartNeeded) - Enabled"
                        $Results.Add(
                            [PSCustomObject]@{
                                FeatureName = $selection.FeatureName
                                RestartNeeded = $Result.RestartNeeded
                                State = "Enabled"
                            }
                        )
                    }
                } else {
                    if (-not($NotificationDisplayed)) {
                        Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                        $NotificationDisplayed = $true
                    } else {
                        Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                    }
                    $Result = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Enable-WindowsOptionalFeature -Online -FeatureName $args[0] -Source $args[1] -NoRestart } -ArgumentList @("$($selection.FeatureName)","$Source") -ErrorAction SilentlyContinue -ErrorVariable errmsg 2>$null
                    if ($errmsg -ne $null) {
                        Write-Verbose -Message "Error found for $($selection.FeatureName)"
                        $Results.Add(
                            [PSCustomObject]@{
                                FeatureName = $selection.FeatureName
                                RestartNeeded = ""
                                State = "$errmsg"
                            }
                        )
                    } else {
                        $Results.Add(
                            [PSCustomObject]@{
                                FeatureName = $selection.FeatureName
                                RestartNeeded = $Result.RestartNeeded
                                State = "Enabled"
                            }
                        )
                    }
                }
            }
            "Enabled" {
                if (-not($NotificationDisplayed)) {
                        Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                        $NotificationDisplayed = $true
                    } else {
                        Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                    }
                $Result = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Disable-WindowsOptionalFeature -Online -FeatureName $args[0] -Remove -NoRestart } -ArgumentList @("$($selection.FeatureName)") -ErrorAction SilentlyContinue -ErrorVariable errmsg 2>$null
                
                if ($errmsg -ne $null) {
                    Write-Verbose -Message "Error found for $($selection.FeatureName)"
                    $Results.Add(
                        [PSCustomObject]@{
                            FeatureName = $selection.FeatureName
                            RestartNeeded = ""
                            State = "$errmsg"
                        }
                    )
                } else {
                    $Results.Add(
                        [PSCustomObject]@{
                            FeatureName = $selection.FeatureName
                            RestartNeeded = $Result.RestartNeeded
                            State = "Disabled"
                        }
                    )
                }
            }
            "*Pending" {
                if (-not($NotificationDisplayed)) {
                    Display-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                    $NotificationDisplayed = $true
                } else {
                    Update-ToastNotification -title "ControlUp SBA : Add/Remove Features" -progressTitle "$computerName" -progressValue "$([math]::Round($($count/$SelectionsCount),1))" -progressStatus "$($selection.FeatureName)" -progressValueStringOverride "$count/$SelectionsCount" -sequenceID 0
                }
               
                $Results.Add(
                    [PSCustomObject]@{
                        FeatureName = $selection.FeatureName
                        RestartNeeded = ""
                        State = $Selection.state
                    }
                )
            }
        }
    }
}

$Results | Out-GridView -Title "$ComputerName : Add or Remove Windows Features Results" -Wait

