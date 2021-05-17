#requires -version 3
$ErrorActionPreference = 'Stop'

<#
    .SYNOPSIS
    Gets the files in the Datastore

    .DESCRIPTION
    This script gets the files in a VMWare Datastore, filterred by the last time they were modified and size.

    .PARAMETER strVCenter
    The name of the vcenter server that will be connected to to run the PowerCLI commands

    .PARAMETER strDatastoreId
    The name of the Datastore the action is to be performed on, will be passed from ControlUp Console.

    .PARAMETER intFileNotModifiedDays
    Minimum days the file has nog changed.

    .PARAMETER intFileSizeMinimumGB
    Minimum size of the file in GB.

    .EXAMPLE
    Example is not relevant as this script will be called through ControlUp Console

    .NOTES
    VMware PowerCLI Core needs to be installed on the machine running the script.
    Loading VMWare PowerCLI will result in a 'Join our CEIP' message. In order to disable these in the future run the following commands on the target system:
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false (or $true, if that's your kind of thing)
    
    Based on work by Luc Dekens
    http://www.lucd.info
#>

[string]$strVCentername = $args[0]
[string]$strDatastoreName = $args[1]
[int]$intFileNotModifiedDays = $args[2]
[int]$intFileSizeMinimumGB = $args[3]

Function Feedback {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$Message,
        [Parameter(Mandatory = $false,
            Position = 1)]
        $Exception,
        [switch]$Oops
    )

    # This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If (!$Exception -and !$Oops) {
        # Write content of feedback string
        Write-Host $Message -ForegroundColor 'Green'
    }

    # If an error occured report it, and exit the script with ErrorLevel 1
    Else {
        # Write content of feedback string but to the error stream
        $Host.UI.WriteErrorLine($Message) 
        
        # Display error details
        If ($Exception) {
            $Host.UI.WriteErrorLine("Exception detail:`n$Exception")
        }
        
        # Exit errorlevel 1
        Exit 1
    }
}

function Load-VMWareModules {
    <# Imports VMware PowerCLI modules, with a -Prefix $Prefix is supplied (desirable to avoid conflict with Hyper-V cmdlets)
      NOTES:
      - The required modules to be loaded are passed as an array.
      - If the PowerCLI versions is below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
    #>

    param (    
        [parameter(Mandatory = $true,
            ValueFromPipeline = $false)]
        [array]$Components
    )

    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    # Import each module, if Import-Module fails try Add-PSSnapin
    foreach ($component in $Components) {
        try {
            $null = Import-Module -Name VMware.$component
        }
        catch {
            try {
                $null = Add-PSSnapin -Name VMware
            }
            catch {
                Write-Host 'The required VMWare PowerCLI components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 6.5 or higher preferred) is installed and available for the user running the script.'
                Exit 1
            }
        }
    }
}

Function Connect-VCenterServer {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        [string]$VCenterName
    )
    Try {
        # Connect to VCenter server
        Connect-VIServer -Server $VCenterName -WarningAction SilentlyContinue -Force
    }
    Catch {
        Feedback -Message "There was a problem connecting to VCenter server $VCenterName. Please correct the error and try again." -Exception $_
    }
}

Function Disconnect-VCenterServer {
    Param (
        [Parameter(Mandatory = $true,
            Position = 0)]
        $VCenter
    )
    # This function closes the connection with the VCenter server 'VCenter'
    try {
        # Disconnect from the VCenter server
        Disconnect-VIServer -Server $VCenter -Confirm:$false
    }
    catch {
        Feedback -Message "There was a problem disconnecting from VCenter server $($VCenter.name)" -Exception $_
    }
}

function Get-VMwareDS {
    param (
        [parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [string]$DatastoreName,
        [parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 1)]
        $VCenter
    )

    # Get the Datastore
    try {
        Get-Datastore -Name $DatastoreName -Server $VCenter
    }
    catch {
        Feedback -Message "Datastore $DatastoreName could not be retreived." -Exception $_
    }
}

function New-VMWareDatastoreSearchSpecification {
    param (
        [parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            Position = 0)]
        [bool]$fileType = $false,
        [parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            Position = 1)]
        [bool]$fileOwner = $false,
        [parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            Position = 2)]
        [bool]$fileSize = $false,
        [parameter(Mandatory = $false,
            ValueFromPipeline = $false,
            Position = 3)]
        [bool]$modification = $false
    )

    # Try to create the SearchSpecification object
    try {
        # Create new search specification for SearchDatastoreSubFolders method
        $SearchSpecification = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec

        # Create the file query flags to return Size and last modified
        $fileQueryFlags = New-Object VMware.Vim.FileQueryFlags
        $fileQueryFlags.fileSize = $fileSize
        $fileQueryFlags.fileType = $fileType
        $fileQueryFlags.fileOwner = $fileOwner
        $fileQueryFlags.modification = $modification

        # Set the flags on the search specification
        $SearchSpecification.Details = $fileQueryFlags

        $SearchSpecification
    }
    catch {
        Feedback -Message 'The HostDatastoreBrowserSearchSpec object (user for searching the datastore) could not be created.' -Exception $_
    }
}

# Check all the arguments have been passsed
if ($args.Count -ne 4) {
    Feedback -Oops  "The script did not get the correct amount of arguments from the Console. This can occur if you are not connected to the VM's hypervisor.`nPlease connect to the hypervisor in the ControlUp Console and try again."
}

# Import the VMWare PowerCLI module
Load-VMwareModules -Components @('VimAutomation.Core')

# Increase web request timeout to three hours as searching the datastore can be slow
$null = Set-PowerCLIConfiguration -Scope Session -WebOperationTimeoutSeconds 10800 -Confirm:$false

# Connect to VCenter server for VMWare
$VCenter = Connect-VCenterServer -VCenterName $strVCenterName

# Get the datastore
$Datastore = Get-VMWareDS -DatastoreName $strDatastoreName -VCenter $VCenter

# Check only one datastore was returned
if ($Datastore.count -gt 1) {
    Feedback -Message "More than one datastore was found using the name $strDatastoreName. The script cannot continue" -Oops
}

# Set View
$DatastoreView = Get-View -Id $Datastore.Id -Server $Vcenter

# Set the browser
$DatastoreBrowser = Get-View $DataStoreView.browser -Server $Vcenter

# Get current date
[datetime]$dtNow = Get-Date

# Set the date the file has to OLDER than
[datetime]$dtBeforeDate = $dtNow.AddDays(-$intFileNotModifiedDays)

# Set root path to search from, we want to search the entire datastore
$RootPath = ("[" + $Datastore.Name + "]")

# Set up the Datastore browser for searching entire datastore including fileType
$SearchSpecification = New-VMWareDatastoreSearchSpecification -fileSize $true -modification $true -fileType $true
    
# Do the search
$SearchResult = $DatastoreBrowser.SearchDatastoreSubFolders($RootPath, $SearchSpecification)

# Now get the files that are old and large enough in that result
$DatastoreFiles = Foreach ($obj in $SearchResult) {
    $objEx = $obj | Select-Object -expandproperty File
    Foreach ($File in $objEx | Where-Object { ($_.Modification.Date -lt $dtBeforeDate) -and (($_.FileSize / 1gb) -gt $intFileSizeMinimumGB) }) {
        [pscustomobject][ordered]@{
            # Strip the datastore name from the File path as it is redundant
            File         = "$($obj.Folderpath.Replace("$RootPath ",'/'))$($File.Path)"
            Size         = $File.FileSize
            LastModified = $File.Modification.Date
            Type         = $File.ToString().Replace('VMware.Vim.', '').TrimEnd('Info')
        }
    }
}

# Set up the output with formating
[array]$TableProperties = @(
    @{ Label = 'File'; Expression = { $_.File }; Alignment = 'Left' }
    @{ Label = 'SizeGB'; Expression = { [math]::Round($_.Size / 1gb, 2) }; Alignment = 'Left' }
    @{ Label = 'Days since modified'; Expression = { ($dtNow - $_.LastModified).Days }; Alignment = 'Left' }
    @{ Label = 'Type'; Expression = { $_.Type }; Alignment = 'Left' }
)

# Set the the size of the PS Buffer
$PSWindow = (Get-Host).UI.RawUI
$WideDimensions = $PSWindow.BufferSize
$WideDimensions.Width = '400'
$PSWindow.BufferSize = $WideDimensions

# Output the result
$DatastoreFiles | Sort-Object 'Size' -Descending | Format-Table -Property $TableProperties -Autosize

# Disconnect from the VCenter server
Disconnect-VCenterServer $VCenter

