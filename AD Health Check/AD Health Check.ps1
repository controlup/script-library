#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Perform an Active Directory Health Check.
.DESCRIPTION
	Perform an Active Directory Health Check based on LDAP queries.
	These are originally based on Jeff Wouters personal best practices.
	No rights can be claimed by this report!

	Founding guidelines for all checks in this script:
	*) Must work for all domains in a forest tree.
	*) Must work without module dependencies, except for the PowerShell core modules.
	*) Must work without Administrator privileges.
	
	You will see a lot of redirection to streams in this script. i.e. 3>$Null, 4>$Null and possibly *>$Null
	This is explained here: 
	https://blogs.technet.microsoft.com/heyscriptingguy/2014/03/30/understanding-streams-redirection-and-write-host-in-powershell/
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER CompanyName
	Companyname to use for the coverpage.
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName
	or HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER Coverpage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
						Subtitle/Subject & Author fields need to be moved 
						after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
						changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
						changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually resized or font 
					changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER MSWord
	SaveAs DOCX file.
	This parameter is set True if no other output format is selected.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be ReportName_2016-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER Sites
	Only perform the checks related to Sites.
.PARAMETER OrganisationalUnit
	Only perform the checks related to OrganisationalUnits.
	This parameter has an alias of OU.
.PARAMETER Users
	Only perform the checks related to Users.
.PARAMETER Computers
	Only perform the checks related to Computers.
.PARAMETER Groups
	Only perform the checks related to Groups.
.PARAMETER All
	Perform all checks.
	This parameter is the default if no other selection parameters are used.
.PARAMETER Log
	Generates a log file for the purpose of troubleshooting.
.PARAMETER Mgmt
	Provides a page at the end of the PDF or DOCX file with information for your manager.
	Listed is the name of the check performed and the number of results found by the check.
.PARAMETER Visible
	Shows Microsoft Word while creating the report.
	This parameter is disabled by default.
.PARAMETER CSV
	For each check, a separate CSV file will be created with the results.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.EXAMPLE
	PS C:\PSScript > .\ADHealthCheck_V2.ps1 -Visible -MSWord

	This will generate a DOCX document with all the checks included.
	Microsoft Word will be visible while creating the DOCX file.
	The file is created at the location of the script that is executed.
.EXAMPLE
	PS C:\PSScript > .\ADHealthCheck_V2.ps1 -Visible -MSWord -Log -CSV

	This will generate a DOCX document with all the checks included.
	Microsoft Word will be visible while creating the DOCX file.
	For each check, a separate CSV file will be created with the results.
	A log file is created for the purpose of troubleshooting.
	All files are created at the location of the script that is executed.
.EXAMPLE
	PS C:\PSScript > .\ADHealthCheck_V2.ps1 -MSWord -Sites -Users -Groups

	This will generate a DOCX document with the checks for Sites, Users and Groups.
.EXAMPLE
	PS C:\PSScript > .\ADHealthCheck_V2.ps1 -Folder \\FileServer\ShareName

	This will generate a DOCX document with all the checks included.
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\ADHealthCheck_V2.ps1 -SmtpServer mail.domain.tld -From ADAdmin@domain.tld -To ITGroup@domain.tld

	Script will use the email server mail.domain.tld, sending from ADAdmin@domain.tld, sending to ITGroup@domain.tld.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\ADHealthCheck_V2.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com

	Script will use the email server smtp.office365.com on port 587 using SSL, sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME        :   AD Health Check.ps1
	AUTHOR      :   Jeff Wouters [MVP Windows PowerShell], Carl Webster and Michael B. Smith
	VERSION     :   2.0
	LAST EDIT   :   8-May-2016

	The Word file generation part of the script is based upon the work done by:

	Carl Webster  | http://www.carlwebster.com | @CarlWebster
	Iain Brighton | http://virtualengine.co.uk | @IainBrighton
	Jeff Wouters  | http://www.jeffwouters.nl  | @JeffWouters

	The Active Directory checks were originally written by:

	Jeff Wouters  | http://www.jeffwouters.nl  | @JeffWouters
	
	Significant Active Directory changes have been implemented by:
	
	Michael B. Smith | http://TheEssentialExchange.com/ | @EssentialExchange
#>

[CmdletBinding( DefaultParameterSetName = 'All', SupportsShouldProcess = $false, ConfirmImpact = 'None' )]
Param(
    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Alias( 'UN' )]
	[ValidateNotNullOrEmpty()]
    [string] $UserName = $env:username,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
	[Alias( 'CN' )]
	[ValidateNotNullOrEmpty()]
    $CompanyName = '',

    [Parameter( Mandatory = $false, Position=1, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, Position=1, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, Position=1, ParameterSetName = 'SMTP' )] 
    [Alias( 'CP' )]
	[ValidateNotNullOrEmpty()]
    [string] $CoverPage = 'Sideline', 

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $MSWord = $false,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $PDF = $false,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
	[Alias( 'ADT' )]
    [Switch] $AddDateTime = $false,
	
    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $Sites,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
	[Alias( 'OU' )]
	[Alias( 'OrganizationalUnit' )]
    [Switch] $OrganisationalUnit,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $Users,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $Computers,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $Groups,

    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $All = $true,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $Log = $false,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )]
	[Alias( 'Management' )]
    [Switch] $Mgmt = $true,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $Visible = $false,

    [Parameter( Mandatory = $false, ParameterSetName = 'Specific' )]
    [Parameter( Mandatory = $false, ParameterSetName = 'All' )]
	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
    [Switch] $CSV = $false,

	[Parameter( Mandatory = $true,Position=2 )] 
	[string] $Folder = '',
	
	[Parameter( Mandatory = $true, ParameterSetName = 'SMTP' )] 
	[string] $SmtpServer = '',

	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
	[int]$SmtpPort = 25,

	[Parameter( Mandatory = $false, ParameterSetName = 'SMTP' )] 
	[Switch] $UseSSL = $false,

	[Parameter( Mandatory = $true, ParameterSetName = 'SMTP' )] 
	[string] $From = '',

	[Parameter( Mandatory = $true, ParameterSetName = 'SMTP' )] 
	[string] $To = '',

	[Parameter( Mandatory = $false )] 
	[Switch] $Dev = $false,
	
	[Parameter( Mandatory = $false )] 
	[Alias( 'SI' )]
	[Switch] $ScriptInfo = $false
)

#region script change log	
#originally written by Jeff Wouters | http://www.jeffwouters.nl | @JeffWouters
# Now maintained by webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#
#Version 2.0 9-May-2016
#	Added alias for AddDateTime of ADT
#	Added alias for CompanyName of CN
#	Added -Dev parameter to create a text file of script errors
#	Added more script information to the console output when script starts
#	Added -ScriptInfo (SI) parameter to create a text file of script information
#	Added support for emailing output report
#	Added support for output folder
#	Added word 2016 support
#	Fixed numerous issues discovered with the latest update to PowerShell V5
#	Fixed several incorrect variable names that kept PDFs from saving in Windows 10 and Office 2013
#	General code cleanup by Michael B. Smith
#	Output to CSV rewritten by Michael B. Smith
#	Removed the 10 second pauses waiting for Word to save and close
#	Removed unused parameters Text, HTML, ComputerName, Hardware
#	Significant Active Directory changes have been implemented by Michael B. Smith
#	Updated help text
#
# Version 1.0 released to the community on July 14, 2014
# http://jeffwouters.nl/index.php/2014/07/an-active-directory-health-check-powershell-script-v1-0/
#endregion

Set-StrictMode -Version 2

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
##$Script:ThisScriptPath = $(Split-Path ((Get-PSCallStack)[0]).ScriptName) -- this is crap after v1
$Script:ThisScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

If($PSBoundParameters.ContainsKey('Log')) 
{
    $Script:LogPath = "$Script:ThisScriptPath\ADHealthCheckTranscript.txt"
    If((Test-Path $Script:LogPath) -eq $true) 
	{
        Write-Verbose "$(Get-Date): Transcript/Log $Script:LogPath already exists"
        $Script:StartLog = $false
    } 
	Else 
	{
        try 
		{
            Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
            Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
            $Script:StartLog = $true
        } 
		catch 
		{
            Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
            $Script:StartLog = $false
        }
    }
}

If($Null -eq $PDF)
{
	$PDF = $False
}
If($Null -eq $MSWord)
{
	$MSWord = $False
}
If($Null -eq $AddDateTime)
{
	$AddDateTime = $False
}
If($Null -eq $Folder)
{
	$Folder = ""
}
If($Null -eq $SmtpServer)
{
	$SmtpServer = ""
}
If($Null -eq $SmtpPort)
{
	$SmtpPort = 25
}
If($Null -eq $UseSSL)
{
	$UseSSL = $False
}
If($Null -eq $From)
{
	$From = ""
}
If($Null -eq $To)
{
	$To = ""
}
If($Null -eq $Dev)
{
	$Dev = $False
}
If($Null -eq $ScriptInfo)
{
	$ScriptInfo = $False
}

If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Folder))
{
	$Folder = ""
}
If(!(Test-Path Variable:SmtpServer))
{
	$SmtpServer = ""
}
If(!(Test-Path Variable:SmtpPort))
{
	$SmtpPort = 25
}
If(!(Test-Path Variable:UseSSL))
{
	$UseSSL = $False
}
If(!(Test-Path Variable:From))
{
	$From = ""
}
If(!(Test-Path Variable:To))
{
	$To = ""
}
If(!(Test-Path Variable:Dev))
{
	$Dev = $False
}
If(!(Test-Path Variable:ScriptInfo))
{
	$ScriptInfo = $False
}

If($Null -eq $MSWord)
{
	If($PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is " $MSWord
		Write-Verbose "$(Get-Date): PDF is " $PDF
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		$ErrorActionPreference = $SaveEAPreference
		Exit
	}
}

If($Dev)
{
	$Error.Clear()
	$pwdPath = $Folder
	If($pwdPath -eq "")
	{
		$pwdpath = $pwd.Path
	}

	[string] $Script:DevErrorFile = Join-Path $pwdPath "ADHealthCheckScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($ScriptInfo)
{
	$pwdPath = $Folder
	If($pwdPath -eq "")
	{
		$pwdpath = $pwd.Path
	}

	[string] $Script:SIFile = Join-Path $pwdPath "ADHealthCheckScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[long]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
}

Function SetWordHashTable
{
	Param(
		[string]$CultureCode
	)

	$hash = @{}
	    
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($CultureCode)
	{
		'ca-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Taula automÃ¡tica 2'
				}
			}

		'da-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatisk tabel 2'
				}
			}

		'de-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatische Tabelle 2'
				}
			}

		'en-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents'  = 'Automatic Table 2'
				}
			}

		'es-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Tabla automÃ¡tica 2'
				}
			}

		'fi-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automaattinen taulukko 2'
				}
			}

		'fr-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Sommaire Automatique 2'
				}
			}

		'nb-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatisk tabell 2'
				}
			}

		'nl-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatische inhoudsopgave 2'
				}
			}

		'pt-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'SumÃ¡rio AutomÃ¡tico 2'
				}
			}

		'sv-'	{
				$hash.($($CultureCode)) = @{
					'Word_TableOfContents' = 'Automatisk innehÃ¥llsfÃ¶rteckning2'
				}
			}

		Default	{$hash.('en-') = @{
					'Word_TableOfContents'  = 'Automatic Table 2'
				}
			}
	}

	$Script:myHash = $hash.$CultureCode

	If($Script:myHash -eq $Null)
	{
		$Script:myHash = $hash.('en-')
	}

	$Script:myHash.Word_NoSpacing = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1 = $wdStyleheading1
	$Script:myHash.Word_Heading2 = $wdStyleheading2
	$Script:myHash.Word_Heading3 = $wdStyleheading3
	$Script:myHash.Word_Heading4 = $wdStyleheading4
	$Script:myHash.Word_TableGrid = $wdTableGrid
}

Function GetCulture
{
	Param(
		[int]$WordValue
	)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray    = ,1027
	$DanishArray     = ,1030
	$DutchArray      = 2067, 1043
	$EnglishArray    = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray    = ,1035
	$FrenchArray     = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray     = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray  = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray    = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray    = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($WordValue)
	{
		{ $CatalanArray    -contains $_ } { $CultureCode = "ca-" }
		{ $DanishArray     -contains $_ } { $CultureCode = "da-" }
		{ $DutchArray      -contains $_ } { $CultureCode = "nl-" }
		{ $EnglishArray    -contains $_ } { $CultureCode = "en-" }
		{ $FinnishArray    -contains $_ } { $CultureCode = "fi-" }
		{ $FrenchArray     -contains $_ } { $CultureCode = "fr-" }
		{ $GermanArray     -contains $_ } { $CultureCode = "de-" }
		{ $NorwegianArray  -contains $_ } { $CultureCode = "nb-" }
		{ $PortugueseArray -contains $_ } { $CultureCode = "pt-" }
		{ $SpanishArray    -contains $_ } { $CultureCode = "es-" }
		{ $SwedishArray    -contains $_ } { $CultureCode = "sv-" }
		Default                           { $CultureCode = "en-" }
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param(
		[int]$xWordVersion, 
		[string]$xCP, 
		[string]$CultureCode
	)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function Stop-WinWord
{
	Write-Debug "***Enter Stop-WinWord"
	
	## determine our login session
	$proc = Get-Process -PID $PID
	If( $null -eq $proc )
	{
		throw "Stop-WinWord: Cannot find process $PID"
	}
	
	$SessionID = $proc.SessionId
	If( $null -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId on $PID is null"
		throw "Can't find a session for pid $PID"
	}

	If( 0 -eq $SessionID )
	{
		Write-Debug "Stop-WinWord: SessionId is 0 -- that is a bug"
		throw "SessionId is zero for pid $PID"
	}
	
	#Find out if winword is running in our session
	try 
	{
		$wordProc = Get-Process 'WinWord' -ErrorAction SilentlyContinue
	}
	catch
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #1"
		Return ## not running
	}

	If( !$wordproc )
	{
		Write-Debug "***Exit Stop-WinWord: no WinWord tasks are running #2"
		Return ## WinWord is not running in ANY session
	}
	
	$wordrunning = $wordProc |? { $_.SessionId -eq $SessionID }
	If( !$wordrunning )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning eq null"
		Return ## not running in the current session
	}
	If( $wordrunning -is [Array] )
	{
		Write-Debug "***Exit Stop-WinWord: wordRunning is an array, elements=$($wordrunning.Count)"
		throw "Multiple Word processes are running in session $SessionID"
	}

	## it is possible for the below to throw a fault if Winword stops before it is executed.
	Stop-Process -Id $wordrunning.Id -ErrorAction SilentlyContinue
	Write-Debug "***Exit Stop-WinWord: sent Stop-Process to $($wordrunning.Id)"
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	# If Word is running - then stop it
	Stop-WinWord
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
#update 5-May-2016 by Michael B. Smith
{
	Param(
		[int] $style       = 0, 
		[int] $tabs        = 0, 
		[string] $name     = '', 
		[string] $value    = '', 
		[string] $fontName = $null,
		[int] $fontSize    = 0,
		[bool] $italics    = $false,
		[bool] $boldface   = $false,
		[Switch] $nonewline
	)
	
	#Build output style
	[string]$output = ''
	Switch ($style)
	{
		0 {$Script:Selection.Style = $myHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $myHash.Word_Heading1}
		2 {$Script:Selection.Style = $myHash.Word_Heading2}
		3 {$Script:Selection.Style = $myHash.Word_Heading3}
		4 {$Script:Selection.Style = $myHash.Word_Heading4}
		Default {$Script:Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t" 
		$tabs-- 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If( !$nonewline )
	{
		$Script:Selection.TypeParagraph()
	}
}
Function _SetDocumentProperty 
{
	#jeff hicks
	Param(
		[object] $Properties,
		[string] $Name,
		[string] $Value
	)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function AbortScript
{
	$Word.Quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject( $Word ) | Out-Null
	If( Get-Variable -Name Word -Scope Global )
	{
		Remove-Variable -Name word -Scope Global
	}
	[GC]::Collect() 
	[GC]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This Function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this Function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>
Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines=$false,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = '-231'
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting)
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This Function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>
Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' 
			{
				ForEach($Cell in $Collection) 
				{
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This Function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This Function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this Function is called by the AddWordTable Function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>
Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this Functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Add DateTime       : $($AddDateTime)"
	Write-Verbose "$(Get-Date): All                : $($All)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name       : $($Script:CoName)"
	}
	Write-Verbose "$(Get-Date): Computers          : $($Computers)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Cover Page         : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): Dev                : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile       : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Filename1          : $($Script:FileName1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2          : $($Script:FileName2)"
	}
	Write-Verbose "$(Get-Date): Folder             : $($Folder)"
	Write-Verbose "$(Get-Date): From               : $($From)"
	Write-Verbose "$(Get-Date): Groups             : $($Groups)"
	Write-Verbose "$(Get-Date): Log                : $($Log)"
	Write-Verbose "$(Get-Date): Mgmt               : $($Mgmt)"
	Write-Verbose "$(Get-Date): Organisational Unit: $($OrganisationalUnit)"
	Write-Verbose "$(Get-Date): Save As PDF        : $($PDF)"
	Write-Verbose "$(Get-Date): Save As WORD       : $($MSWORD)"
	Write-Verbose "$(Get-Date): Script Info        : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Sites              : $($Sites)"
	Write-Verbose "$(Get-Date): Smtp Port          : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server        : $($SmtpServer)"
	Write-Verbose "$(Get-Date): To                 : $($To)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name          : $($UserName)"
	}
	Write-Verbose "$(Get-Date): Users              : $($Users)"
	Write-Verbose "$(Get-Date): Visible            : $($Visible)"
	Write-Verbose "$(Get-Date): Use SSL            : $($UseSSL)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected        : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version       : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture          : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture        : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language      : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version       : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start       : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp
{
	Param(
		[object] $object,
		[string] $topLevel,
		[string] $secondLevel 
	)

	#Function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If( ( Get-Member -Name $topLevel -InputObject $object ) )
		{
			If( ( Get-Member -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Script:Word -eq $Null)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | ForEach-Object {
		If($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($BuildingBlocks -ne $Null)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($part -ne $Null)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Script:Doc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Script:Selection -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2015 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($toc -eq $Null)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param(
		[string] $AbstractTitle, 
		[string] $SubjectTitle
	)

	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running Word 2010 and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running Word 2013 and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					Stop-WinWord
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If( Test-Path variable:global:Word )
	{
		Remove-Variable -Name Word -Scope Global 4>$Null
	}
	If( Get-Variable -Name Word -Scope Script -ErrorAction SilentlyContinue )
	{
		Remove-Variable -Name Word -Scope Script 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	Stop-WinWord
}

Function SetFileName1andFileName2
{
	Param(
		[string] $OutputFileName
	)

	$pwdPath = $Folder
	If($pwdPath -eq "")
	{
		$pwdpath = $pwd.Path
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string] $Script:FileName1 = Join-Path $pwdPath $OutputFileName
		If($PDF)
		{
			[string] $Script:FileName2 = Join-Path $pwdPath $OutputFileName
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string] $Script:FileName1 = ( Join-Path $pwdPath $OutputFileName ) + '.docx'
			If($PDF)
			{
				[string] $Script:FileName2 = ( Join-Path $pwdPath $OutputFileName ) + '.pdf'
			}
		}

		SetupWord
	}
}

#region email Function
Function SendEmail
{
	Param(
		[string]$Attachments
	)

	Write-Verbose "$(Get-Date): Prepare to email"
	
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()

	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

#Script begins

$script:startTime = Get-Date

#The Function SetFileName1andFileName2 needs your script output filename
SetFileName1andFileName2 "ADHealthCheck"

#change title for your report
[string]$Script:Title = "Active Directory Health Check"

###REPLACE AFTER THIS SECTION WITH YOUR SCRIPT###

Function Split-IntoGroups 
{
    # Written by 'The Masked Avenger with the Cheetos'
    [CmdletBinding()]
    param (
        [parameter(mandatory=$true,position=0,valuefrompipeline=$true)][Object[]]$InputObject,
        [parameter(mandatory=$false,position=1)][ValidateRange(1, ([int]::MaxValue))][int]$Number=10000
    )

    begin 
	{
        $currentGroup = New-Object System.Collections.ArrayList($Number)
    } 
	process 
	{
        ForEach($object in $InputObject) {
            $index = $currentGroup.Add($object)
            If($index -ge $Number - 1) {
                ,$currentGroup.ToArray()
                $currentGroup.Clear()
            }
        }
    } 
	end 
	{
        If($currentGroup.Count -gt 0) {
            ,$currentGroup.ToArray()
        }
    }
}

Function Generate-CheckListResults 
{
    [CmdletBinding()]
    Param(
        [Parameter()]
		$Name,
		
        [Parameter()]
		$Count
    )

    $Object = New-Object -TypeName PSObject
    $Object | Add-Member -MemberType NoteProperty -Name 'Check' -value $Name
    $Object | Add-Member -MemberType NoteProperty -Name 'Results' -value $Count
    $Object
}

$global:someCallers = 0

Function Write-ToCSV 
{
	[CmdletBinding()]
	Param(
		[Parameter( Mandatory = $true,  Position = 0, ValuefromPipeline = $true )]
		$Content,
		
		[Parameter( Mandatory = $true,  Position = 1 )]
		[string] $Name,
		
		[Parameter( Mandatory = $false, Position = 2 )]
		[string] $Path = $Script:ThisScriptPath
	)

	$global:someCallers++
	If( $null -eq $Content )
	{
		Write-Debug "***Write-ToCSV: Content is empty, for call count $($global:someCallers)"
		Return
	}

    ## This code makes some assumptions (which were true at the time that
    ## the code was written):
    ## 1. All entries in $Content are the same type of PSObject.
    ## 2. Each entry in $Content is a PSObject.
    ## 3. $Content contains at least one entry.
    ## 4. PowerShell version 3 or higher.
    ## 5. Each PSObject property value is represented in the PSObject
    ##    by a string or integer.
    ## 6. No property values contain a double quote ('"').
    ## MBS - 3-May-16

    $sample = $null
	$count  = 0
    If( $Content -is [Array] )
    {
        $sample = $Content[ 0 ]
		$count  = $Content.Count
    }
    Else 
    {
        $sample = $Content
		$count  = 1
    }
 
	Write-Debug "***Write-ToCSV: content count $count, call count $($global:someCallers), content type $($content.GetType().Fullname)"
 
    $output  = @()
    $headers = ''
    
    $properties = $sample.PSObject.Properties
    ForEach( $property in $properties )
    {
        $headers += '"' + $property.Name + '"' + ','
    }

    $output += $headers.SubString( 0, $headers.Length - 1 )
    
    ForEach( $item in $Content )
    {
        $properties = $item.PSObject.Properties
        $line = ''
        ForEach( $property in $properties )
        {
            $line += '"' + "$($property.Value)" + '"' + ','
        }
        $output += $line.SubString( 0, $line.Length - 1 )
    }
    
    ## $filename = Join-Path "." ( $i.ToString() + '.csv' )
	$filename = Join-Path $Path ( $Name + '.csv' )
    $output | Out-File $filename -Force -Encoding ascii 4>$Null
} 

Function Write-ToWord 
{
    [CmdletBinding()]
    Param(
        [Parameter( Mandatory = $true, Position = 0)]
		$TableContent,
		
        [Parameter( Mandatory = $true, Position = 1)]
		[string]$Name
    )

    Write-Debug "$(Get-Date):      Writing '$Name' to Word"
    WriteWordLine -Style 3 -Tabs 0 -Name $Name
    FindWordDocumentEnd
    $TableContent | Split-IntoGroups | ForEach {
        AddWordTable -CustomObject ($TableContent) | Out-Null
        FindWordDocumentEnd
        WriteWordLine -Style 0 -Tabs 0 -Name ''
    }
    WriteWordLine -Style 0 -Tabs 0 -Name ''
}

Function ConvertTo-FQDN
{
	Param (
		[Parameter( Mandatory = $true )]
		[string] $DomainFQDN
	)

	$result = "DC=" + $DomainFQDN.Replace( ".", ",DC=" )
	Write-Debug "***ConvertTo-FQDN DomainFQDN='$DomainFQDN', result='$result'"
	Return $result
}

Function Get-Domains
{
	( [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest() ).Domains
}

Function Get-ADDomains
{
	$Domains = Get-Domains
	ForEach($Domain in $Domains) 
	{
		$DomainName = $Domain.Name
		$DomainFQDN = ConvertTo-FQDN $DomainName
		
		$ADObject   = [ADSI]"LDAP://$DomainName"
		$sidObject = New-Object System.Security.Principal.SecurityIdentifier( $ADObject.objectSid[ 0 ], 0 )

		Write-Debug "***Get-AdDomains DomName='$DomainName', sidObject='$($sidObject.Value)', name='$DomainFQDN'"

		$Object = New-Object -TypeName PSObject
		$Object | Add-Member -MemberType NoteProperty -Name 'Name'      -Value $DomainFQDN
		$Object | Add-Member -MemberType NoteProperty -Name 'FQDN'      -Value $DomainName
		$Object | Add-Member -MemberType NoteProperty -Name 'ObjectSID' -Value $sidObject.Value
		$Object
	}
}

Function Get-PrivilegedGroupsMemberCount 
{
	Param (
		[Parameter( Mandatory = $true, ValueFromPipeline = $true )]
		$Domains
	)

	## Jeff W. said this was original code, but until I got ahold of it and
	## rewrote it, it looked only slightly changed from:
	## https://gallery.technet.microsoft.com/scriptcenter/List-Membership-In-bff89703
	## So I give them both credit. :-)
	
	## the $Domains param is the output from Get-AdDomains above
	ForEach( $Domain in $Domains ) 
	{
		$DomainSIDValue = $Domain.ObjectSID
		$DomainName     = $Domain.Name
		$DomainFQDN     = $Domain.FQDN

		Write-Debug "***Get-PrivilegedGroupsMemberCount: domainName='$domainName', domainSid='$domainSidValue'"

		## Carefully chosen from a more complete list at:
		## https://support.microsoft.com/en-us/kb/243330
		## Administrator (not a group, just FYI)    - $DomainSidValue-500
		## Domain Admins                            - $DomainSidValue-512
		## Schema Admins                            - $DomainSidValue-518
		## Enterprise Admins                        - $DomainSidValue-519
		## Group Policy Creator Owners              - $DomainSidValue-520
		## BUILTIN\Administrators                   - S-1-5-32-544
		## BUILTIN\Account Operators                - S-1-5-32-548
		## BUILTIN\Server Operators                 - S-1-5-32-549
		## BUILTIN\Print Operators                  - S-1-5-32-550
		## BUILTIN\Backup Operators                 - S-1-5-32-551
		## BUILTIN\Replicators                      - S-1-5-32-552
		## BUILTIN\Network Configuration Operations - S-1-5-32-556
		## BUILTIN\Incoming Forest Trust Builders   - S-1-5-32-557
		## BUILTIN\Event Log Readers                - S-1-5-32-573
		## BUILTIN\Hyper-V Administrators           - S-1-5-32-578
		## BUILTIN\Remote Management Users          - S-1-5-32-580
		
		## FIXME - we report on all these groups for every domain, however
		## some of them are forest wide (thus the membership will be reported
		## in every domain) and some of the groups only exist in the
		## forest root.
		$PrivilegedGroups = "$DomainSidValue-512", "$DomainSidValue-518",
		                    "$DomainSidValue-519", "$DomainSidValue-520",
							"S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549",
							"S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552",
							"S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573",
							"S-1-5-32-578", "S-1-5-32-580"

		ForEach( $PrivilegedGroup in $PrivilegedGroups ) 
		{
			$source = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
			$source.SearchScope = 'Subtree'
			$source.PageSize    = 1000
			$source.Filter      = "(objectSID=$PrivilegedGroup)"
			
			Write-Debug "***Get-PrivilegedGroupsMemberCount: LDAP://$DomainName, (objectSid=$PrivilegedGroup)"
			
			$Groups = $source.FindAll()
			ForEach( $Group in $Groups )
			{
				$DistinguishedName = $Group.Properties.Item( 'distinguishedName' )
				$groupName         = $Group.Properties.Item( 'Name' )

				Write-Debug "***Get-PrivilegedGroupsMemberCount: searching group '$groupName'"

				$Source.Filter = "(memberOf:1.2.840.113556.1.4.1941:=$DistinguishedName)"
				$Users = $null
				## CHECK: I don't think a try/catch is necessary here - MBS
				try 
				{
					$Users = $Source.FindAll()
				} 
				catch 
				{
					# nothing
				}
				If( $null -eq $users )
				{
					## Obsolete: F-I-X-M-E: we should probably Return a PSObject with a count of zero
					## Write-ToCSV and Write-ToWord understand empty Return results.

					Write-Debug "***Get-PrivilegedGroupsMemberCount: no members found in $groupName"
				}
				Else 
				{
					Function GetProperValue
					{
						Param(
							[Object] $object
						)

						If( $object -is [System.DirectoryServices.SearchResultCollection] )
						{
							Return $object.Count
						}
						If( $object -is [System.DirectoryServices.SearchResult] )
						{
							Return 1
						}
						If( $object -is [Array] )
						{
							Return $object.Count
						}
						If( $null -eq $object )
						{
							Return 0
						}

						Return 1
					}

					[int]$script:MemberCount = GetProperValue $Users

					Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' user count before first filter $MemberCount"

					$Object = New-Object -TypeName PSObject
					$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
					$Object | Add-Member -MemberType NoteProperty -Name 'Group'  -Value $groupName

					$Members = $Users | Where-Object { $_.Properties.Item( 'objectCategory' ).Item( 0 ) -like 'cn=person*' }
					$script:MemberCount = GetProperValue $Members

					Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' user count after first filter $MemberCount"

					Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' has $MemberCount members"

					$Object | Add-Member -MemberType NoteProperty -Name 'Members' -Value $MemberCount
					$Object
				}
			}
		}
	}
}

Function Get-AllADDomainControllers 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	
	$adsiSearcher        = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$adsiSearcher.Filter = '(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))'
	$Servers             = $adsiSearcher.FindAll() 
	
	ForEach( $Server in $Servers ) 
	{
		$dcName = $Server.Properties.item( 'Name' )

		Write-Debug "***Get-AllAdDomainControllers DomainName='$DomainName', DomainFQDN='$($DomainFQDN)', DCname='$dcName'"

		$Object = New-Object -TypeName PSObject
		$Object | Add-Member -MemberType NoteProperty -Name 'Domain'      -Value $DomainFQDN
		$Object | Add-Member -MemberType NoteProperty -Name 'Name'        -Value $dcName
		$Object | Add-Member -MemberType NoteProperty -Name 'LastContact' -Value $Server.Properties.Item( 'whenchanged' )
		$Object
	}
}

Function Get-AllADMemberServers 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN

	Write-Debug "***Enter: Get-AllAdMemberServers DomainName='$domainName'"

	$adsiSearcher        = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$adsiSearcher.Filter = '(&(objectCategory=computer)(operatingSystem=*server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8192)))"'
	$Servers             = $adsiSearcher.FindAll()
	
	If( $null -eq $servers )
	{
		Write-Debug '***Get-AllAdMemberServers: no member servers were found'
		Return
	}

	ForEach( $Server in $Servers ) 
	{
		$serverName = $Server.Properties.Item( 'Name' )

		Write-Debug "***Get-AllAdMemberServers DomainName='$DomainName', DomainFQDN='$DomainFQDN', serverName='$serverName'"

		$Object = New-Object -TypeName PSObject
		$Object | Add-Member -MemberType NoteProperty -Name 'Domain'       -Value $DomainFQDN
		$Object | Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $serverName
		$Object
	}
}

Function Get-AllADMemberServerObjects 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Parametersetname = 'PasswordNeverExpires' )]
		[Switch]$PasswordNeverExpires,

		[Parameter( Mandatory = $true, Parametersetname = 'PasswordExpiration' )]
		[int]$PasswordExpiration,

		[Parameter( Mandatory = $true, Parametersetname = 'AccountNeverExpires' )]
		[Switch]$AccountNeverExpires,

		[Parameter( Mandatory = $true, Parametersetname = 'Disabled' )]
		[Switch]$Disabled,

		[Parameter( Mandatory = $true, Position = 1, ValueFromPipeline = $true, Parametersetname = 'PasswordNeverExpires' )]
		[Parameter( Mandatory = $true, Position = 1, ValueFromPipeline = $true, Parametersetname = 'PasswordExpiration' )]
		[Parameter( Mandatory = $true, Position = 1, ValueFromPipeline = $true, Parametersetname = 'AccountNeverExpires' )]
		[Parameter( Mandatory = $true, Position = 1, ValueFromPipeline = $true, Parametersetname = 'Disabled' )]
		$Domain
	)

	$DomainName    = $Domain.Name
	$DomainFQDN    = $Domain.FQDN
	$localParamset = $PSCmdlet.ParameterSetName

	Write-Debug "***Enter Get-AllADMemberServerObjects, DomainName='$DomainName', ParamSet='$localParamset'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	
	Switch ( $localParamset ) 
	{
		'PasswordNeverExpires'
		{
			$source.Filter = "(&(objectCategory=computer)(operatingSystem=*server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(userAccountControl:1.2.840.113556.1.4.803:=65536))"
		}
		'PasswordExpiration'
		{
			$source.Filter = "(&(objectCategory=computer)(operatingSystem=*server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8192)))"
		}
		'AccountNeverExpires' 
		{
			$source.Filter = "(&(objectCategory=computer)(operatingSystem=*server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(|(accountExpires=0)(accountExpires=9223372036854775807)))"
		}
		'Disabled'
		{
			#$source.Filter = "(&(objectCategory=computer)(operatingSystem=*server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8194)))"
			$source.Filter = "(&(&(objectCategory=computer)(objectClass=computer)(operatingSystem=*server*)(useraccountcontrol:1.2.840.113556.1.4.803:=2)))"
		}
	}
	
	If( $localParamset -eq 'PasswordExpiration' ) 
	{
		try 
		{
			$source.FindAll() | ForEach-Object {
				$fileTime = $null
				$passLast = $_.Properties[ 'PwdLastSet' ].Item( 0 )
				If( $null -ne $passLast )
				{
					$fileTime = [DateTime]::FromFileTime( $passLast )
				}
				
				If( $null -ne $passLast -and
				    $fileTime -lt ( [DateTime]::Now ).AddMonths( -$PasswordExpiration ) )
				{
					$serverName = $_.Properties.Item( 'Name' )

					Write-Debug "***Get-AllADMemberServerObjects, paramset='$localParamset', found server='$serverName'"

					$Object = New-Object -TypeName PSObject
					$Object | Add-Member -MemberType NoteProperty -Name 'Domain'          -Value $DomainFQDN
					$Object | Add-Member -MemberType NoteProperty -Name 'Name'            -Value $serverName
					$Object | Add-Member -MemberType NoteProperty -Name 'PasswordLastSet' -Value $fileTime
					$Object
				}
			}     
		}
		catch
		{
		}
	} 
	Else 
	{
		try 
		{
			$source.FindAll() | ForEach-Object {
				$serverName = $_.Properties.Item( 'Name' )

				Write-Debug "***Get-AllADMemberServerObjects, paramset='$localParamset', found server='$serverName'"

				$Object = New-Object -TypeName PSObject
				$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
				$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value $serverName
				$Object
			}
		} 
		catch 
		{
		}
	}
}

Function Get-ADUserObjects 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Parametersetname = 'PasswordNeverExpires')]
		[Switch]$PasswordNeverExpires,

		[Parameter( Mandatory = $true, Parametersetname = 'PasswordNotRequired')]
		[Switch]$PasswordNotRequired,

		[Parameter( Mandatory = $true, Parametersetname = 'PasswordChangeAtNextLogon')]
		[Switch]$PasswordChangeAtNextLogon,

		[Parameter( Mandatory = $true, Parametersetname = 'PasswordExpiration')]
		[int]$PasswordExpiration,

		[Parameter( Mandatory = $true, Parametersetname = 'NotRequireKerbereosAuthentication')]
		[Switch]$NotRequireKerbereosAuthentication,

		[Parameter( Mandatory = $true, Parametersetname = 'AccountNoExpire')]
		[Switch]$AccountNoExpire,

		[Parameter( Mandatory = $true, Parametersetname = 'Disabled')]
		[Switch]$Disabled,

		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'PasswordNeverExpires' )]
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'PasswordNotRequired' )]
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'PasswordChangeAtNextLogon' )]
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'PasswordExpiration' )]
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'NotRequireKerbereosAuthentication' )]
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'AccountNoExpire' )]
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true, Parametersetname = 'Disabled' )]
		$Domain
	)

	## this doesn't know how to process passwordSettingsObjects (fine-grained passwords) -- FIXME

	$DomainName    = $Domain.Name
	$DomainFQDN    = $Domain.FQDN
	$localParamset = $PSCmdlet.ParameterSetName

	Write-Debug "***Enter Get-ADUserObjects: domain='$DomainName', paramset='$localParamset'"

	$source             = New-Object System.Directoryservices.Directorysearcher( "LDAP://$DomainName" )
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	
	Switch ( $localParamset )
	{
		'PasswordNeverExpires'
		{
			$source.filter = "(&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=65536))"
		}
		'PasswordNotRequired' 
		{
			$source.filter = "(&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=32))"
		}
		'PasswordChangeAtNextLogon' 
		{
			$source.filter = "(&(sAMAccountType=805306368)(pwdLastSet=0))"
		}
		'PasswordExpiration'
		{
			$source.filter = "(&(sAMAccountType=805306368)(pwdLastSet>=0))"
		}
		'NotRequireKerbereosAuthentication' 
		{
			$source.filter = "(&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=4194304))"
		}
		'AccountNoExpire'
		{
			$source.filter = "(&(sAMAccountType=805306368)(|(accountExpires=0)(accountExpires=9223372036854775807)))"
		}
		'Disabled' 
		{
			$source.filter = "(&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=2))"
		}
	}

	If( $localParamset -eq 'PasswordExpiration' ) 
	{
		try 
		{
			$source.FindAll() | ForEach-Object {
				$fileTime = $null
				$passLast = $_.Properties[ 'PwdLastSet' ].Item( 0 )
				If( $null -ne $passLast )
				{
					$fileTime = [DateTime]::FromFileTime( $passLast )
				}
				
				If( $null -ne $passLast -and
				    $fileTime -lt ( [DateTime]::Now ).AddMonths( -$PasswordExpiration ) )
				{
					$userName   = $_.Properties.Item( 'Name' )

					Write-Debug "***Get-ADUserObjects: domain='$DomainFQDN', paramset='$localParamset', username='$userName'"

					$Object = New-Object -TypeName PSObject
					$Object | Add-Member -MemberType NoteProperty -Name 'Domain'          -Value $DomainFQDN
					$Object | Add-Member -MemberType NoteProperty -Name 'Name'            -Value $userName
					$Object | Add-Member -MemberType NoteProperty -Name 'PasswordLastSet' -Value $fileTime
					$Object
				}
			}
		} 
		catch 
		{
		}
	}
	Else 
	{
		try 
		{
			$source.FindAll() | ForEach-Object {
				$userName = $_.Properties.Item( 'Name' )

				Write-Debug "***Get-ADUserObjects: domain='$DomainFQDN', paramset='$localParamset', username='$userName'"

				$Object = New-Object -TypeName PSObject
				$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
				$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value $userName
				$Object
			}
		} 
		catch 
		{
		}
	}
}

Function Get-OUGPInheritanceBlocked 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	
	Write-Debug "***Enter: Get-OUGPInheritanceBlocked, DomainName '$DomainName'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	$source.filter      = '(&(objectclass=OrganizationalUnit)(gpoptions=1))'
	try 
	{
		$source.FindAll() | ForEach-Object {
			$ouName = $_.Properties.Item( 'Name' )

			Write-Debug "***Get-OuGpInheritanceBlocked: Inheritance blocked on OU '$ouName' in domain '$DomainName'"

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value $ouName 
			$Object
		}
	} 
	catch 
	{
	}
}

Function Get-ADSites 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	$searchRoot = "LDAP://CN=Sites,CN=Configuration,$DomainName"

	Write-Debug "***Enter: Get-AdSites, DomainName='$($DomainName)', SearchRoot='$searchRoot'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher
	$source.SearchScope = 'Subtree'
	$source.SearchRoot  = $searchRoot
	$source.PageSize    = 1000
	$source.Filter      = '(objectclass=site)'
	
	try 
	{
		$source.FindAll() | ForEach-Object {
			$siteName = $_.Properties.Item( 'Name' )
			$desc     = $_.Properties.Item( 'Description' )

			If( [String]::IsNullOrEmpty( $desc ) )
			{
				$desc = ' '
			}
			
			Write-Debug "***Get-AdSites: domainFQDN='$DomainFQDN', sitename='$sitename', desc='$desc'"

			$subnets = @()
			$siteBL  = $_.Properties.Item( 'siteObjectBL' )
			ForEach( $item in $siteBL )
			{
				$temp = $item.SubString( 0, $item.IndexOf( ',' ) )  ## up to first ","
				$temp = $temp.SubString( 3 )                        ## drop CN=

				Write-Debug "***Get-AdSites: sitename='$sitename', subnet='$temp'"

				$subnets += $temp
			}
			If( $subnets.Count -eq 0 )
			{
				$subnets = $null
			}

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain'      -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Site'        -Value $siteName
			$Object | Add-Member -MemberType NoteProperty -Name 'Description' -Value $desc
			$Object | Add-Member -MemberType NoteProperty -Name 'Subnets'     -Value $subnets
			$Object
		}
	} 
	catch 
	{
	}
}

Function Get-ADSiteServer 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		$Domain,

		[Parameter( Mandatory = $true )]
		$Site
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	$searchRoot = "LDAP://CN=Servers,CN=$Site,CN=Sites,CN=Configuration,$DomainName"

	Write-Debug "***Enter: Get-AdSiteServer DomainName='$domainName', DomainFQDN='$domainFQDN', searchRoot='$searchRoot'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher
	$source.SearchRoot  = $searchRoot 
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	$source.Filter      = '(objectclass=server)'
	
	try 
	{
		$SiteServers = $source.FindAll()
		If( $null -ne $SiteServers ) 
		{
			ForEach( $SiteServer in $SiteServers ) 
			{
				$serverName = $SiteServer.Properties.Item( 'Name' )

				Write-Debug "***Get-AdSiteServer: serverName='$serverName' found in site '$site' in domain '$domainFQDN'"

				$Object = New-Object -TypeName PSObject
				$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
				$Object | Add-Member -MemberType NoteProperty -Name 'Site'   -Value $Site
				$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value $serverName
				$Object
			}
		} 
		Else 
		{
			Write-Debug "***Get-AdSiteServer: No server found in site '$site' in domain '$domainFQDN'"

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Site'   -Value $Site
			$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value ' '
			$Object            
		}
	} 
	catch 
	{
	}
}

Function Get-ADSiteConnection 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain,

		[Parameter( Mandatory = $true )]
		$Site
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	$searchRoot = "LDAP://CN=$Site,CN=Sites,CN=Configuration,$DomainName"

	Write-Debug "***Enter: Get-ADSiteConnection DomainName='$DomainName', DomainFQDN='$DomainFQDN', searchRoot='$searchRoot'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher
	$source.SearchRoot  = $searchRoot 
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	$source.Filter      = '(objectclass=nTDSConnection)'
	
	try 
	{
		$SiteConnections = $source.FindAll()
		If( $null -ne $SiteConnections ) 
		{
			ForEach( $SiteConnection in $SiteConnections ) 
			{
				$connectName   = $SiteConnection.Properties.Item( 'Name' )
				$connectServer = $SiteConnection.Properties.Item( 'FromServer' )

				Write-Debug "***Get-ADSiteConnection DomainFQDN='$DomainFQDN', site='$Site', connectionName='$connectName'"

				$Object = New-Object -TypeName PSObject
				$Object | Add-Member -MemberType NoteProperty -Name 'Domain'     -Value $DomainFQDN
				$Object | Add-Member -MemberType NoteProperty -Name 'Site'       -Value $Site
				$Object | Add-Member -MemberType NoteProperty -Name 'Name'       -Value $connectName
				$Object | Add-Member -MemberType NoteProperty -Name 'FromServer' -Value $($connectServer -split ',' -replace 'CN=','')[3]
				$Object
			}
		} 
		Else 
		{
			Write-Debug "***Get-ADSiteConnection DomainFQDN='$DomainFQDN', site='$Site', no connections"

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain'     -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Site'       -Value $Site
			$Object | Add-Member -MemberType NoteProperty -Name 'Name'       -Value ' '
			$Object | Add-Member -MemberType NoteProperty -Name 'FromServer' -Value ' '
			$Object        
		}
	} 
	catch 
	{
	}
}

Function Get-ADSiteLink 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	$searchRoot = "LDAP://CN=Sites,CN=Configuration,$DomainName"

	Write-Debug "***Enter: Get-AdSiteLink DomainName='$DomainName', DomainFQDN='$DomainFQDN', searchRoot='$searchRoot'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher
	$source.SearchRoot  = $searchRoot
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	$source.Filter      = '(objectclass=sitelink)'
	
	try 
	{
		$SiteLinks = $source.FindAll()
		ForEach( $SiteLink in $SiteLinks ) 
		{
			$siteLinkName = $SiteLink.Properties.Item( 'Name' )
			$siteLinkDesc = $SiteLink.Properties.Item( 'Description' )
			$siteLinkRepl = $SiteLink.Properties.Item( 'replinterval' )
			$siteLinkSite = $SiteLink.Properties.Item( 'Sitelist' )
			$siteLinkCt   = 0

			If( $siteLinkSite )
			{
				$siteLinkCt = $siteLinkSite.Count
			}

			$sites = @()
			ForEach( $item in $siteLinkSite )
			{
				$temp  = $item.SubString( 0, $item.IndexOf( ',' ) )
				$temp  = $temp.SubString( 3 )
				$sites += $temp
			}
			If( $sites.Count -eq 0 )
			{
				$sites      = $null
				$siteLinkCt = 0
			}

			Write-Debug "***Get-AdSiteLink: Name='$siteLinkName', Desc='$siteLinkDesc', Repl='$siteLinkRepl', Count='$siteLinkCt'"

			If( [String]::IsNullOrEmpty( $siteLinkDesc ) )
			{
				$siteLinkDesc = ' '
			}

			If( $null -ne $sites ) 
			{
				ForEach( $Site in $Sites ) 
				{
					Write-Debug "***Get-AdSiteLink: siteLinkName='$siteLinkName', sitename='$site'"

					$Object = New-Object -TypeName PSObject
					$Object | Add-Member -MemberType NoteProperty -Name 'Domain'               -Value $DomainFQDN
					$Object | Add-Member -MemberType NoteProperty -Name 'Name'                 -Value $siteLinkName
					$Object | Add-Member -MemberType NoteProperty -Name 'Description'          -Value $siteLinkDesc
					$Object | Add-Member -MemberType NoteProperty -Name 'Replication Interval' -Value $siteLinkRepl
					$Object | Add-Member -MemberType NoteProperty -Name 'Site'                 -Value $site
					$Object | Add-Member -MemberType NoteProperty -Name 'Site Count'           -Value $siteLinkCt
					$Object
				}
			} 
			Else 
			{
				Write-Debug "***Get-AdSiteLink: siteLinkName='$siteLinkName', siteName='<empty>'"

				$Object = New-Object -TypeName PSObject
				$Object | Add-Member -MemberType NoteProperty -Name 'Domain'               -Value $DomainFQDN
				$Object | Add-Member -MemberType NoteProperty -Name 'Name'                 -Value $siteLinkName
				$Object | Add-Member -MemberType NoteProperty -Name 'Description'          -Value $siteLinkDesc
				$Object | Add-Member -MemberType NoteProperty -Name 'Replication Interval' -Value $siteLinkRepl
				$Object | Add-Member -MemberType NoteProperty -Name 'Site'                 -Value ' '
				$Object | Add-Member -MemberType NoteProperty -Name 'Site Count'           -Value '0'
				$Object
			}
		}
	} 
	catch 
	{
	}
}

Function Get-ADSiteSubnet 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN
	$searchRoot = "LDAP://CN=Subnets,CN=Sites,CN=Configuration,$DomainName"

	Write-Debug "***Enter Get-AdSiteSubnet DomainName='$DomainName', DomainFQDN='$DomainFQDN', searchRoot='$searchRoot'"

	$source             = New-Object System.DirectoryServices.DirectorySearcher
	$source.SearchRoot  = $searchRoot
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	$source.Filter      = '(objectclass=subnet)'
	
	try 
	{
		$source.FindAll() | ForEach-Object {
			$subnetSite = ($_.Properties.Item( 'SiteObject' ) -split ',' -replace 'CN=','')[0]
			$subnetName = $_.Properties.Item( 'Name' )
			$subnetDesc = $_.Properties.Item( 'Description' )

			Write-Debug "***Get-AdSiteSubnet: site='$subnetSite', name='$subnetName', desc='$subnetDesc'"

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain'      -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Site'        -Value $subnetSite
			$Object | Add-Member -MemberType NoteProperty -Name 'Name'        -Value $subnetName
			$Object | Add-Member -MemberType NoteProperty -Name 'Description' -Value $subnetDesc
			$Object
		}
	} 
	catch 
	{
	}
}

Function Get-ADEmptyGroups 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	## $exclude includes (punny, aren't I?) the list of groups commonly used as a 
	## 'Primary Group' in Active Directory. While, theoretically, ANY group can be
	## a primary group, that is quite rare. 
	$exclude = 'Domain Users', 'Domain Computers', 'Domain Controllers', 'Domain Guests'
	
	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN

	Write-Debug "***Enter Get-AdEmptyGroups DomainName='$DomainName', DomainFQDN='$DomainFQDN'"

	$source             = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$source.SearchScope = 'Subtree'
	$source.PageSize    = 1000
	$source.Filter      = '(&(objectCategory=Group)(!member=*))'

	try 
	{
		$groups = $source.FindAll()
		$groups = (($groups | ? { $exclude -notcontains $_.Properties[ 'Name' ].Item( 0 ) } ) | % { $_.Properties[ 'Name' ].Item( 0 ) }) | sort
		ForEach( $group in $groups )
		{
			Write-Debug "***Get-AdEmptyGroups: DomainFQDN='$DomainFQDN', empty groupname='$group'"

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value $group
			$Object
		}
	}
	catch 
	{
	}
}

Function Get-ADDomainLocalGroups 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN

	Write-Debug "***Enter Get-AdDomainLocalGroups DomainName='$DomainName', DomainFQDN='$DomainFQDN'"

	$search             = New-Object System.DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$search.SearchScope = 'Subtree'
	$search.PageSize    = 1000
	$search.Filter      = '(&(groupType:1.2.840.113556.1.4.803:=4)(!(groupType:1.2.840.113556.1.4.803:=1)))'
	
	try 
	{
		$search.FindAll() | ForEach-Object {
			$groupName = $_.Properties.Item( 'Name' )
			$groupDN   = $_.Properties.Item( 'Distinguishedname' )

			Write-Debug "***Get-AdDomainLocalGroups groupName='$groupName', dn='$groupDN'"

			$Object = New-Object -TypeName PSObject
			$Object | Add-Member -MemberType NoteProperty -Name 'Domain'            -Value $DomainFQDN
			$Object | Add-Member -MemberType NoteProperty -Name 'Name'              -Value $groupName
			$Object | Add-Member -MemberType NoteProperty -Name 'DistinguishedName' -Value $groupDN
			$Object
		}
	} 
	catch 
	{
	}
}

Function Get-ADUsersInDomainLocalGroups 
{
	[CmdletBinding()]
	Param (
		[Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
		$Domain
	)

	$DomainName = $Domain.Name
	$DomainFQDN = $Domain.FQDN

	Write-Debug "***Enter Get-AdUsersInDomainLocalGroups DomainName='$DomainName', DomainFQDN='$DomainFQDN'"

	$search             = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
	$search.SearchScope = 'Subtree'
	$search.PageSize    = 1000
	$search.Filter      = '(&(groupType:1.2.840.113556.1.4.803:=4)(!(groupType:1.2.840.113556.1.4.803:=1)))'
	
	try 
	{
		## $search was being used twice.
		$results = $search.FindAll() 
		$results | ForEach-Object {
			$groupName         = $_.Properties.Item( 'Name' )
			$DistinguishedName = $_.Properties.Item( 'DistinguishedName' )

			Write-Debug "***Get-AdUsersInDomainLocalGroups name='$groupName', dn='$distinguishedName'"

			$search.Filter = "(&(memberOf=$DistinguishedName)(objectclass=User))"
			$search.FindAll() | ForEach-Object {
				$userName = $_.Properties.Item( 'Name' )

				Write-Debug "***Get-AdUsersInDomainLocalGroups name='$groupName', user='$userName'" 

				$Object = New-Object -TypeName PSObject
				$Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $DomainFQDN
				$Object | Add-Member -MemberType NoteProperty -Name 'Group'  -Value $groupName
				$Object | Add-Member -MemberType NoteProperty -Name 'Name'   -Value $userName
				$Object
			}
		}
	} 
	catch 
	{
	}
}

#region process document output
Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}

	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
			Write-Verbose "$(Get-Date): "
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
			Write-Verbose "$(Get-Date): "
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}

	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}

	Write-Verbose "$(Get-Date): "
}
#endregion

#region end script
Function ProcessScriptEnd
{
	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}
	
	If($ScriptInfo)
	{
		Out-File -FilePath $Script:SIFile -InputObject "" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Add DateTime       : $($AddDateTime)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "All                : $($All)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $Script:SIFile -Append -InputObject "Company Name       : $($Script:CoName)" 4>$Null		
		}
		Out-File -FilePath $Script:SIFile -Append -InputObject "Computers          : $($computers)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $Script:SIFile -Append -InputObject "Cover Page         : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $Script:SIFile -Append -InputObject "Dev                : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $Script:SIFile -Append -InputObject "DevErrorFile      : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $Script:SIFile -Append -InputObject "Filename1          : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $Script:SIFile -Append -InputObject "Filename2          : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $Script:SIFile -Append -InputObject "Folder             : $($Folder)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "From               : $($From)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Groups             : $($groups)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Mgmt               : $($mgmt)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Organisational Unit: $($OrganisationalUnit)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Save As PDF        : $($PDF)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Save As WORD       : $($MSWORD)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Script Info        : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Sites              : $($Sites)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Smtp Port          : $($SmtpPort)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Smtp Server        : $($SmtpServer)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "To                 : $($To)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Use SSL            : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $Script:SIFile -Append -InputObject "User Name          : $($UserName)" 4>$Null
		}
		Out-File -FilePath $Script:SIFile -Append -InputObject "Users              : $($users)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Visible            : $($Visible)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "PSCulture          : $($PSCulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $Script:SIFile -Append -InputObject "Word language      : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $Script:SIFile -Append -InputObject "Word version       : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $Script:SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Script start       : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $Script:SIFile -Append -InputObject "Elapsed time       : $($Str)" 4>$Null
	}
	
	$runtime = $Null
	$Str = $Null
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region Content
$Script:MgmtPage = @()

Function Generate-TableContent
{
	[CmdletBinding()]
	Param(
		$content,
		$hashParam,
		$title
	)

	$count = 0
	If( $null -eq $content )
	{
		## do not early-Return, because the MgmtPage needs to be updated
		Write-Debug "***Generate-TableContent: empty for title='$title'"
	}
	Else
	{
		$count = 1
		If( $content -is [Array] )
		{
			$count = $content.Count
		}

		Write-Debug "***Generate-TableContent: count=$count for title='$title'"

		If( $hashParam.ContainsKey( 'CSV' ) )
		{
			Write-ToCSV -Name $title -Content $content
		}
		Write-ToWord -Name $title -TableContent $content
	}
	
	If( $hashParam.ContainsKey( 'Mgmt' ) ) 
	{
		$script:MgmtPage += Generate-CheckListResults -Name $title -Count $count
	}
}

Function IsInDomain
{
	$computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue -verbose:$False
	If( !$? -or $null -eq $computerSystem )
	{
		$computerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
		If( !$? -or $null -eq $computerSystem )
		{
			Write-Error 'IsInDomain: fatal error: cannot obtain Win32_ComputerSystem from CIM or WMI.'
			AbortScript
		}
	}
	
	Return $computerSystem.PartOfDomain
}

If( -not ( IsInDomain ) )
{
	Write-Error 'ADHealthCheck must be run from a computer that is a member of a domain.'
	AbortScript
}

FindWordDocumentEnd
$Script:Selection.InsertNewPage()
Write-Verbose "$(Get-Date): Get domains" 
$Domains = Get-ADDomains
If( $null -eq $Domains )
{
	Write-Error 'ADHealthCheck cannot obtain a list of domains in the forest.'
	AbortScript
}

$parameters = $PSBoundParameters
$paramset   = $PSCmdlet.ParameterSetName

ForEach( $Domain in $Domains ) 
{
	$DomainFQDN = $Domain.FQDN
	Write-Verbose "$(Get-Date): Domain $DomainFQDN"
	WriteWordLine -Style 1 -Tabs 0 -Name "Domain $DomainFQDN"
	FindWordDocumentEnd
	If(($parameters.ContainsKey('Sites')) -or ($paramset -eq 'All') -or ($paramset -eq 'SMTP')) 
	{
		#Sites
		$Script:Selection.InsertNewPage()
		FindWordDocumentEnd
		Write-Verbose "$(Get-Date):  Sites"
		WriteWordLine -Style 2 -Tabs 0 -Name 'Sites'
		FindWordDocumentEnd
		$TableContentTemp = Get-ADSites -Domain $Domain
		
		#Sites - Description empty
		$CheckTitle = 'Sites - Without a description'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		If($TableContentTemp -ne $null) 
		{
			$TableContent = $TableContentTemp | Where-Object {$_.Description -eq $null}
			Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

			#Sites - No subnet
			$CheckTitle = 'Sites - Without one or more subnet(s)'
			Write-Verbose "$(Get-Date):   $CheckTitle"
			$TableContent = $TableContentTemp | Where-Object {$_.Subnets -eq $null}
			Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

			#Sites - No server
			$CheckTitle = 'Sites - No server(s)'
			Write-Verbose "$(Get-Date):   $CheckTitle"
			$TableContent = $TableContentTemp | ForEach-Object { Get-ADSiteServer -Site $_.Site -Domain $Domain } | Where-Object {$_.Name -eq $null}
			Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

			#Sites - No connection
			$CheckTitle = 'Sites - Without a connection'
			Write-Verbose "$(Get-Date):   $CheckTitle"
			$TableContent = $TableContentTemp | ForEach-Object { Get-ADSiteConnection -Site $_.site -Domain $Domain } | Where-Object {$_.Name -eq $null}
			WriteWordLine -Style 3 -Tabs 0 -Name $CheckTitle
			FindWordDocumentEnd
			Generate-TableContent $TableContent $PSBoundParameters $CheckTitle
			WriteWordLine -Style 0 -Tabs 0 -Name ''
			FindWordDocumentEnd
		}

		$allSiteLinks = Get-AdSiteLink -Domain $Domain
		
		#Sites - No sitelink
		$CheckTitle = 'Sites - No sitelink(s)'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = $allSiteLinks | Where-Object {$_.'Site Count' -eq '0'}
		Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

		#Sitelinks - One site
		$CheckTitle = 'Sites - With one sitelink'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = $allSiteLinks | Where-Object {$_.'Site Count' -eq '1'}
		Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

		#Sitelinks - More than two sites
		$CheckTitle = 'SiteLinks - More than two sitelinks'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = $allSiteLinks | Where-Object {$_.'Site Count' -gt '2'}
		Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

		#Sitelinks - No description
		$CheckTitle = 'SiteLinks - Without a description'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = $allSiteLinks | Where-Object {$_.Description -eq $null}
		Generate-TableContent $TableContent $PSBoundParameters $CheckTitle

		#ADSubnets - Available but not in use
		$CheckTitle = 'Subnets in Sites - Not in use'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$AvailableSubnets = Get-ADSiteSubnet -Domain $Domain | select -ExpandProperty 'name'
		$InUseSubnets = Get-ADSites -Domain $Domain | select -ExpandProperty 'subnets'
		If(($AvailableSubnets -ne $Null) -and ($InUseSubnets -ne $null)) 
		{
			$TableContent = Compare-Object -DifferenceObject $InUseSubnets -ReferenceObject $AvailableSubnets
			Generate-TableContent $TableContent $parameters $CheckTitle
		}
	}
	If(($parameters.ContainsKey('OrganisationalUnit')) -or ($paramset -eq 'All') -or ($paramset -eq 'SMTP')) 
	{
		#OrganisationalUnit
		$Script:Selection.InsertNewPage()
		FindWordDocumentEnd
		Write-Verbose "$(Get-Date):  OU"
		WriteWordLine -Style 2 -Tabs 0 -Name 'Organisational Units'
		## FIXME - no organizational units shown
		#OU - GPO inheritance blocked
		$CheckTitle = 'OU - GPO inheritance blocked'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-OUGPInheritanceBlocked -Domain $Domain
		WriteWordLine -Style 3 -Tabs 0 -Name $CheckTitle
		FindWordDocumentEnd
		Generate-TableContent $TableContent $parameters $CheckTitle
	}
	If(($parameters.ContainsKey('Computers')) -or ($paramset -eq 'All') -or ($paramset -eq 'SMTP')) 
	{
		#Domain Controllers
		$Script:Selection.InsertNewPage()
		FindWordDocumentEnd
		Write-Verbose "$(Get-Date):  Domain Controllers"
		WriteWordLine -Style 2 -Tabs 0 -Name 'Domain Controllers'
		## FIXME - write all domain controller names? Domain? OS Version? Etc.
		FindWordDocumentEnd

		#Domain Controllers - No contact
		$CheckTitle = 'Domain Controllers - No contact in the last 3 months'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-AllADDomainControllers -Domain $Domain | Where-Object {$_.LastContact -lt (([datetime]::Now).AddMonths(-6))} | Sort-Object -Property LastContact -Descending 
		WriteWordLine -Style 3 -Tabs 0 -Name $CheckTitle
		FindWordDocumentEnd
		Generate-TableContent $TableContent $parameters $CheckTitle
		WriteWordLine -Style 0 -Tabs 0 -Name ''
		FindWordDocumentEnd

		#Member Servers
		Write-Verbose "$(Get-Date):  Member Servers"
		WriteWordLine -Style 2 -Tabs 0 -Name 'Member Servers'
		FindWordDocumentEnd

		#Member Servers - Password never expires
		$CheckTitle = 'Member Servers - Password never expires'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-AllADMemberServerObjects -Domain $Domain -PasswordNeverExpires | Sort -Property Name
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Computers - Password expired
		$CheckTitle = 'Member Servers - Password more than 6 months old'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-AllADMemberServerObjects -Domain $Domain -PasswordExpiration '6' | Sort -Property Name
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Member Servers - Account never expires
		$CheckTitle = 'Member Servers - Account never expires'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-AllADMemberServerObjects -Domain $Domain -AccountNeverExpires | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Member Servers - Account disabled
		$CheckTitle = 'Member Servers - Account disabled'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-AllADMemberServerObjects -Domain $Domain -Disabled | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle
	}

	If(($parameters.ContainsKey('Users')) -or ($paramset -eq 'All') -or ($paramset -eq 'SMTP')) 
	{
		#Users
		$Script:Selection.InsertNewPage()
		FindWordDocumentEnd
		Write-Verbose "$(Get-Date):  Users"
		WriteWordLine -Style 2 -Tabs 0 -Name 'Users'
		FindWordDocumentEnd

		#Users in Domain Local Groups
		$CheckTitle = 'Users - Direct member of a Domain Local Group'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUsersInDomainLocalGroups -Domain $Domain | Sort -Property Group, Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Password never expires
		$CheckTitle = 'Users - Password never expires'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -PasswordNeverExpires | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Password not required
		$CheckTitle = 'Users - Password not required'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -PasswordNotRequired | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Password needs to be changed at next logon
		$CheckTitle = 'Users - Change password at next logon'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -PasswordChangeAtNextLogon | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Password not changed in last 12 months
		$CheckTitle = 'Users - Password not changed in last 12 months'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -PasswordExpiration '12' | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Account without expiration date
		$CheckTitle = 'Users - Account without expiration date'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -AccountNoExpire | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Do not require kerberos preauthentication
		$CheckTitle = 'Users - Do not require kerberos preauthentication'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -NotRequireKerbereosAuthentication | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle

		#Users - Disabled
		$CheckTitle = 'Users - Disabled'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADUserObjects -Domain $Domain -Disabled | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle
	}

	If(($parameters.ContainsKey('Groups')) -or ($paramset -eq 'All') -or ($paramset -eq 'SMTP')) 
	{
		#Groups
		Write-Verbose "$(Get-Date):  Groups"
		$Script:Selection.InsertNewPage()
		FindWordDocumentEnd
		WriteWordLine -Style 2 -Tabs 0 -Name 'Groups'
		FindWordDocumentEnd
		#Privileged Groups
		Write-Verbose "$(Get-Date):   Groups - Privileged groups"
		$TableContentTemp = Get-PrivilegedGroupsMemberCount -Domains $Domain | Sort -Property Group

		#Groups - Privileged with many members
		$CheckTitle = 'Groups - Privileged - More than 5 members'
		Write-Verbose "$(Get-Date):    $CheckTitle"
		If($TableContentTemp -ne $null) 
		{
			$TableContent = $TableContentTemp | Where {$_.Members -gt '5'} | Sort -Property Group 
			Generate-TableContent $TableContent $parameters $CheckTitle
		}

		#Groups - Privileged with no members
		$CheckTitle = 'Groups - Privileged - No members'
		Write-Verbose "$(Get-Date):    $CheckTitle"
		If($TableContentTemp -ne $null) 
		{
			$TableContent = $TableContentTemp | Where {$_.Members -eq '0'} | Sort -Property Group 
			Generate-TableContent $TableContent $parameters $CheckTitle
		}

		#Groups - Empty
		$CheckTitle = 'Groups - Primary - Empty (no members)'
		Write-Verbose "$(Get-Date):   $CheckTitle"
		$TableContent = Get-ADEmptyGroups -Domain $Domain | Sort -Property Name 
		Generate-TableContent $TableContent $parameters $CheckTitle
	}
	
	$CheckTitle = 'Management'
	Write-Verbose "$(Get-Date):   $CheckTitle"
	If($parameters.ContainsKey('Mgmt')) 
	{
		If($parameters.ContainsKey('CSV')) 
		{
			Write-ToCSV -Name $CheckTitle -Content $MgmtPage         
		}
		$Script:Selection.InsertNewPage()
		FindWordDocumentEnd
		WriteWordLine -Style 2 -Tabs 0 -Name $CheckTitle
		FindWordDocumentEnd
		Write-ToWord -Name 'Management Table' -TableContent $MgmtPage
	}
}
#endregion Content

###REPLACE BEFORE THIS SECTION WITH YOUR SCRIPT###

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

###Change the two lines below for your script
$AbstractTitle = "AD Health Check Report"
$SubjectTitle = "Active Directory Health Check Report"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd

# open Word document automatically for the user here
Start $Script:FileName1 -Verb open

If($parameters.ContainsKey('Log')) 
{
	If($Script:StartLog -eq $true) 
	{
		try 
		{
			Stop-Transcript | Out-Null
			Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
		} 
		catch 
		{
			Write-Verbose "$(Get-Date): Transcript/log stop failed"
		}
	}
}

