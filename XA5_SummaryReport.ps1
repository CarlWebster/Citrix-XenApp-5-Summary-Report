#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a Summary Report of the inventory of a Citrix XenApp 5 farm using Microsoft Word.
.DESCRIPTION
	Creates a Summary Report of the inventory of a Citrix XenApp 5 farm using Microsoft Word.l.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
.EXAMPLE
	PS C:\PSScript > .\XA5_SummaryReport.ps1
	
	Runs and creates a one page report.
	
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.NOTES
	NAME: XA5_SummaryReport.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster
	LASTEDIT: November 2, 2013
#>

Set-StrictMode -Version 2

#the following values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
[int]$wdMove = 0
[int]$wdSeekMainDocument = 0
[int]$wdStory = 6
[int]$wdWord2007 = 12
[int]$wdWord2010 = 14
[int]$wdWord2013 = 15
[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

$hash = @{}

# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleHeading3 = -4
$wdStyleHeading4 = -5
$wdStyleNoSpacing = -158

$myHash = $hash

$myHash.Word_NoSpacing = $wdStyleNoSpacing
$myHash.Word_Heading1 = $wdStyleheading1
$myHash.Word_Heading2 = $wdStyleheading2
$myHash.Word_Heading3 = $wdStyleheading3
$myHash.Word_Heading4 = $wdStyleheading4

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
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

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {Write-Warning "($_)"}
		return $False
	}
	Else
	{
		Return $True
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
{
	Param([int]$style = 0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [Switch]$nonewline)
	$output=""
	#Build output style
	Switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	#build # of tabs
	While($tabs -gt 0) { 
		$output += "`t"; $tabs--; 
	}
		
	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
	
	#test for new WriteWordLine 0.
	If($nonewline){
		# Do nothing.
	} Else {
		$Selection.TypeParagraph()
	}
}

Function AbortScript
{
	$Word.quit()
	Write-Host "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	Remove-Variable -Name word -Scope Global -EA 0
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Host "$(Get-Date): Script has been aborted"
	Exit
}

#script begins

If(!(Check-NeededPSSnapins "Citrix.XenApp.Commands")){
    #We're missing Citrix Snapins that we need
    Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 5 Server? Script will now close."
    Exit
}

CheckWordPreReq

Write-Host "$(Get-Date): Getting Farm data"
$farm = Get-XAFarm -EA 0
If($?)
{
	Write-Host "$(Get-Date): Verify farm version"
	#first check to make sure this is a XenApp 5 farm
	If($Farm.ServerVersion.ToString().SubString(0,1) -ne "6")
	{
		If($Farm.ServerVersion.ToString().SubString(0,1) -eq "4")
		{
			$FarmOS = "2003"
		}
		Else
		{
			$FarmOS = "2008"
		}
		Write-Host "$(Get-Date): Farm OS is $($FarmOS)"
		#this is a XenApp 5 farm, script can proceed
		#XenApp 5 for server 2003 shows as version 4.6
		#XenApp 5 for server 2008 shows as version 5.0
	}
	Else
	{
		#this is not a XenApp 5 farm, script cannot proceed
		Write-Warning "This script is designed for XenApp 5 and should not be run on XenApp 6.x"
		Return 1
	}
	
	$FarmName = $farm.FarmName
	$filename1="$($pwd.path)\Summary Report for $($FarmName).docx"
} 
Else 
{
	Write-Warning "Farm information could not be retrieved"
	Write-Error "Farm information could not be retrieved.  Script cannot continue."
	Exit
}
 
$farm = $Null

Write-Host "$(Get-Date): Setting up Word"

# Setup word for output
Write-Host "$(Get-Date): Create Word comObject.  Ignore the next message."
$Word = New-Object -comobject "Word.Application" -EA 0

If(!$? -or $Word -eq $Null)
{
	Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
	Write-Error "The Word object could not be created.  You may need to repair your Word installation.  Script cannot continue."
	Exit
}

[int]$WordVersion = [int] $Word.Version
If($WordVersion -eq $wdWord2013)
{
	$WordProduct = "Word 2013"
}
ElseIf($WordVersion -eq $wdWord2010)
{
	$WordProduct = "Word 2010"
}
ElseIf($WordVersion -eq $wdWord2007)
{
	$WordProduct = "Word 2007"
}
Else
{
	Write-Error "You are running an untested or unsupported version of Microsoft Word.  Script will end.  Please send info on your version of Word to webster@carlwebster.com"
	AbortScript
}

Write-Host "$(Get-Date): Running Microsoft $WordProduct"

$Word.Visible = $False

Write-Host "$(Get-Date): Create empty word doc"
$Doc = $Word.Documents.Add()
If($Doc -eq $Null)
{
	Write-Host "$(Get-Date): "
	Write-Error "An empty Word document could not be created.  Script cannot continue."
	AbortScript
}

$Selection = $Word.Selection
If($Selection -eq $Null)
{
	Write-Host "$(Get-Date): "
	Write-Error "An unknown error happened selecting the entire Word document for default formatting options.  Script cannot continue."
	AbortScript
}

#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
#36 = .50"
$Word.ActiveDocument.DefaultTabStop = 36

#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
Write-Host "$(Get-Date): Disable grammar and spell checking"
$Word.Options.CheckGrammarAsYouType=$False
$Word.Options.CheckSpellingAsYouType=$False

#return focus to main document
Write-Host "$(Get-Date): Return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Host "$(Get-Date): Move to the end of the current document"
Write-Host "$(Get-Date):"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 

#process the nodes in the Delivery Services Console (XA5/2003) and the Access Management Console (XA5/2008)

$ConfigLog = $False
$farm = Get-XAFarmConfiguration -EA 0

If($?)
{
	If($Farm.ConfigLogEnabled)
	{
		$ConfigLog = $True
	}
}
$farm = $null

Write-Host "$(Get-Date): Processing Administrators"
[int]$TotalFullAdmins = 0
[int]$TotalViewAdmins = 0
[int]$TotalCustomAdmins = 0

Write-Host "$(Get-Date): `tRetrieving Administrators"
$Administrators = Get-XAAdministrator -EA 0 | Sort-Object AdministratorName

If($?)
{
	ForEach($Administrator in $Administrators)
	{
		Write-Host "$(Get-Date): `t`tProcessing administrator $($Administrator.AdministratorName)"
		Switch ($Administrator.AdministratorType)
		{
			"Unknown"  {}
			"Full"     {$TotalFullAdmins++}
			"ViewOnly" {$TotalViewAdmins++}
			"Custom"   {$TotalCustomAdmins++}
			Default    {}
		}
	}
}
Else 
{
	Write-Warning "Administrator information could not be retrieved"
}
$Administrators = $Null
Write-Host "$(Get-Date): Finished Processing Administrators"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Applications"

[int]$TotalPublishedApps = 0
[int]$TotalPublishedContent = 0
[int]$TotalPublishedDesktops = 0
[int]$TotalStreamedApps = 0

Write-Host "$(Get-Date): `tRetrieving Applications"
$Applications = Get-XAApplication -EA 0 | Sort-Object FolderPath, DisplayName

If($? -and $Applications -ne $Null)
{
	ForEach($Application in $Applications)
	{
		Write-Host "$(Get-Date): `t`tProcessing application $($Application.BrowserName)"
		
		Switch ($Application.ApplicationType)
		{
			"Unknown"                            {}
			"ServerInstalled"                    {$TotalPublishedApps++}
			"ServerDesktop"                      {$TotalPublishedDesktops++}
			"Content"                            {$TotalPublishedContent++}
			"StreamedToServer"                   {$TotalStreamedApps++}
			"StreamedToClient"                   {$TotalStreamedApps++}
			"StreamedToClientOrInstalled"        {$TotalStreamedApps++}
			"StreamedToClientOrStreamedToServer" {$TotalStreamedApps++}
			Default {}
		}
	}
}
ElseIf($Applications -eq $Null)
{
	Write-Host "$(Get-Date): There are no Applications published"
}
Else 
{
	Write-Warning "Application information could not be retrieved"
}
$Applications = $Null
Write-Host "$(Get-Date): Finished Processing Applications"
Write-Host "$(Get-Date): "

#servers
Write-Host "$(Get-Date): Processing Servers"
[int]$TotalServers = 0

Write-Host "$(Get-Date): `tRetrieving Servers"
$servers = Get-XAServer -EA 0 | Sort-Object FolderPath, ServerName

If($?)
{
	ForEach($server in $servers)
	{
		$TotalServers++
		Write-Host "$(Get-Date): `t`tProcessing server $($server.ServerName)"
	}
}
Else 
{
	Write-Warning "Server information could not be retrieved"
}
$servers = $Null
Write-Host "$(Get-Date): Finished Processing Servers"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Zones"
Write-Host "$(Get-Date): `tSetting summary variables"
[int]$TotalZones = 0

Write-Host "$(Get-Date): `tRetrieving Zone Information"
$Zones = Get-XAZone -EA 0 | Sort-Object ZoneName
If($?)
{
	ForEach($Zone in $Zones)
	{
		$TotalZones++
	}
}
Else 
{
	Write-Warning "Zone information could not be retrieved"
}
$Servers = $Null
$Zone = $Null
Write-Host "$(Get-Date): Finished Processing Zones"
Write-Host "$(Get-Date): "

#Process the nodes in the Advanced Configuration Console

Write-Host "$(Get-Date): Processing Load Evaluators"
#load evaluators
[int]$TotalLoadEvaluators = 0

Write-Host "$(Get-Date): `tRetrieving Load Evaluators"
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | Sort-Object LoadEvaluatorName

If($?)
{
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		$TotalLoadEvaluators++
	}
}
Else 
{
	Write-Warning "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $Null
Write-Host "$(Get-Date): Finished Processing Load Evaluators"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Policies"
[int]$TotalPolicies = 0

Write-Host "$(Get-Date): `tRetrieving Policies"
$Policies = Get-XAPolicy -EA 0 | Sort-Object PolicyName
If($? -and $Policies -ne $Null)
{
	ForEach($Policy in $Policies)
	{
		$TotalPolicies++
		Write-Host "$(Get-Date): `tProcessing policy $($Policy.PolicyName)"

	}
}
ElseIf($Policies -eq $Null)
{
	Write-Host "$(Get-Date): There are no Policies created"
}
Else 
{
	Write-Warning "Citrix Policy information could not be retrieved."
}
$Policies = $Null
Write-Host "$(Get-Date): Finished Processing Policies"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Print Drivers"
#printer drivers
[int]$TotalPrintDrivers = 0

Write-Host "$(Get-Date): `tRetrieving Print Drivers"
$PrinterDrivers = Get-XAPrinterDriver -EA 0 | Sort-Object DriverName

If($? -and $PrinterDrivers -ne $Null)
{
	ForEach($PrinterDriver in $PrinterDrivers)
	{
		$TotalPrintDrivers++
		Write-Host "$(Get-Date): `t`tProcessing driver $($PrinterDriver.DriverName)"
	}
}
ElseIf($PrinterDrivers -eq $Null)
{
	Write-Host "$(Get-Date): There are no Printer Drivers created"
}
Else 
{
	Write-Warning "Printer driver information could not be retrieved"
}
$PrintDrivers = $Null
Write-Host "$(Get-Date): Finished Processing Print Drivers"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Processing Printer Driver Mappings"
#printer driver mappings
[int]$TotalPrintDriverMappings = 0

Write-Host "$(Get-Date): `tRetrieving Print Driver Mappings"
$PrinterDriverMappings = Get-XAPrinterDriverMapping -EA 0 | Sort-Object ClientDriverName

If($? -and $PrinterDriverMappings -ne $Null)
{
	ForEach($PrinterDriverMapping in $PrinterDriverMappings)
	{
		$TotalPrintDriverMappings++
		Write-Host "$(Get-Date): `t`tProcessing drive $($PrinterDriverMapping.ClientDriverName)"
	}
}
ElseIf($PrinterDriverMappings -eq $Null)
{
	Write-Host "$(Get-Date): There are no Printer Driver Mappings created"
}
Else 
{
	Write-Warning "Printer driver mapping information could not be retrieved"
}
$PrintDriverMappings = $Null
Write-Host "$(Get-Date): Finished Processing Printer Driver Mappings"
Write-Host "$(Get-Date): "

[int]$TotalConfigLogItems = 0

If($ConfigLog)
{
	Write-Host "$(Get-Date): Processing the Configuration Logging Report"
	#Configuration Logging report
	#only process if $ConfigLog = $True and XA5ConfigLog.udl file exists
	#build connection string for Microsoft SQL Server
	#User ID is account that has access permission for the configuration logging database
	#Initial Catalog is the name of the Configuration Logging SQL Database
	If(Test-Path “$($pwd.path)\XA5ConfigLog.udl”)
	{
		Write-Host "$(Get-Date): `tRetrieving Configuration Logging Data"
		$ConnectionString = Get-Content “$($pwd.path)\XA5ConfigLog.udl” | select-object -last 1
		$ConfigLogReport = get-XAConfigurationLog -connectionstring $ConnectionString -EA 0

		If($? -and $ConfigLogReport)
		{
			Write-Host "$(Get-Date): `t`tProcessing $($ConfigLogReport.Count) configuration logging items"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				$TotalConfigLogItems++
			}
		} 
		Else 
		{
			Write-Warning "Configuration log report could not be retrieved"
		}
		$ConfigLogReport = $Null
	}
	Else 
	{
		$selection.InsertNewPage()
		Write-Warning "Configuration Logging is enabled but the XA5ConfigLog.udl file was not found"
	}
	Write-Host "$(Get-Date): Finished Processing the Configuration Logging Report"
}
Write-Host "$(Get-Date): "

#summary report
Write-Host "$(Get-Date): Create Summary Report"
WriteWordLine 1 0 "Summary Report for the $($farmname) Farm"
Write-Host "$(Get-Date): Add administrator summary info"
WriteWordLine 0 0 "Administrators"
WriteWordLine 0 1 "Total Full Administrators`t: " $TotalFullAdmins
WriteWordLine 0 1 "Total View Administrators`t: " $TotalViewAdmins
WriteWordLine 0 1 "Total Custom Administrators`t: " $TotalCustomAdmins
WriteWordLine 0 2 "Total Administrators`t: " ($TotalFullAdmins + $TotalViewAdmins + $TotalCustomAdmins)
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add application summary info"
WriteWordLine 0 0 "Applications"
WriteWordLine 0 1 "Total Published Applications`t: " $TotalPublishedApps
WriteWordLine 0 1 "Total Published Content`t`t: " $TotalPublishedContent
WriteWordLine 0 1 "Total Published Desktops`t: " $TotalPublishedDesktops
WriteWordLine 0 1 "Total Streamed Applications`t: " $TotalStreamedApps
WriteWordLine 0 2 "Total Applications`t: " ($TotalPublishedApps + $TotalPublishedContent + $TotalPublishedDesktops + $TotalStreamedApps)
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add server summary info"
WriteWordLine 0 0 "Servers"
WriteWordLine 0 2 "Total Servers`t`t: " $TotalServers
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add zone summary info"
WriteWordLine 0 0 "Zones"
WriteWordLine 0 2 "Total Zones`t`t: " $TotalZones
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add load evaluator summary info"
WriteWordLine 0 0 "Load Evaluators"
WriteWordLine 0 2 "Total Load Evaluators`t: " $TotalLoadEvaluators
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add policy summary info"
WriteWordLine 0 0 "Policies"
WriteWordLine 0 2 "Total Policies`t`t: " $TotalPolicies
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add print driver summary info"
WriteWordLine 0 0 "Print Drivers"
WriteWordLine 0 2 "Total Print Drivers`t: " $TotalPrintDrivers
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add print driver mapping summary info"
WriteWordLine 0 0 "Print Driver Mappingss"
WriteWordLine 0 2 "Total Prt Drvr Mappings: " $TotalPrintDriverMappings
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): Add configuration logging summary info"
WriteWordLine 0 0 "Configuration Logging"
WriteWordLine 0 2 "Total Config Log Items`t: " $TotalConfigLogItems 
WriteWordLine 0 0 ""
Write-Host "$(Get-Date): `tFinished Create Summary Page"
Write-Host "$(Get-Date): "
Write-Host "$(Get-Date): Finishing up Word document"

#end of document processing

#the $saveFormat below passes StrictMode 2
#I found this at the following two links
#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
Write-Host "$(Get-Date): Save and Close document and Shutdown Word"

$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)

Write-Host "$(Get-Date): Closing Word"
$doc.Close()
$Word.Quit()
Write-Host "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word -Scope Global -EA 0
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): Script has completed"
Write-Host "$(Get-Date): "

Write-Host "$(Get-Date): $($filename1) is ready for use"
Write-Host "$(Get-Date): "
