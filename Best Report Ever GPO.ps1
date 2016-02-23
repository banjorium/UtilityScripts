#========================================================================
# GPOreport
# Created by: Micky Balladelli
#
# This script goes through all the GPOs in all the domains in the $domains array 
# and verifies them according to best practices.
#
# The script then generates an Excel file with all the data, and uses pivot tables to generate
# a chart for each domain in the forest.
# 
# It checks:
#    - Unlinked GPOs
#    - GPO links that are disabled
#	 - Disabled GPO
#	 - Empty GPOs
#	 - Enabled GPOs without settings
#	 - WMI filters used in GPOs, WMI filter info is retrieved such as name, author, and code
#	 - GPOs with tombstone owners
#
#
#========================================================================
 
$domains = @()
$domains += New-Object -TypeName PSCustomObject -Property @{
			domain = "snhu.edu";
			server = "snhu-dc8.snhu.edu";
			}
#$domains += New-Object -TypeName PSCustomObject -Property @{
#			domain = "ad2.local";
#			server = "dc0.ad2.local";
#			}
 
# Import the necessary modules
try
{
	import-module grouppolicy -ErrorAction Stop
}
catch
{
	# the grouppolicy module requires at least Windows 7 or Windows Server 2008.
	[System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error")| Out-Null
}
 
# define a newline variable, it will be handy in the multiline textbox
$nl = "`r`n"
 
# define comments that will be used to give move info depending on the issue type
$comment0 = "Each GPO should have an identified owner. This is recommended in case of trouble-shooting,"+$nl+"and when modifications are required on the settings defined in a policy."
$comment1 = "GPOs that are disabled, or have no links, or all links disabled, or no data"+$nl+"should be removed."
$comment2 = "GPO has some disabled links. Those links should be removed."
$comment3 = "GPO has its Computer configuration active, however it is empty."+$nl+"it is recommended to disable unused configurations"
$comment4 = "GPO has its User configuration active, however it is empty."+$nl+"it is recommended to disable unused configurations"
$comment5 = "WMI filters may have a performance impact on the GPO target computers."
 
# save the location of the current script
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
 
# Sort a range of cells, this will make charts look nicer
function Sort-Range
{
	Param($range1, $range2)
 
	$i = 0
	$link1 = $range1.Cells.Next 
	$link2 = $range2.Cells.Next
	$newdata = @()
	$data = @()
	do 
	{ 
		$data += New-Object -TypeName PSCustomObject -Property @{
								domain 	= $link1.Value2;
								data 	= [int] $link2.Value2;
								}	
 
		$link1 = $link1.Next
		$link2 = $link2.Next			
 
		$i++
	} while ( $link1.Value2 -ne $null )
	$newdata = $data | Sort-Object -Descending -property data 
 
	$link1 = $range1.Cells.Next
	$link2 = $range2.Cells.Next
	for ($j = 0; $j -le $i ; $j++) 
	{
		$link1.Value2 = $newdata[$j].domain
		$link2.Value2 = $newdata[$j].data
		$link1 = $link1.Next
		$link2 = $link2.Next							
	}
 
}
function Report-GPO
{
	Param($GPO, $GPOreport, [string]$issue, $issueID, $WMI)
 
	$linkArray = @()
	$disLinkArray = @()
 
	foreach($link in $GPOReport.GPO.LinksTo)
	{
		if ($link -ne $null)
 
		{
			if ($link.Enabled -eq $true)
			{
				$linkArray+= $link.SOMPath
			}
			else
			{
				$disLinkArray+= $link.SOMPath
			}
		}
	}
 
	$linkString = "$linkArray"
	$disLinkString = "$disLinkArray"
 
 
	$retReport = New-Object -TypeName PSCustomObject -Property @{
			                    Name              = $GPO.Displayname;
								IssueFound		  = $issue;
			                    GPOStatus         = $GPO.GpoStatus;
								Owner	 		  = $GPO.Owner;
								Description		  = $GPO.Description;
								Links			  = $linkString;
								DisabledLinks	  = $disLinkString;
								DisabledLinkArray = $disLinkArray;
								GUID			  = $GPO.Id;
								IssueID			  = $issueID;
								CreationTime	  = $GPO.CreationTime;
								ModificationTime  = $GPO.ModificationTime;
			                    WMIFilter         = $GPO.WMIFilter.Name;
			                    WMIdata		      = $WMI;
			                    Path              = $GPO.Path
							} 
	return $retReport
}
 
# the following function's code is from Glenn Sizemore to convert a CN to DN
function ConvertFrom-Canonical 
{
	param([string]$canoincal=(throw '$Canonical is required!'))
 
	$obj = $canoincal.Replace(',','\,').Split('/')
    [string]$DN = "CN=" + $obj[$obj.count - 1]
 
	for ($i = $obj.count - 2;$i -ge 1;$i--){$DN += ",OU=" + $obj[$i]}
    $obj[0].split(".") | ForEach-Object { $DN += ",DC=" + $_}
    return $dn
}
 
function SaveCSV
{
	$fileout = $dir + "\report.csv"
 
	if ($report -ne $null)
	{
		$report |
	 	 Select-Object Name, IssueFound, Links, DisabledLinks, GPOStatus, CreationTime, ModificationTime, WMIFilter, WMIdata, Owner, Path | Export-CSV $fileout -NoTypeInformation
	}
}
 
 
 
# Load excel and define constant enums
[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")|Out-Null
# Excel enums
$xlConditionValues 	= [Microsoft.Office.Interop.Excel.XLConditionValueTypes]
$xlTheme		   	= [Microsoft.Office.Interop.Excel.XLThemeColor]
$xlChart 			= [Microsoft.Office.Interop.Excel.XLChartType]
$xlIconSet			= [Microsoft.Office.Interop.Excel.XLIconSet]
$xlDirection		= [Microsoft.Office.Interop.Excel.XLDirection]
$xlIcon				= [Microsoft.Office.Interop.Excel.XlIcon]
$xlPosition			= [Microsoft.Office.Interop.Excel.XlLegendPosition]
 
# Excel constants
$xlPivotTableVersion12     = 3
$xlPivotTableVersion10     = 1
$xlCount                   = -4112
$xlDescending              = 2
$xlDatabase                = 1
$xlHidden                  = 0
$xlRowField                = 1
$xlColumnField             = 2
$xlPageField               = 3
$xlDataField               = 4 
$xlCenter 				   = -4108
 
# Start a new Excel instance and create a workbook
$Excel = New-Object Microsoft.Office.Interop.Excel.ApplicationClass
$Excel.Visible = $True 
$wb = $Excel.Workbooks.Add()
 
$wb.worksheets.Item(1).delete()
$wb.worksheets.Item(1).delete()
 
# Loop through the list of domains (defined in an array at the beginning of this script
foreach ($domainElem in $domains)
{
	$domain = $domainElem.domain
	$server = $domainElem.server
	$ws = $wb.Worksheets.Add()
	$ws.name=$domain
 
	# Generate the data titles
	$ws.Cells.Item(1,1).Value2 = $domain
	$ws.Cells.Item(1,1).font.bold = 1
	$ws.Cells.Item(1,1).font.size=18
	$ws.Cells.Item(18,1).Value2 = "Total number of GPOs:"
	$ws.Cells.Item(19,1).Value2 = "Total number of issues found:"
 
	$row = 20
	$ws.range("C20:P20").font.bold = 1
	$ws.Cells.Item($row,3).Value2 = "Name"
	$ws.Cells.Item($row,4).Value2 = "Issue Found"
	$ws.Cells.Item($row,5).Value2 = "Proposed resolution"
	$ws.Cells.Item($row,6).Value2 = "GPO Status"
	$ws.Cells.Item($row,7).Value2 = "Owner"
	$ws.Cells.Item($row,8).Value2 = "Description"
	$ws.Cells.Item($row,9).Value2 = "Links"
	$ws.Cells.Item($row,10).Value2 = "DisabledLinks"
	$ws.Cells.Item($row,11).Value2 = "GUID"
	$ws.Cells.Item($row,12).Value2 = "CreationTime"
	$ws.Cells.Item($row,13).Value2 = "ModificationTime"
	$ws.Cells.Item($row,14).Value2 = "WMIFilter"
	$ws.Cells.Item($row,15).Value2 = "WMIdata"
	$ws.Cells.Item($row,16).Value2 = "Path"
 
	$ws.Columns.AutoFit()|Out-Null
 
	$row++
	"Retrieving data for the $domain domain"| Out-Host
 
 
	$report = @()
	$GPOPolicies = @()
	[Int32]$count = 0;
 
	# Get all GPOs from the given domain
	try
	{
		if ($server.Length -eq 0)
		{
			[array]$GPOPolicies = Get-GPO -domain $domain -All
		}
		else
		{ 
			[array]$GPOPolicies = Get-GPO -domain $domain -server $server -All
		}
 
	}
	catch
	{
	    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error")| Out-Null 
		return
	}
 
	ForEach ($GPO in $GPOPolicies) 
	{
		if ($server.Length -eq 0)
		{
			[xml]$GPOReport = (get-gporeport -guid $GPO.Id -ReportType xml -Domain $domain )
		}
		else
		{ 
			[xml]$GPOReport = (get-gporeport -guid $GPO.Id -ReportType xml -Domain $domain -Server $server)
		}
 
		$count++;
 
		# test if the GPO's is linked
		if ($GPOReport.gpo.LinksTo -eq $null) 
		{		
			$report += Report-GPO $GPO $GPOReport "GPO has no links" 1
 
			# no point in checking the other points. This GPO is not used
			continue
		}
		else
		{
			# test if the GPO links are disabled
			if (-not (($GPOReport.GPO.LinksTo | Foreach {$_.Enabled}) -eq $true))
			{
				$report += Report-GPO $GPO $GPOReport "All GPO links are disabled" 1
			}	
			else
			{
				if (($GPOReport.GPO.LinksTo | Foreach {$_.Enabled}) -eq $false)
				{
					$report += Report-GPO $GPO $GPOReport "Some GPO links are disabled" 2
				}					
			}
		}
 
		# test if the GPO's has no settings
		if (!$GPOReport.gpo.Computer.ExtensionData -and !$GPOReport.GPO.User.ExtensionData)
		{
			$report += Report-GPO $GPO $GPOReport "GPO is empty" 1
 
			# no point in checking the other points. This GPO is empty
			continue
		}
		else
		{
			# test if the extension is empty and the GPO is enabled
			if (!$GPOReport.gpo.Computer.ExtensionData -and $GPO.GpoStatus -eq "ComputerSettingsEnabled")
			{
				$report += Report-GPO $GPO $GPOReport "Computer settings enabled without data" 3
			}
			if (!$GPOReport.gpo.Computer.ExtensionData -and $GPO.GpoStatus -eq "AllSettingsEnabled")
			{
				$report += Report-GPO $GPO $GPOReport "Computer settings enabled without data" 3
			}
			if (!$GPOReport.gpo.User.ExtensionData -and $GPO.GpoStatus -eq "UserSettingsEnabled")
			{
				$report += Report-GPO $GPO $GPOReport "User settings enabled without data" 4
			}
			if (!$GPOReport.gpo.User.ExtensionData -and $GPO.GpoStatus -eq "AllSettingsEnabled")
			{
				$report += Report-GPO $GPO $GPOReport "User settings enabled without data" 4
			}			
		}
 
		# test if the GPO is disabled
		if ($GPO.GpoStatus -eq "AllSettingsDisabled")
		{
			$report += Report-GPO $GPO $GPOReport "GPO is disabled" 1
		}	
 
		# test if the GPO's has a WMI filter
		$result = @{}
		if ($GPO.WMIFilter.Name.Length -gt 0)
		{
 
 
			# let's find out what this filter does
			$wmiFilterAttr = "msWMI-Name", "msWMI-Parm1", "msWMI-Parm2", "msWMI-Author", "msWMI-ID"
 
		 	$search = New-Object System.DirectoryServices.DirectorySearcher([ADSI]("LDAP://" + $domain))
			$wminame = $GPO.WMIFilter.Name
			$search.Filter = "(&amp;(objectClass=msWMI-Som)(msWMI-Name=$wminame))"
			$search.PropertiesToLoad.AddRange($wmiFilterAttr)
			$result = $search.FindOne()
			$WMI = New-Object -TypeName PSCustomObject -Property @{
						                    Name 	= [string]$result.Properties["mswmi-name"];
											Parm1	= [string]$result.Properties["mswmi-parm1"];
											Parm2	= [string]$result.Properties["mswmi-parm2"];
											Author  = [string]$result.Properties["mswmi-author"];
											ID	    = [string]$result.Properties["mswmi-ID"]
											}
 
			$report += Report-GPO $GPO $GPOReport "WMI filter found" 5 $WMI 
		}
 
		# test if the GPO's owner is a tumbstone
		$GPOowner =  $GPO.Owner | Out-String
 
		if ($GPOowner.Length -eq 0)
		{
			$report += Report-GPO $GPO $GPOReport "Owner not found" 0
		}
		else 
		{
			if ($GPOowner.StartsWith("S-") -eq $true)
			{
				$report += Report-GPO $GPO $GPOReport "Owner is a tombstone" 0
			}
		}
	}	
 
 
	# Generate the excel worksheet for this domain
	if ($report -ne $null)
	{
		"Generating Excel worksheet" | Out-Host
		foreach ($elem in $report)
		{
			$ws.Cells.Item($row,2).Value2 = [string]$elem.IssueID
			$ws.Cells.Item($row,3).Value2 = [string]$elem.Name
			$ws.Cells.Item($row,4).Value2 = [string]$elem.IssueFound
 
			switch ($elem.IssueID)
			{
				# Owner not found
				"0"
				{
					$ws.Cells.Item($row,5).Value2 = "Every GPO should have an identified owner"
					$ws.Cells.Item($row,4).AddComment($comment0) |Out-Null
				}
				# Remove the GPO
				"1"
				{
					$ws.Cells.Item($row,5).Value2 = "Remove the GPO"
					$ws.Cells.Item($row,4).AddComment($comment1)|Out-Null
				}
				# Remove disabled Links
				"2"
				{
					$ws.Cells.Item($row,5).Value2 = "Remove disabled links"
					$ws.Cells.Item($row,4).AddComment($comment2)|Out-Null
				}
				# Computer settings disabled
				"3"
				{
					$ws.Cells.Item($row,5).Value2 = "Status need to change to ComputerSettingsDisabled"
					$ws.Cells.Item($row,4).AddComment($comment3)|Out-Null
				}
				# User settings disabled
				"4"
				{
					$ws.Cells.Item($row,5).Value2 = "Status need to change to UserSettingsDisabled"
					$ws.Cells.Item($row,4).AddComment($comment4)|Out-Null
				}
				# WMI filter found
				"5"
				{
					$ws.Cells.Item($row,5).Value2 = "Redesign GPO recommended"
					$ws.Cells.Item($row,4).AddComment($comment5)|Out-Null
				}
			}
			$ws.Cells.Item($row,6).Value2 = [string]$elem.GPOStatus
			$ws.Cells.Item($row,7).Value2 = [string]$elem.Owner
			$ws.Cells.Item($row,8).Value2 = [string]$elem.Description
			$ws.Cells.Item($row,9).Value2 = [string]$elem.Links
			$ws.Cells.Item($row,10).Value2 = [string]$elem.DisabledLinks
			$ws.Cells.Item($row,11).Value2 = [string]$elem.GUID
			$ws.Cells.Item($row,12).Value2 = [string]$elem.CreationTime
			$ws.Cells.Item($row,13).Value2 = [string]$elem.ModificationTime
			$ws.Cells.Item($row,14).Value2 = [string]$elem.WMIFilter
			$ws.Cells.Item($row,15).Value2 = [string]$elem.WMIdata
			$ws.Cells.Item($row,16).Value2 = [string]$elem.Path
 
			$row++
		}
	}
 
	$ws.Columns.AutoFit()|Out-Null
 
	# Generate a Pivot table containing the total number of issues
	$start=$ws.range("D20")
	$selection=$ws.Range($start,$start.End($xlDirection::xlDown))
	$PivotTable = $wb.PivotCaches().Create($xlDatabase,$selection,$xlPivotTableVersion10)
	$PivotTable.CreatePivotTable("R2C1", $domain) | Out-Null
	$wb.ShowPivotTableFieldList = $true
 
	$PivotFields = $ws.PivotTables($domain).PivotFields("Issue Found")
	$PivotFields.Orientation = $xlRowField
	$PivotFields.Position = 1
	$PivotFields = $ws.PivotTables($domain).PivotFields("Issue Found")
	$PivotFields.Orientation = $xlDataField
	$PivotFields.Position = 1
 
	#Use format conditioning for the Issue ID
	$start=$ws.range("B21")
	$Selection=$ws.Range($start,$start.End($xlDirection::xlDown))
 
	$Selection.FormatConditions.AddIconSetCondition()|Out-Null 
	$Selection.FormatConditions.item($($Selection.FormatConditions.Count)).SetFirstPriority()|Out-Null
	$Selection.FormatConditions.item(1).ReverseOrder = $False
	$Selection.FormatConditions.item(1).ShowIconOnly = $True
	$Selection.FormatConditions.item(1).IconSet = $xlIconSet::xl3TrafficLights1
 
	$Selection.FormatConditions.item(1).IconCriteria.Item(1).Operator=7
	$Selection.FormatConditions.item(1).IconCriteria.Item(1).Icon = $xlIcon::xlIconGreenFlag
	$Selection.FormatConditions.item(1).IconCriteria.Item(2).Type=$xlConditionValues::xlConditionValueNumber
	$Selection.FormatConditions.item(1).IconCriteria.Item(2).Value=1
	$Selection.FormatConditions.item(1).IconCriteria.Item(2).Operator=7
	$Selection.FormatConditions.item(1).IconCriteria.Item(2).Icon = $xlIcon::xlIconRedFlag
	$Selection.FormatConditions.item(1).IconCriteria.Item(3).Type=$xlConditionValues::xlConditionValueNumber
	$Selection.FormatConditions.item(1).IconCriteria.Item(3).Value=2
	$Selection.FormatConditions.item(1).IconCriteria.Item(3).Operator=7
	$Selection.FormatConditions.item(1).IconCriteria.Item(3).Icon = $xlIcon::xlIconYellowFlag
 
	# Generate the chart with results for this domain
	$chart=$ws.Shapes.AddChart().Chart
 
	$col1=$ws.range("A1")
	$col1=$ws.Range($col1,$col1.End($xlDirection::xlDown))
	$col2=$ws.range("B1")
	$col2=$ws.Range($col2,$col2.End($xlDirection::xlDown))
	$chart.chartType=$xlChart::xlBarClustered
	$chartdata = $ws.Range($col1,$col2)
	$chart.SetSourceData($chartdata)
 
	$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
	$chart.ChartTitle.Text = ($domain+": "+$report.count+" issues found out of "+$GPOPolicies.Count+" GPOs")
	$chart.ChartTitle.Font = "Arial,10pt"
 
	# Save the total number of GPOs and issue count
	$ws.Cells.Item(18,2).Value2 = $GPOPolicies.Count
	$ws.Cells.Item(19,2).Value2 = $report.count
 
	$ws.shapes.item("Chart 1").top=$col2.Top+$col2.Width
	$ws.shapes.item("Chart 1").left=300
	$ws.shapes.item("Chart 1").width=700
}
 
# Generate the results worksheet with the summary of all the data
"Generating the results worksheet with the summary of all the data"| Out-Host
$ws = $wb.Worksheets.Add()
$ws.name="Results"
 
$ws.Cells.Range("E2").font.bold = 1
$ws.Cells.Range("E2").font.size = 17
$ws.Cells.Range("E3").font.bold = 1
$ws.Cells.Range("E3").font.size = 17
$ws.Cells.Range("F2").font.bold = 1
$ws.Cells.Range("F2").font.size = 17
$ws.Cells.Range("F3").font.bold = 1
$ws.Cells.Range("F3").font.size = 17
 
$ws.Cells.Range("L1").Value2  = "GPO report generated:"  
$ws.Cells.Range("L1").VerticalAlignment = $xlCenter 
$ws.Cells.Range("E2").Value2  = "Total number of GPOs"
$ws.Cells.Range("E3").Value2  = "Number of issues found"
 
$row = 150
$col = 5
$basecol = $col 
 
$ws.Cells.Item($row+1,$col).Value2  = "Total number of GPOs"
$ws.Cells.Item($row+3,$col).Value2  = "Total number of issues found"
$ws.Cells.Item($row+5,$col).Value2  = "% of issues vs GPOs"
$ws.Cells.Item($row+7,$col).Value2  = "Most GPOs required to be removed"
$ws.Cells.Item($row+9,$col).Value2  = "Most WMI scripts"
$ws.Cells.Item($row+11,$col).Value2 = "Most invalid configurations"
$ws.Cells.Item($row+13,$col).Value2 = "Most disabled links"
$ws.Cells.Item($row+15,$col).Value2 = "Most invalid owners"
 
foreach ($domainElem in $domains)
{
	$col++
	$domain = $domainElem.domain
 
	# Total number of GPOs found
	$ws.Cells.Item($row,$col).Value2 = $domain
	$ws.Cells.Item($row+1,$col).Value2 = $wb.sheets.item($domain).cells.range("B18").Value2
 
	# Total number of issues found
	$ws.Cells.Item($row+2,$col).Value2 = $domain
	$ws.Cells.Item($row+3,$col).Value2 = $wb.sheets.item($domain).Cells.range("B19").Value2
 
	# % of issues vs GPOs
	$ws.Cells.Item($row+4,$col).Value2 = $domain
	$ws.Cells.Item($row+5,$col).Value2 = [Int32]($ws.Cells.Item($row+3,$col).Value2*100/$ws.Cells.Item($row+1,$col).Value2)
 
	$wsDom = $wb.sheets.item($domain)
	$range=$wsDom.range("B21")
	$range=$wsDom.Range($range,$range.End($xlDirection::xlDown))
 
	# Most GPO required to be removed
	$count = 0
	foreach ($item in $range)
	{
		if ($item.value2 -eq 1)
		{
			$count++
		}
	}
	$ws.Cells.Item($row+6,$col).Value2 = $domain
	$ws.Cells.Item($row+7,$col).Value2 = $count
 
	# Most WMI scripts
	$count = 0
	foreach ($item in $range)
	{
		if ($item.value2 -eq 5)
		{
			$count++
		}
	}
	$ws.Cells.Item($row+8,$col).Value2 = $domain
	$ws.Cells.Item($row+9,$col).Value2 = $count
 
	# Most invalid configurations
	$count = 0
	foreach ($item in $range)
	{
		if ($item.value2 -eq 3 -or $item.value2 -eq 4)
		{
			$count++
		}
	}
	$ws.Cells.Item($row+10,$col).Value2 = $domain
	$ws.Cells.Item($row+11,$col).Value2 = $count
 
	# Most disabled links
	$count = 0
	foreach ($item in $range)
	{
		if ($item.value2 -eq 2)
		{
			$count++
		}
	}
	$ws.Cells.Item($row+12,$col).Value2 = $domain
	$ws.Cells.Item($row+13,$col).Value2 = $count	
 
	# Most invalid owners
	$count = 0
	foreach ($item in $range)
	{
		if ($item.value2 -eq 0)
		{
			$count++
		}
	}
	$ws.Cells.Item($row+14,$col).Value2 = $domain
	$ws.Cells.Item($row+15,$col).Value2 = $count	
}
$ws.Columns.AutoFit()|Out-Null
 
#Compute total number of GPOs and issues
$range= $ws.range("F151")
$range= $ws.Range($range,$range.End($xlDirection::xlToRight))
$count = 0
foreach ($item in $range)
{
	$count += $item.value2
}
 
$ws.Cells.Range("M1").Value2  = Get-Date |Out-String
$ws.Cells.Range("F2").Value2  = $count
 
$range= $ws.range("F153")
$range= $ws.Range($range,$range.End($xlDirection::xlToRight))
$count = 0
foreach ($item in $range)
{
	$count += $item.value2
}
$ws.Cells.Range("F3").Value2  = $count
 
#Create the final results charts
"Creating the final charts"|Out-Host
 
 
# Total number of GPOs found
$range1=$ws.range("E150")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E151")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 1").top=$ws.Range("A6").Top
$ws.shapes.item("Chart 1").left=$ws.Range("A6").Left
$ws.shapes.item("Chart 1").height=250
# Total number of issues found
$range1=$ws.range("E152")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E153")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 2").top=$ws.Range("F6").Top
$ws.shapes.item("Chart 2").left=$ws.Range("F6").Left
$ws.shapes.item("Chart 2").height=250
 
# % issues vs GPOs
$range1=$ws.range("E154")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E155")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 3").top=$ws.Range("A23").Top
$ws.shapes.item("Chart 3").left=$ws.Range("A23").Left
$ws.shapes.item("Chart 3").height=250
 
# Most GPO required to be removed
$range1=$ws.range("E156")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E157")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 4").top=$ws.Range("F23").Top
$ws.shapes.item("Chart 4").left=$ws.Range("F23").Left
$ws.shapes.item("Chart 4").height=250
 
# Most WMI scripts
$range1=$ws.range("E158")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E159")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 5").top=$ws.Range("A40").Top
$ws.shapes.item("Chart 5").left=$ws.Range("A40").Left
$ws.shapes.item("Chart 5").height=250
 
# Most invalid configurations
$range1=$ws.range("E160")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E161")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 6").top=$ws.Range("F40").Top
$ws.shapes.item("Chart 6").left=$ws.Range("F40").Left
$ws.shapes.item("Chart 6").height=250
 
# Most disabled links
$range1=$ws.range("E162")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E163")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 7").top=$ws.Range("A57").Top
$ws.shapes.item("Chart 7").left=$ws.Range("A57").Left
$ws.shapes.item("Chart 7").height=250
 
# Most disabled links
$range1=$ws.range("E164")
$range1=$ws.Range($range1,$range1.End($xlDirection::xlToRight))
$range2=$ws.range("E165")
$range2=$ws.Range($range2,$range2.End($xlDirection::xlToRight))
Sort-Range $range1 $range2
$selection = $ws.Range($range1, $range2)
$chart=$ws.Shapes.AddChart().Chart
$chart.chartType=$xlChart::xl3DPieExploded
$chart.HasDataTable = 1
$chart.SetSourceData($selection)
$chart.SeriesCollection(1).HasDataLabels = "true"
$chart.SeriesCollection(1).ApplyDataLabels() | out-Null
$ws.shapes.item("Chart 8").top=$ws.Range("F57").Top
$ws.shapes.item("Chart 8").left=$ws.Range("F57").Left
$ws.shapes.item("Chart 8").height=250
 
# Remove last unused worksheet
$wb.worksheets.Item($wb.Worksheets.Count).delete()
 
# Save and done
$excel.DisplayAlerts = $False
$wb.SaveAs(($dir + "\report.xlsx"))
"Done"|Out-Host