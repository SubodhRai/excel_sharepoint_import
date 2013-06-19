#################################################
# Powershell script to import from Excel into a
# Sharepoint list.
# 
# Needs to be run on sharepoint server with Excel
#
# Matthew Hilton 2013-06-19
#
#################################################
# Excel Bits
#################################################
# Variables
$Incoming="C:\Temp\Incoming"					# Folder where excel files are stored
#################################################
# Setup Excel
#################################################
$excel = New-Object -ComObject excel.application
$excel.visible = $true
$excel.DisplayAlerts = $false
#################################################
# Create Dataset table
$ExcelData = new-object System.Data.DataSet
$ExcelData.Tables.Add("ImportData")
[void]$ExcelData.Tables["ImportData"].Columns.Add("Request_Type",[string])
[void]$ExcelData.Tables["ImportData"].Columns.Add("Title",[string])
[void]$ExcelData.Tables["ImportData"].Columns.Add("Office",[string])
[void]$ExcelData.Tables["ImportData"].Columns.Add("User_ID",[string])
#################################################
# Process Excel Files
foreach ($ExcelFile in (Get-childitem -path $Incoming -filter *.xls)){ # Files in the incoming folder will need to be excel 2003 format
$ExcelFile.Fullname
$excelfileFull=$ExcelFile.Fullname
# Open Excel 2003 workbook
$ImportData=$excel.workbooks.open($excelfileFull) 
$ImportDataSheet=$ImportData.worksheets | where {$_.name -eq "ImportData"} # Selects ImportData sheet
Write-Host "ImportData File is being processed"
# Get last row, this doesn't always work, if the sheet has been edited you may end up added blank data. See Column C check.
$LastRow=$null
$LastRow=($ImportDataSheet.UsedRange.Rows.Count)
# Reset top row count to row 2, this is done to exclude the header row.
$Row=$null
$Row=2
Do {
$cell=$null
$cell=$ImportDataSheet.Cells.Item($Row,3).Text # check to see if column C is blank
IF (!$cell) {$LastRow=$Row-1} #IF cell is null set last row to $row -1
IF ($cell){
# Loop through Excel rows
    $ImportDataRow = $ExcelData.Tables["ImportData"].NewRow()
    $ImportDataRow["Request_Type"] = $ImportDataSheet.Cells.Item($Row,1).Text
    $ImportDataRow["Title"] = $ImportDataSheet.Cells.Item($Row,2).Text
    $ImportDataRow["Office"] = $ImportDataSheet.Cells.Item($Row,3).Text
	$ImportDataRow["User_ID"] = $ImportDataSheet.Cells.Item($Row,4).Text
    $ExcelData.Tables["ImportData"].Rows.Add($ImportDataRow)
	$ImportDataRow=$null
}
$Row++
}
While ($Row -le $LastRow)
#################################################
# Close Workbook
$ImportData.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ImportData) | Out-Null
}
#################################################
# Exit Excel
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers() 
#################################################
# Data in Excel sheet is now in powershell Dataset
#################################################
# Sharepoint bits
#################################################
## Load SharePoint Powershell Snapin
Add-PsSnapin Microsoft.SharePoint.PowerShell
#################################################
# Setup Variables for Sharepoint
# Root Site URL
$siteURL="http://abc.cbd.com"
# List Name
$listname="Log"

# Pull all items in the $listname
$site=Get-SPSite $siteURL
$web=$site.RootWeb
$list=$web.Lists[$listname]

foreach ($item in ($ExcelData.Tables["ImportData"]))
{
# Create New Item sharepoint list item
$newitem=$list.items.Add()
$newitem["Request_Type"]=$item["Request_Type"] 
$newitem["Title"]=$item["Title"]     
$newitem["Office"]=$item["Office"]    
$newitem["User_ID"]=$item["User_ID"]   
$newitem.update()
}

