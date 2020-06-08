#FORMAT EXCEL FILES FOR PRINTING TRANSFER COSTS FORMS IN THE RIGHT FORMAT
#THEN CREATE PDF FILE
#Author: S. Giodini
#Version: 11-3-2020
write-host "start"
#prepare paths and shortcuts
[string]$path = "C:\Users\SGiodini\Rode Kruis\510 - Staff - Monthly Hour Transfers\may2020"  
#Path to Excel spreadsheets to save to PDF
[string]$savepath = "C:\Users\SGiodini\Rode Kruis\510 - Staff - Monthly Hour Transfers\may2020"
[string]$dToday = Get-Date -Format "yyyyMMdd"

#get all invoices (previously downloaded from forecast as csv in the $path)
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] 
#$excelFiles = Get-ChildItem -Path $path -include *.xls, *.xlsx -recurse | Where-Object { $_.FullName -match 'PRJ*' } 
$excelFiles = Get-ChildItem $path *.csv


# Create the Excel application object
$objExcel = New-Object -ComObject excel.application 
$objExcel.visible = $false   #Do not open individual windows

foreach($wb in $excelFiles) 
{ 
# Path to new PDF with date 
 $filepath = Join-Path -Path $savepath -ChildPath ($wb.BaseName + "_" + $dtoday + ".pdf") 
 # Open workbook - 3 refreshes links
 $workbook = $objExcel.workbooks.open($wb.fullname, 3)
 #apply texttocolumns
 $sheet=$workbook.Worksheets.item(1)
 $sheet.activate
 $range=$sheet.usedrange
 $colA=$sheet.range("A1").EntireColumn
 $colrange=$sheet.range("A1")
 $xlDelimited = 1
 $xlTextQualifier = 2
 $xlTextFormat = 2
 $colA.texttocolumns($colrange,$xlDelimited,$xlTextQualifier,$false,$false,$false,$true,$false)
 #AutoFit Columns
 $sheet.columns.autofit() 
 $workbook.RefreshAll()
 
 # Give delay to save
 Start-Sleep -s 5
 
 # Save Workbook with the required formatting
 #Landscape orientation
 $workbook.Worksheets.Item(1).PageSetup.Orientation = 2
 #required footer for approval by P&C
 $workbook.Worksheets.Item(1).pageSetup.LeftFooter = "Transfer from KPL748 to project 67EUR/h"
 $workbook.Worksheets.Item(1).pageSetup.CenterFooter = "Prepared by Giodini, Stefania &D"
 $workbook.Worksheets.Item(1).pageSetup.RightFooter = "Page &P of &N"
 #fit to one page only
 $workbook.Worksheets.Item(1).pageSetup.Zoom = $false
 $workbook.Worksheets.Item(1).pageSetup.FitToPagesWide = 1
 $workbook.Worksheets.Item(1).pageSetup.FitToPagesTall = 1
 $workbook.Worksheets.Item(1).PageSetup.PrintGridlines = $true
 $workbook.Saved = $true 
"saving $filepath" 
 #Export as PDF
 $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath) 
 $objExcel.Workbooks.close() 
} 
$objExcel.Quit()