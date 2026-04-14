# Script: Excel-URL-column-update.ps1
# Description: Updates the URL hyperlnks in a Excel column
# Full documentation: https://www.camerongriffiths.com/scripts/updateurlsinanexcelcolumn
# Author: Cameron Griffiths | camerongriffiths.com
# Requirements: PnP.PowerShell, Site Collection Administrator permissions

# Load the Excel file
$excelPath = "C:\path\to\your\file.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$worksheet = $workbook.Sheets.Item("YourWorksheetName")

# Get the last row with data in Column C (or change column number as needed)
$lastRow = $worksheet.Cells($worksheet.Rows.Count, 3).End(-4162).Row

Write-Host "Processing $lastRow rows..." -ForegroundColor Cyan

# Loop through Column C and update hyperlinks
$count = 0
for ($i = 2; $i -le $lastRow; $i++) {
    $cell = $worksheet.Cells.Item($i, 3)
    
    # Check if cell has a hyperlink
    if ($cell.Hyperlinks.Count -gt 0) {
        $hyperlink = $cell.Hyperlinks.Item(1)
        $oldUrl = $hyperlink.Address
        
        $updated = $false
        $newUrl = $oldUrl
        
        # Check for URL-encoded version (Old%20Value)
        if ($oldUrl -like "*Old%20Value*") {
            $newUrl = $oldUrl -replace "Old%20Value", "New%20Value"
            $updated = $true
        }
        # Check for space version (Old Value)
        elseif ($oldUrl -like "*Old Value*") {
            $newUrl = $oldUrl -replace "Old Value", "New Value"
            $updated = $true
        }
        # Check for additional pattern (Value you want to replace)
        elseif ($oldUrl -like "*Value you want to replace*") {
            $newUrl = $oldUrl -replace "Value you want to replace", "Replacement Value"
            $updated = $true
        }
        # Add more replacement patterns as needed
        elseif ($oldUrl -like "*AnotherOldValue*") {
            $newUrl = $oldUrl -replace "AnotherOldValue", "AnotherNewValue"
            $updated = $true
        }
        
        if ($updated) {
            $hyperlink.Address = $newUrl
            $count++
            Write-Host "Row $i - UPDATED" -ForegroundColor Green
            Write-Host "  Old: $oldUrl" -ForegroundColor Gray
            Write-Host "  New: $newUrl" -ForegroundColor Cyan
        }
    }
}

Write-Host "`nTotal URLs updated: $count" -ForegroundColor Yellow

# Save and close
$workbook.Save()
$workbook.Close()
$excel.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "`nDone! File saved successfully." -ForegroundColor Green
