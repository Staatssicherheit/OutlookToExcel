# Define the paths for the source and destination Excel files
$sourceFile = "C:\path\to\source.xlsx"
$destinationFile = "C:\path\to\destination.xlsx"

# Create a new Excel application object
$excel = New-Object -ComObject Excel.Application

# Disable alerts to prevent prompts during copy/paste
$excel.DisplayAlerts = $false

# Open the source Excel file
$workbook = $excel.Workbooks.Open($sourceFile)

# Define the source worksheet and cell range to copy
$sourceWorksheet = $workbook.Sheets.Item("Sheet1") # Change "Sheet1" to your source sheet name
$sourceRange = $sourceWorksheet.Range("A1:B10")    # Change "A1:B10" to your source range

# Open the destination Excel file
$destinationWorkbook = $excel.Workbooks.Open($destinationFile)

# Define the destination worksheet and cell range to paste
$destinationWorksheet = $destinationWorkbook.Sheets.Item("Sheet1") # Change "Sheet1" to your destination sheet name
$destinationRange = $destinationWorksheet.Range("C1")             # Change "C1" to your destination cell

# Copy the source range
$sourceRange.Copy()

# Paste the copied range to the destination
$destinationWorksheet.Paste($destinationRange)

# Save changes to the destination workbook
$destinationWorkbook.Save()

# Close both workbooks
$destinationWorkbook.Close($true)
$workbook.Close($true)

# Quit Excel application
$excel.Quit()

# Release COM objects from memory
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($destinationWorksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sourceWorksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($destinationWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
