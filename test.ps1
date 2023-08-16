# Add a reference to the Microsoft Excel Object Library
$excel = New-Object -ComObject Excel.Application

# Open the Excel file
$workbook = $excel.Workbooks.Open("C:/filepath")

# Select the worksheet
$worksheet = $workbook.Sheets.Item(1)

# Get the used range of the worksheet
$usedRange = $worksheet.UsedRange

# Get the column index where you want to search for "Project Manager"
$searchColumn = 3  # Adjust this to the actual column index (1-based)

# Get the dimensions of the used range
$rows = $usedRange.Rows.Count
$columns = $usedRange.Columns.Count

# Initialize an array to store data from matching rows
$matchingRowData = @()

# Loop through rows to collect data from matching rows
for ($row = 1; $row -le $rows; $row++) {
    $cellValue = $usedRange.Item($row, $searchColumn).Value2
    if ($cellValue -eq "Project Manager") {
        $rowData = @()
        for ($col = 1; $col -le $columns; $col++) {
            $cellValue = $usedRange.Item($row, $col).Value2
            $rowData += $cellValue
        }
        $matchingRowData += , $rowData
    }
}

# Close Excel on Source Workbook
$workbook.Close()

# Open the destination Excel file
$destinationWorkbook = $excel.Workbooks.Open("C:/filepath")

# Select the worksheet
$destinationWorksheet = $destinationWorkbook.Sheets.Item(1)

# Determine the starting destination row in the destination worksheet
$destinationRowIndex = 7  # Adjust this to your desired starting row

# Loop through matching rows' data and write to the destination worksheet
foreach ($matchingRow in $matchingRowData) {
    for ($om = 0; $om -lt $matchingRow.Length; $om++) {
        $destinationWorksheet.Cells.Item($destinationRowIndex, $om + 1).Value2 = $matchingRow[$om]
    }
    
    # Increment the destination row index for the next row of data
    $destinationRowIndex++
}

# Save changes to the destination workbook and close it
$destinationWorkbook.Save()
$destinationWorkbook.Close()

$excel.Quit()

# Release COM objects and perform garbage collection
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($destinationWorksheet)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($destinationWorkbook)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
