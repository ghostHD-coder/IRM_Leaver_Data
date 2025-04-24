# Please modify the script to align with your specific environment.


# Install the ImportExcel module if not already installed

Set-ExecutionPolicy Unrestricted -Scope CurrentUser

Install-Module -Name ImportExcel -Force -Scope CurrentUser

 

# Define paths for the Excel files

$file1Path = "C:\Users\dimdi\Downloads\irm_hr_leaver_data\report1.xlsx"

$file2Path = "C:\Users\dimdi\Downloads\irm_hr_leaver_data\report2.xlsx"

$mergedFilePath = "C:\Users\dimdi\Downloads\irm_hr_leaver_data\merged_file.xlsx"

 

# Load the Excel files

$file1 = Import-Excel -Path $file1Path

$file2 = Import-Excel -Path $file2Path

 

# Merge the data on 'Username'

$mergedData = @()

foreach ($row1 in $file1) {

    foreach ($row2 in $file2) {

        if ($row1.'username' -eq $row2.'username') {

            $mergedRow = $row1 | Select-Object *, @{Name='UPN'; Expression={$row2.'UPN'}}

            $mergedData += $mergedRow

        }

    }

}

# Export the merged data to a new Excel file

$mergedData | Export-Excel -Path $mergedFilePath

Write-Output "The data has been merged and saved to 'merged_file.xlsx'."

# Define the path to your Excel file for inserting the formula

$excelFilePath = $mergedFilePath

# Create an Excel application object

$excel = New-Object -ComObject Excel.Application

$excel.Visible = $true  # Set to $true if you want to see the Excel window

# Open the workbook

$workbook = $excel.Workbooks.Open($excelFilePath)

$worksheet = $workbook.Worksheets.Item(1)


# Get the last row with data in the "LastWorkingDate" column

$lastRow = $worksheet.Cells($worksheet.Rows.Count, 3).End(-4162).Row  # -4162 is the constant for xlUp


# Loop through each cell in the "LastWorkingDate" column and set the formula in the corresponding cell in column E

for ($row = 2; $row -le $lastRow; $row++) {

    $worksheet.Cells($row, 5).Formula = "=TEXT(C$row, ""yyyy-mm-dd"") & ""T"" & TEXT(C$row,""hh:mm:ss"") & ""+01:00"""

}

# Convert formulas to values

for ($row = 2; $row -le $lastRow; $row++) {

    if ($worksheet.Cells($row, 5).Formula -ne $null) {

        $worksheet.Cells($row, 5).Value2 = $worksheet.Cells($row, 5).Value2

    }

}

 

# Add header to column E

$worksheet.Cells(1, 5).Value = "LastWorkingDate"

 

# Remove column C

$worksheet.Columns(3).Delete()

 

# Save and close the workbook

$workbook.Save()

$workbook.Close()

# Quit Excel

$excel.Quit()

# Release the COM objects

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# Collect garbage

[GC]::Collect()

[GC]::WaitForPendingFinalizers()

 

Write-Output "Workbook has been saved."

 

