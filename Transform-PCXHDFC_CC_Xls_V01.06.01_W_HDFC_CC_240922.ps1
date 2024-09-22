<#
.SYNOPSIS
    This script automates the conversion of `.xls` files to `.xlsx` format and then transforms the data in the resulting `.xlsx` files.
    It processes files located in the current directory where the script is executed, ensuring efficient handling and transformation of data.
    This code is used for HDFC bank credit card data to transfer to different format.
    Filename should beging with HDFC_CC_<XX>_. XX refers to any charector for identification of the file. Preferably 2 charector of the person name.

.DESCRIPTION
    The script performs the following tasks:
    1. **Initial Setup:**
       - Clears the host screen for a clean output.
       - Retrieves the path of the current directory where the script is located.
    
    2. **Conversion of `.xls` to `.xlsx`:**
       - Identifies all `.xls` files in the folder with the naming pattern `HDFC_CC_*.xls`.
       - Uses Excel COM object to open each `.xls` file, convert it to `.xlsx` format, and save it with the suffix `_ConvertedFromXls`.
       - Logs the conversion progress to the console.
    
    3. **Processing and Transformation of `.xlsx` Files:**
       - Identifies all `.xlsx` files in the folder, including those converted from `.xls`.
       - Extracts the "MOP" (Method of Payment) from specific segments of the file name to be included in the transformed data.
       - Opens each `.xlsx` file and processes it:
         - Finds header positions within the first 50 rows of the sheet.
         - Creates a new workbook for storing transformed data.
         - Writes new headers to the new worksheet.
         - Transfers and formats data based on the identified headers and the MOP.
         - Uses a custom function to normalize date formats to `dd/MM/yyyy`.
         - Ensures that certain columns (e.g., "Item", "Category", etc.) remain blank as per the requirements.
         - Concatenates "Transaction Remarks" across multiple rows when other related columns are empty.
       - Saves the transformed data with the suffix `_Transformed`.
       - Logs the transformation progress to the console.
    
    4. **Cleanup:**
       - Closes all open workbooks and releases COM objects to free up resources.
       - Ensures that the Excel application quits and cleans up any remaining COM objects to prevent memory leaks.

.FEATURES
    - **Dynamic File Handling:** Automatically detects and processes relevant files based on their extensions and naming patterns.
    - **Custom Date Formatting:** A robust function to handle various date formats and ensure consistency in the `dd/MM/yyyy` format.
    - **Blank Column Handling:** Ensures that specific columns remain blank as needed.
    - **Detailed Logging:** Provides real-time feedback and logs significant actions and errors during processing.
    - **Resource Management:** Properly releases COM objects and performs garbage collection to avoid memory issues.

.ERROR_HANDLING
    - The script includes error handling for opening workbooks and saving transformed files, logging errors to the console when encountered.
    - It ensures that operations on files are executed only if the files are successfully opened, maintaining script stability.

#>

# Clear the host screen
Clear-Host

# Get the current directory (script root)
$scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Convert .xls files to .xlsx with a "_ConvertedFromXls" suffix
# Get all .xls files in the folder
$xlsFiles = Get-ChildItem -Path $scriptRoot -Filter HDFC_CC_*.xls

# Create an Excel COM object for conversion
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($file in $xlsFiles) {
    # Define file paths for conversion
    $xlsFilePath = $file.FullName
    $xlsxFilePath = [System.IO.Path]::Combine($scriptRoot, "$($file.BaseName)_ConvertedFromXls.xlsx")

    # Open the .xls file
    $workbook = $excel.Workbooks.Open($xlsFilePath)
    
    # Save as .xlsx file
    $workbook.SaveAs($xlsxFilePath, 51)  # 51 is the Excel constant for .xlsx

    # Close the workbook
    $workbook.Close()
    
    Write-Host "Conversion complete: $xlsFilePath -> $xlsxFilePath"
}

# Quit Excel for conversion step
$excel.Quit()

# Release COM objects for conversion
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

# Now process all .xlsx files in the folder (including converted ones)
$excelFiles = Get-ChildItem -Path $scriptRoot -Filter HDFC_CC_*.xlsx

# Create a new Excel COM object for transformation
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($excelFile in $excelFiles) {
    $excelFilePath = $excelFile.FullName

    # Extract MOP from the file name
    $fileNameParts = (Get-Item $excelFilePath).BaseName -split '_'
    $MOP = "$($fileNameParts[0])_$($fileNameParts[1])_$($fileNameParts[2])"

    # Open the existing Excel file
    try {
        $workbook = $excel.Workbooks.Open($excelFilePath)
        Write-Host "Opened Excel file: $excelFilePath"
    } catch {
        Write-Host "Failed to open Excel file: $excelFilePath"
        continue
    }

    $worksheet = $workbook.Worksheets.Item(1)
    Write-Host "Accessed the first worksheet."

# Function to parse dates and extract only the date part (dd/MM/yyyy)
function Get-FormattedDate {
    param (
        [string]$inputString
    )

    # Take the first 10 characters and replace '/' or '-' with '-'
    $formattedDate = $inputString.Substring(0, [Math]::Min(10, $inputString.Length)) -replace '[/-]', '-'

    return $formattedDate
}

    # Define the headers to be found and their new names
    $requiredHeaders = @{
        "DATE" = "Date"
        "Description" = "Narration"
        "AMT" = "Amt (Dr)"
        "Debit / Credit" = "Value Dt"
    }

    # Define the new headers and their order
    $newHeaders = @("Date", "Narration", "Item", "Catogery", "Place", "Freq", "For", "MOP", "Amt (Dr)", "Chq./Ref.No.", "Value Dt", "Amt (Cr)")

    # Find the header rows and columns
    $headerPositions = @{ }
    foreach ($header in $requiredHeaders.Keys) {
        for ($row = 27; $row -le 50; $row++) {  # Dynamically search within the first 50 rows
            for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                $cellValue = $worksheet.Cells.Item($row, $col).Value2
                if ($cellValue -eq $header) {
                    $headerPositions[$header] = @{ Row = $row; Column = $col }
                    Write-Host "Found required header: $header at row $row, column $col"
                    break
                }
            }
            if ($headerPositions.ContainsKey($header)) { break }
        }
    }

    # Check if all required headers were found
    if ($headerPositions.Count -ne $requiredHeaders.Count) {
        $missingHeaders = $requiredHeaders.Keys | Where-Object { -not $headerPositions.ContainsKey($_) }
        Write-Host "Not all required headers were found: $($missingHeaders -join ', ')"
        $workbook.Close($false)
        continue
    }

    # Create a new workbook for the filtered data
    $newWorkbook = $excel.Workbooks.Add()
    $newWorksheet = $newWorkbook.Worksheets.Item(1)

    # Write the new headers to the new worksheet
    $colIndex = 1
    foreach ($newHeader in $newHeaders) {
        $newWorksheet.Cells.Item(1, $colIndex) = $newHeader
        Write-Host "Filtered header '$newHeader' written to new worksheet."
        $colIndex++
    }

    # Write the filtered data rows to the new worksheet
$rowIndex = 2
for ($i = $headerPositions["DATE"].Row + 1; $i -le $worksheet.UsedRange.Rows.Count; $i++) {
    $colIndex = 1

    foreach ($newHeader in $newHeaders) {
        $data = ""

        switch ($newHeader) {
            "Date" {
                #$dateValue = $worksheet.Cells.Item($i, $headerPositions["DATE"].Column).Text
                $dateValue = $worksheet.Cells.Item($i, $headerPositions["DATE"].Column).Value2
                $data = Get-FormattedDate $dateValue

                # Set cell format to "Text" only for date columns
                $newWorksheet.Cells.Item($rowIndex, $colIndex).NumberFormat = "@"
            }
            "Narration" {
                $data = $worksheet.Cells.Item($i, $headerPositions["Description"].Column).Value2
            }
            "Amt (Dr)" {
                $debitCredit = $worksheet.Cells.Item($i, $headerPositions["Debit / Credit"].Column).Value2
                if ($debitCredit -ne "Cr") {
                    $data = $worksheet.Cells.Item($i, $headerPositions["AMT"].Column).Value2
                } else {
                    $data = ""
                }
            }
            "Amt (Cr)" {
                $debitCredit = $worksheet.Cells.Item($i, $headerPositions["Debit / Credit"].Column).Value2
                if ($debitCredit -eq "Cr") {
                    $data = $worksheet.Cells.Item($i, $headerPositions["AMT"].Column).Value2
                } else {
                    $data = ""
                }
            }
            "Value Dt" {
                $data = $worksheet.Cells.Item($i, $headerPositions["Debit / Credit"].Column).Value2
            }
            "MOP" {
                $data = $MOP
            }
        }

        # Ensure data is converted to string to avoid type mismatch
        $newWorksheet.Cells.Item($rowIndex, $colIndex).Value2 = [string]$data

        Write-Host "Processed row $i, column $colIndex ($newHeader): $data"
        $colIndex++
    }
    $rowIndex++
}




    # Define transformed file path
    $newExcelFilePath = [System.IO.Path]::Combine($scriptRoot, "$($excelFile.BaseName)_Transformed.xlsx")

    # Save the new workbook as XLSX
    try {
        $newWorkbook.SaveAs($newExcelFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
        Write-Host "Filtered data written to $newExcelFilePath"
    } catch {
        Write-Host "Failed to save the new Excel file: $newExcelFilePath"
    }

    # Close the workbooks and release COM objects
    $newWorkbook.Close()
    $workbook.Close($false)
}

# Quit Excel
$excel.Quit()

# Release COM objects for transformation
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($newWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
