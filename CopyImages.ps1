# Prompt user for input
$excelPath = Read-Host -Prompt 'Enter the full path to your Excel file (e.g., C:\Users\YourName\Documents\catalogue.xlsx)'
$sheetName = Read-Host -Prompt 'Enter the name of the sheet in the Excel file (e.g., Sheet1)'
$columnName = Read-Host -Prompt 'Enter the name of the column containing catalogue numbers'
$sourceFolder = Read-Host -Prompt 'Enter the full path to the folder containing your images (e.g., C:\Users\YourName\Pictures)'

# Get the Excel file name without extension
$excelFileName = [System.IO.Path]::GetFileNameWithoutExtension($excelPath)

# Create destination folder as a subfolder where the script is run
$scriptFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$destinationFolder = Join-Path -Path $scriptFolder -ChildPath $excelFileName

# Create destination folder if it doesn't exist
if (!(Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Force -Path $destinationFolder
}

Write-Host "Images will be copied to: $destinationFolder"

# Load Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item($sheetName)

# Find the column with catalogue numbers
$range = $sheet.UsedRange
$columns = $range.Columns.Count
$catalogueColumn = 1
for ($i = 1; $i -le $columns; $i++) {
    if ($sheet.Cells.Item(1, $i).Text -eq $columnName) {
        $catalogueColumn = $i
        break
    }
}

# Read catalogue numbers
$catalogueNumbers = @()
$row = 2
while ($sheet.Cells.Item($row, $catalogueColumn).Text -ne "") {
    $catalogueNumbers += $sheet.Cells.Item($row, $catalogueColumn).Text
    $row++
}

# Close Excel
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Function to get the folder name for a catalogue number
function Get-FolderName {
    param($catalogueNumber)
    $parts = $catalogueNumber.Split('.')
    if ($parts.Count -ge 3) {
        $prefix = "$($parts[0]).$($parts[1])"
        $number = [int]$parts[2]
        $start = [Math]::Floor($number / 1000) * 1000 + 1
        $end = $start + 999
        return "$prefix.$start-$prefix.$end"
    }
    return $null
}

# Copy files
foreach ($number in $catalogueNumbers) {
    $folderName = Get-FolderName $number
    if ($folderName) {
        $searchPath = Join-Path -Path $sourceFolder -ChildPath $folderName
        if (Test-Path $searchPath) {
            $files = Get-ChildItem -Path $searchPath -File | Where-Object {
                $_.Extension -match '\.(jpg|tif|dng)$' -and
                ($_.BaseName -eq $number -or $_.BaseName -like "$number *")
            }
            foreach ($file in $files) {
                $destinationPath = Join-Path -Path $destinationFolder -ChildPath $file.Name
                Copy-Item -Path $file.FullName -Destination $destinationPath
                Write-Host "Copied: $($file.FullName)"
            }
        } else {
            Write-Host "Folder not found for catalogue number: $number"
        }
    } else {
        Write-Host "Could not determine folder for catalogue number: $number"
    }
}

Write-Host "Process completed. Images copied to: $destinationFolder"