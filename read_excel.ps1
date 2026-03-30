$excelPath = Get-Item "c:\Users\callcenter\Desktop\dashboard\data\*INCENTIVOS*" | Select-Object -ExpandProperty FullName
Write-Host "Arquivo: $excelPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.ActiveSheet

Write-Host "Sheet Name: $($sheet.Name)"

$used = $sheet.UsedRange
$rows = $used.Rows.Count
$cols = $used.Columns.Count

Write-Host "Dimensions: $rows rows x $cols columns"
Write-Host "---"

# Show first 10 rows
for ($i=1; $i -le [Math]::Min(10, $rows); $i++) {
    $row = @()
    for ($j=1; $j -le [Math]::Min(15, $cols); $j++) {
        $val = $sheet.Cells($i,$j).Value
        $row += if ($null -eq $val) { "" } else { $val }
    }
    Write-Host ($row -join "`t")
}

$workbook.Close($false)
$excel.Quit()
