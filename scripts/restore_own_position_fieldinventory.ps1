$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$systemRoot = 'C:\Users\kazu2\Documents\案件情報System'
$kernelPath = Join-Path $systemRoot '案件情報System_Kernel.xlsx'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupPath = "$kernelPath.$timestamp.bak"

Copy-Item -LiteralPath $kernelPath -Destination $backupPath -Force

$excel = $null
$workbook = $null
$worksheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($kernelPath)
    $worksheet = $workbook.Worksheets.Item('CaseList_FieldInventory')

    $targetRow = 22
    $existingKey = [string]$worksheet.Cells.Item($targetRow, 6).Text

    if ($existingKey -ne '当方_地位') {
        $worksheet.Rows.Item($targetRow).Insert()
    }

    $lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row

    $worksheet.Cells.Item($targetRow, 1).Value2 = 21
    $worksheet.Cells.Item($targetRow, 2).Value2 = '当方'
    $worksheet.Cells.Item($targetRow, 3).Value2 = 'B21'
    $worksheet.Cells.Item($targetRow, 4).Value2 = '当方_地位'
    $worksheet.Cells.Item($targetRow, 5).Value2 = ''
    $worksheet.Cells.Item($targetRow, 6).Value2 = '当方_地位'
    $worksheet.Cells.Item($targetRow, 7).Value2 = 'cf_r021'

    for ($row = $targetRow + 1; $row -le $lastRow; $row++) {
        $sequence = $row - 1
        $worksheet.Cells.Item($row, 1).Value2 = $sequence
        $worksheet.Cells.Item($row, 3).Value2 = 'B' + $sequence
        $worksheet.Cells.Item($row, 7).Value2 = ('cf_r{0:D3}' -f $sequence)
    }

    $workbook.Save()
    $workbook.Close($true)
}
finally {
    if ($worksheet -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
    if ($workbook -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Updated: $kernelPath"
Write-Output "Backup : $backupPath"
