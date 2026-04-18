[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]$ComObject
    )

    process {
        if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        }
    }
}

function New-BackupFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupPath = '{0}.{1}.bak' -f $Path, $timestamp
    Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    return $backupPath
}

function Get-OwnPositionLabel {
    $own = ([char]0x5F53).ToString() + [char]0x65B9
    $suffix = ([char]0x5730).ToString() + [char]0x4F4D
    return "$own`_$suffix"
}

function Get-KernelSheetNames {
    $user = ([char]0x30E6).ToString() + [char]0x30FC + [char]0x30B6 + [char]0x30FC + [char]0x60C5 + [char]0x5831
    $template = ([char]0x96DB).ToString() + [char]0x5F62 + [char]0x4E00 + [char]0x89A7
    $caseList = ([char]0x6848).ToString() + [char]0x4EF6 + [char]0x4E00 + [char]0x89A7

    return @($user, $template, $caseList)
}

$rootPath = Split-Path -Parent $PSScriptRoot
$basePath = (Get-ChildItem -LiteralPath $rootPath -Filter '*_Base.xlsx' | Select-Object -First 1 -ExpandProperty FullName)
$kernelPath = (Get-ChildItem -LiteralPath $rootPath -Filter '*_Kernel.xlsx' | Select-Object -First 1 -ExpandProperty FullName)

if ([string]::IsNullOrWhiteSpace($basePath) -or [string]::IsNullOrWhiteSpace($kernelPath)) {
    throw 'Base or Kernel workbook was not found.'
}

$baseBackup = New-BackupFile -Path $basePath
$kernelBackup = New-BackupFile -Path $kernelPath

$excel = $null
$baseWorkbook = $null
$kernelWorkbook = $null
$baseSheet = $null
$inventorySheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $label = Get-OwnPositionLabel
    $xlPasteFormats = -4122
    $xlPasteValidation = 6

    $baseWorkbook = $excel.Workbooks.Open($basePath)
    $baseSheet = $baseWorkbook.Worksheets.Item(1)

    if ([string]$baseSheet.Cells.Item(21, 1).Text -ne $label) {
        $baseSheet.Rows.Item(21).Insert()
        $baseSheet.Rows.Item(20).Copy() | Out-Null
        $baseSheet.Rows.Item(21).PasteSpecial($xlPasteFormats)
        $baseSheet.Rows.Item(21).PasteSpecial($xlPasteValidation)
        $baseSheet.Rows.Item(21).RowHeight = $baseSheet.Rows.Item(20).RowHeight
        $baseSheet.Cells.Item(21, 1).Value2 = $label
        $null = $baseSheet.Cells.Item(21, 2).ClearContents()
        $baseWorkbook.Save()
    }

    $baseWorkbook.Close($true)
    $baseWorkbook = $null
    $baseSheet = $null

    $kernelWorkbook = $excel.Workbooks.Open($kernelPath)

    $desiredSheetNames = Get-KernelSheetNames
    for ($index = 1; $index -le 3; $index++) {
        $worksheet = $kernelWorkbook.Worksheets.Item($index)
        if ([string]$worksheet.Name -ne $desiredSheetNames[$index - 1]) {
            $worksheet.Name = $desiredSheetNames[$index - 1]
        }
        Release-ComObject $worksheet
    }

    $inventorySheet = $kernelWorkbook.Worksheets.Item('CaseList_FieldInventory')

    if ([string]$inventorySheet.Cells.Item(22, 4).Text -ne $label) {
        $inventorySheet.Rows.Item(22).Insert()
        $inventorySheet.Rows.Item(21).Copy() | Out-Null
        $inventorySheet.Rows.Item(22).PasteSpecial($xlPasteFormats)
        $inventorySheet.Rows.Item(22).PasteSpecial($xlPasteValidation)
        $inventorySheet.Rows.Item(22).RowHeight = $inventorySheet.Rows.Item(21).RowHeight

        $inventorySheet.Cells.Item(22, 2).Value2 = $inventorySheet.Cells.Item(21, 2).Value2
        $inventorySheet.Cells.Item(22, 3).Value2 = 'B21'
        $inventorySheet.Cells.Item(22, 4).Value2 = $label
        $null = $inventorySheet.Cells.Item(22, 5).ClearContents()
        $inventorySheet.Cells.Item(22, 6).Value2 = $label
        $inventorySheet.Cells.Item(22, 7).Value2 = 'cf_r021'
        $inventorySheet.Cells.Item(22, 8).Value2 = 'Text'
        $inventorySheet.Cells.Item(22, 9).Value2 = 'Trim'
        $null = $inventorySheet.Cells.Item(22, 10).ClearContents()
    }

    $lastInventoryRow = $inventorySheet.Cells.Item($inventorySheet.Rows.Count, 1).End(-4162).Row
    for ($row = 2; $row -le $lastInventoryRow; $row++) {
        $baseRow = $row - 1
        $inventorySheet.Cells.Item($row, 1).Value2 = [string]$baseRow
        $inventorySheet.Cells.Item($row, 3).Value2 = "B$baseRow"
        $inventorySheet.Cells.Item($row, 7).Value2 = ('cf_r{0:000}' -f $baseRow)
    }

    $kernelWorkbook.Save()
    $kernelWorkbook.Close($true)
    $kernelWorkbook = $null
    $inventorySheet = $null

    [PSCustomObject]@{
        BaseBackup = $baseBackup
        KernelBackup = $kernelBackup
        AddedField = 'Added own position field'
        FixedSheets = 'Renamed first three Kernel sheets'
    } | Format-List
}
finally {
    if ($inventorySheet) { Release-ComObject $inventorySheet }
    if ($kernelWorkbook) { $kernelWorkbook.Close($false) | Out-Null; Release-ComObject $kernelWorkbook }
    if ($baseSheet) { Release-ComObject $baseSheet }
    if ($baseWorkbook) { $baseWorkbook.Close($false) | Out-Null; Release-ComObject $baseWorkbook }
    if ($excel) {
        $excel.EnableEvents = $true
        $excel.ScreenUpdating = $true
        $excel.DisplayAlerts = $true
        $excel.Quit()
        Release-ComObject $excel
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
