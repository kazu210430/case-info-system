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

function Get-AgentLabel {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    $positionSuffix = [string]([char]0x5730) + [char]0x4F4D
    return ("{0}_{1}" -f (Get-AgentBaseName -AgentNumber $AgentNumber), $positionSuffix)
}

function Get-AgentBaseName {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    $proxySuffix = [string]([char]0x4EE3) + [char]0x7406 + [char]0x4EBA
    return ("{0}{1}{2}" -f ([char]0x76F8 + [string][char]0x624B), $AgentNumber, $proxySuffix)
}

function Get-InsertedBaseRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    return 50 + (($AgentNumber - 1) * 21)
}

function Get-OriginalBaseInsertRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    return 50 + (($AgentNumber - 1) * 20)
}

function Get-OriginalInventoryInsertRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    return 51 + (($AgentNumber - 1) * 20)
}

function Get-AdjustedBaseRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$OriginalRow
    )

    $insertedRows = @(50, 71, 92, 113, 134, 155, 176, 197, 218, 239)
    $shiftCount = ($insertedRows | Where-Object { $_ -le $OriginalRow }).Count
    return $OriginalRow + $shiftCount
}

function Insert-BaseRows {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet
    )

    $xlPasteFormats = -4122
    $xlPasteValidation = 6

    foreach ($agent in 10..1) {
        $insertRow = Get-OriginalBaseInsertRow -AgentNumber $agent
        $sourceRow = $insertRow - 1
        $label = Get-AgentLabel -AgentNumber $agent

        $Worksheet.Rows.Item($insertRow).Insert()
        $Worksheet.Rows.Item($sourceRow).Copy() | Out-Null
        $Worksheet.Rows.Item($insertRow).PasteSpecial($xlPasteFormats)
        $Worksheet.Rows.Item($insertRow).PasteSpecial($xlPasteValidation)
        $Worksheet.Rows.Item($insertRow).RowHeight = $Worksheet.Rows.Item($sourceRow).RowHeight

        $Worksheet.Cells.Item($insertRow, 1).Value2 = $label
        $Worksheet.Cells.Item($insertRow, 2).ClearContents()
    }
}

function Insert-InventoryRows {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet
    )

    $xlPasteFormats = -4122
    $xlPasteValidation = 6

    foreach ($agent in 10..1) {
        $insertRow = Get-OriginalInventoryInsertRow -AgentNumber $agent
        $sourceRow = $insertRow - 1
        $label = Get-AgentLabel -AgentNumber $agent
        $baseRow = Get-InsertedBaseRow -AgentNumber $agent

        $Worksheet.Rows.Item($insertRow).Insert()
        $Worksheet.Rows.Item($sourceRow).Copy() | Out-Null
        $Worksheet.Rows.Item($insertRow).PasteSpecial($xlPasteFormats)
        $Worksheet.Rows.Item($insertRow).PasteSpecial($xlPasteValidation)
        $Worksheet.Rows.Item($insertRow).RowHeight = $Worksheet.Rows.Item($sourceRow).RowHeight

        $Worksheet.Cells.Item($insertRow, 2).Value2 = $Worksheet.Cells.Item($sourceRow, 2).Value2
        $Worksheet.Cells.Item($insertRow, 3).Value2 = "B$baseRow"
        $Worksheet.Cells.Item($insertRow, 4).Value2 = $label
        $Worksheet.Cells.Item($insertRow, 5).ClearContents()
        $Worksheet.Cells.Item($insertRow, 6).Value2 = $label
        $Worksheet.Cells.Item($insertRow, 7).Value2 = ('cf_r{0:000}' -f $baseRow)
        $Worksheet.Cells.Item($insertRow, 8).Value2 = 'Text'
        $Worksheet.Cells.Item($insertRow, 9).Value2 = 'Trim'
        $Worksheet.Cells.Item($insertRow, 10).ClearContents()
    }
}

function Update-InventoryRowNumbers {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet
    )

    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, 1).End(-4162).Row

    $insertedLabels = 1..10 | ForEach-Object { Get-AgentLabel -AgentNumber $_ }

    for ($row = 2; $row -le $lastRow; $row++) {
        $Worksheet.Cells.Item($row, 1).Value2 = $row - 1

        $sourceCellText = [string]$Worksheet.Cells.Item($row, 3).Text
        if ($sourceCellText -match '^B(\d+)$') {
            $originalBaseRow = [int]$Matches[1]
            $isInsertedRow = $insertedLabels -contains [string]$Worksheet.Cells.Item($row, 4).Text

            if (-not $isInsertedRow) {
                $adjustedBaseRow = Get-AdjustedBaseRow -OriginalRow $originalBaseRow
                $Worksheet.Cells.Item($row, 3).Value2 = "B$adjustedBaseRow"
            }
        }

        $namedRangeText = [string]$Worksheet.Cells.Item($row, 7).Text
        $currentSourceCellText = [string]$Worksheet.Cells.Item($row, 3).Text
        if ($namedRangeText -match '^cf_r\d+$' -and $currentSourceCellText -match '^B(\d+)$') {
            $currentBaseRow = [int]$Matches[1]
            $Worksheet.Cells.Item($row, 7).Value2 = ('cf_r{0:000}' -f $currentBaseRow)
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

function Find-LabelRow {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$ColumnIndex,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, $ColumnIndex).End(-4162).Row
    for ($row = 1; $row -le $lastRow; $row++) {
        if ([string]$Worksheet.Cells.Item($row, $ColumnIndex).Text -eq $Label) {
            return $row
        }
    }

    return $null
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$runtimeRoot = Split-Path -Parent $repoRoot
$basePath = (Get-ChildItem -LiteralPath $runtimeRoot -Filter '*_Base.xlsx' | Select-Object -First 1 -ExpandProperty FullName)
$kernelPath = (Get-ChildItem -LiteralPath $runtimeRoot -Filter '*_Kernel.xlsx' | Select-Object -First 1 -ExpandProperty FullName)

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

    $baseWorkbook = $excel.Workbooks.Open($basePath)
    $baseSheet = $baseWorkbook.Worksheets.Item(1)
    if (-not (Find-LabelRow -Worksheet $baseSheet -ColumnIndex 1 -Label (Get-AgentLabel -AgentNumber 1))) {
        Insert-BaseRows -Worksheet $baseSheet
        $baseWorkbook.Save()
    }
    $baseWorkbook.Close($true)
    $baseWorkbook = $null
    $baseSheet = $null

    $kernelWorkbook = $excel.Workbooks.Open($kernelPath)
    $inventorySheet = $kernelWorkbook.Worksheets.Item('CaseList_FieldInventory')
    if (-not (Find-LabelRow -Worksheet $inventorySheet -ColumnIndex 4 -Label (Get-AgentLabel -AgentNumber 1))) {
        Insert-InventoryRows -Worksheet $inventorySheet
        Update-InventoryRowNumbers -Worksheet $inventorySheet
        $kernelWorkbook.Save()
    }
    $kernelWorkbook.Close($true)
    $kernelWorkbook = $null
    $inventorySheet = $null

    [PSCustomObject]@{
        BaseBackup = $baseBackup
        KernelBackup = $kernelBackup
        AddedFields = 'Added 10 proxy position fields'
    } | Format-List
}
finally {
    if ($inventorySheet) { Release-ComObject $inventorySheet }
    if ($kernelWorkbook) { $kernelWorkbook.Close($false) | Out-Null; Release-ComObject $kernelWorkbook }
    if ($baseSheet) { Release-ComObject $baseSheet }
    if ($baseWorkbook) { $baseWorkbook.Close($false) | Out-Null; Release-ComObject $baseWorkbook }
    if ($excel) {
        $excel.CutCopyMode = 0
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
