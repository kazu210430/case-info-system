[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-LabelText {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    $base = ([char]0x76F8).ToString() + [char]0x624B + $AgentNumber + [char]0x4EE3 + [char]0x7406 + [char]0x4EBA
    $suffix = ([char]0x5730).ToString() + [char]0x4F4D
    return "$base`_$suffix"
}

function Get-BaseInsertedRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    return 50 + (($AgentNumber - 1) * 21)
}

function Get-InventoryInsertedRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$AgentNumber
    )

    return 51 + (($AgentNumber - 1) * 21)
}

function Get-AdjustedBaseRow {
    param(
        [Parameter(Mandatory = $true)]
        [int]$OriginalRow
    )

    $insertedRows = @(50, 70, 90, 110, 130, 150, 170, 190, 210, 230)
    return $OriginalRow + (@($insertedRows | Where-Object { $_ -le $OriginalRow }).Count)
}

function Set-CellReference {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Cell,

        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    $reference = $Cell.GetAttribute('r')
    if ($reference -match '^([A-Z]+)\d+$') {
        $newReference = '{0}{1}' -f $Matches[1], $RowNumber
        $Cell.SetAttribute('r', $newReference)
    }
}

function Get-CellNode {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Row,

        [Parameter(Mandatory = $true)]
        [string]$ColumnLetter,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    return $Row.SelectSingleNode(("d:c[starts-with(@r,'{0}')]" -f $ColumnLetter), $NamespaceManager)
}

function Set-InlineStringCell {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Cell,

        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory = $true)]
        [string]$NamespaceUri
    )

    $Cell.SetAttribute('t', 'inlineStr')
    while ($Cell.HasChildNodes) {
        [void]$Cell.RemoveChild($Cell.FirstChild)
    }

    $inlineString = $Document.CreateElement('is', $NamespaceUri)
    $textNode = $Document.CreateElement('t', $NamespaceUri)
    $textNode.InnerText = $Text
    [void]$inlineString.AppendChild($textNode)
    [void]$Cell.AppendChild($inlineString)
}

function Set-NumericCell {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Cell,

        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory = $true)]
        [string]$NamespaceUri
    )

    if ($Cell.HasAttribute('t')) {
        [void]$Cell.RemoveAttribute('t')
    }

    while ($Cell.HasChildNodes) {
        [void]$Cell.RemoveChild($Cell.FirstChild)
    }

    $valueNode = $Document.CreateElement('v', $NamespaceUri)
    $valueNode.InnerText = $Value
    [void]$Cell.AppendChild($valueNode)
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

$repoRoot = Split-Path -Parent $PSScriptRoot
$runtimeRoot = Split-Path -Parent $repoRoot
$kernelPath = (Get-ChildItem -LiteralPath $runtimeRoot -Filter '*_Kernel.xlsx' | Select-Object -First 1 -ExpandProperty FullName)
$extractPath = Join-Path $runtimeRoot 'tmp\\kernel_xml'
$sheetPath = Join-Path $extractPath 'xl\\worksheets\\sheet4.xml'
$sharedStringsPath = Join-Path $extractPath 'xl\\sharedStrings.xml'
$workbookPath = Join-Path $extractPath 'xl\\workbook.xml'
$outputZipPath = Join-Path $runtimeRoot 'tmp\\kernel_fieldinventory_patched.zip'

if (-not (Test-Path -LiteralPath $kernelPath)) {
    throw 'Kernel workbook was not found.'
}

if (-not (Test-Path -LiteralPath $sheetPath)) {
    throw 'Expanded kernel XML was not found. Re-extract kernel_xml first.'
}

$backupPath = New-BackupFile -Path $kernelPath

$sheetDocument = New-Object System.Xml.XmlDocument
$sheetDocument.PreserveWhitespace = $true
$sheetDocument.Load($sheetPath)

$namespaceUri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$ns = New-Object System.Xml.XmlNamespaceManager($sheetDocument.NameTable)
$ns.AddNamespace('d', $namespaceUri)

$sheetData = $sheetDocument.SelectSingleNode('/d:worksheet/d:sheetData', $ns)
$dimensionNode = $sheetDocument.SelectSingleNode('/d:worksheet/d:dimension', $ns)

$originalRows = @($sheetData.SelectNodes('d:row', $ns))
$inventoryInsertedRows = 1..10 | ForEach-Object { Get-InventoryInsertedRow -AgentNumber $_ }
$inventoryOriginalInsertRows = @(51, 71, 91, 111, 131, 151, 171, 191, 211, 231)

for ($index = $originalRows.Count - 1; $index -ge 0; $index--) {
    $rowNode = [System.Xml.XmlElement]$originalRows[$index]
    $oldRowNumber = [int]$rowNode.GetAttribute('r')
    $newRowNumber = $oldRowNumber + (@($inventoryOriginalInsertRows | Where-Object { $_ -le $oldRowNumber }).Count)
    $rowNode.SetAttribute('r', [string]$newRowNumber)

    foreach ($cellNode in @($rowNode.SelectNodes('d:c', $ns))) {
        Set-CellReference -Cell $cellNode -RowNumber $newRowNumber
    }
}

foreach ($agent in 1..10) {
    $insertedRowNumber = Get-InventoryInsertedRow -AgentNumber $agent
    $baseRowNumber = Get-BaseInsertedRow -AgentNumber $agent
    $labelText = Get-LabelText -AgentNumber $agent
    $mailRowNode = [System.Xml.XmlElement]($sheetData.SelectSingleNode(("d:row[@r='{0}']" -f ($insertedRowNumber - 1)), $ns))
    $memoRowNode = [System.Xml.XmlElement]($sheetData.SelectSingleNode(("d:row[@r='{0}']" -f ($insertedRowNumber + 1)), $ns))

    if ($null -eq $mailRowNode -or $null -eq $memoRowNode) {
        throw ("Inventory insertion anchor was not found for agent {0}." -f $agent)
    }

    $newRowNode = [System.Xml.XmlElement]$mailRowNode.CloneNode($true)
    $newRowNode.SetAttribute('r', [string]$insertedRowNumber)

    foreach ($cellNode in @($newRowNode.SelectNodes('d:c', $ns))) {
        Set-CellReference -Cell $cellNode -RowNumber $insertedRowNumber
    }

    $columnACell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'A' -NamespaceManager $ns)
    $columnCCell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'C' -NamespaceManager $ns)
    $columnDCell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'D' -NamespaceManager $ns)
    $columnFCell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'F' -NamespaceManager $ns)
    $columnGCell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'G' -NamespaceManager $ns)
    $columnHCell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'H' -NamespaceManager $ns)
    $columnICell = [System.Xml.XmlElement](Get-CellNode -Row $newRowNode -ColumnLetter 'I' -NamespaceManager $ns)

    Set-NumericCell -Cell $columnACell -Value ([string]($insertedRowNumber - 1)) -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnCCell -Text ("B{0}" -f $baseRowNumber) -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnDCell -Text $labelText -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnFCell -Text $labelText -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnGCell -Text ('cf_r{0:000}' -f $baseRowNumber) -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnHCell -Text 'Text' -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnICell -Text 'Trim' -Document $sheetDocument -NamespaceUri $namespaceUri

    [void]$sheetData.InsertBefore($newRowNode, $memoRowNode)
}

$allRows = @($sheetData.SelectNodes('d:row', $ns))

foreach ($rowNode in $allRows) {
    $rowNumber = [int]$rowNode.GetAttribute('r')
    if ($rowNumber -ge 2) {
        $columnACell = [System.Xml.XmlElement](Get-CellNode -Row $rowNode -ColumnLetter 'A' -NamespaceManager $ns)
        if ($null -ne $columnACell) {
            Set-NumericCell -Cell $columnACell -Value ([string]($rowNumber - 1)) -Document $sheetDocument -NamespaceUri $namespaceUri
        }
    }

    $columnCCell = [System.Xml.XmlElement](Get-CellNode -Row $rowNode -ColumnLetter 'C' -NamespaceManager $ns)
    $columnGCell = [System.Xml.XmlElement](Get-CellNode -Row $rowNode -ColumnLetter 'G' -NamespaceManager $ns)

    if ($null -eq $columnCCell -or $null -eq $columnGCell) {
        continue
    }

    if ($inventoryInsertedRows -contains $rowNumber) {
        continue
    }

    $columnACell = [System.Xml.XmlElement](Get-CellNode -Row $rowNode -ColumnLetter 'A' -NamespaceManager $ns)
    if ($null -eq $columnACell) {
        continue
    }

    $valueNode = $columnACell.SelectSingleNode('d:v', $ns)
    if ($null -eq $valueNode) {
        continue
    }

    $originalBaseRow = [int]$valueNode.InnerText
    $adjustedBaseRow = Get-AdjustedBaseRow -OriginalRow $originalBaseRow
    Set-InlineStringCell -Cell $columnCCell -Text ('B{0}' -f $adjustedBaseRow) -Document $sheetDocument -NamespaceUri $namespaceUri
    Set-InlineStringCell -Cell $columnGCell -Text ('cf_r{0:000}' -f $adjustedBaseRow) -Document $sheetDocument -NamespaceUri $namespaceUri
}

$dimensionNode.SetAttribute('ref', 'A1:J241')
$sheetDocument.Save($sheetPath)

[xml]$workbookDocument = Get-Content -LiteralPath $workbookPath
$workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookDocument.NameTable)
$workbookNs.AddNamespace('d', $namespaceUri)
$definedNameNode = $workbookDocument.SelectSingleNode("/d:workbook/d:definedNames/d:definedName[@name='ExternalData_1' and @localSheetId='3']", $workbookNs)
if ($null -ne $definedNameNode) {
    $definedNameNode.InnerText = 'CaseList_FieldInventory!$A$1:$J$242'
}
$workbookDocument.Save($workbookPath)

if (Test-Path -LiteralPath $outputZipPath) {
    Remove-Item -LiteralPath $outputZipPath -Force
}

Compress-Archive -Path (Join-Path $extractPath '*') -DestinationPath $outputZipPath -Force
Copy-Item -LiteralPath $outputZipPath -Destination $kernelPath -Force

[PSCustomObject]@{
    BackupPath = $backupPath
    KernelPath = $kernelPath
    AddedFields = 'Added 10 proxy position inventory rows'
} | Format-List
