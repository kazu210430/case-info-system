[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$KernelWorkbookPath = '',
    [string]$BaseWorkbookPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml.Linq

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$runtimeRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot '..'))

if ([string]::IsNullOrWhiteSpace($KernelWorkbookPath)) {
    $kernelCandidate = Get-ChildItem -LiteralPath $runtimeRoot -Filter '*Kernel.xlsx' | Select-Object -First 1
    if ($null -eq $kernelCandidate) {
        throw "Kernel workbook not found under $runtimeRoot"
    }

    $KernelWorkbookPath = $kernelCandidate.FullName
}

if ([string]::IsNullOrWhiteSpace($BaseWorkbookPath)) {
    $baseCandidate = Get-ChildItem -LiteralPath $runtimeRoot -Filter '*Base.xlsx' | Select-Object -First 1
    if ($null -eq $baseCandidate) {
        throw "Base workbook not found under $runtimeRoot"
    }

    $BaseWorkbookPath = $baseCandidate.FullName
}

$customPropsNamespace = [System.Xml.Linq.XNamespace]::Get('http://schemas.openxmlformats.org/officeDocument/2006/custom-properties')
$xmlNamespace = [System.Xml.Linq.XNamespace]::Xml
$snapshotChunkSize = 240
$snapshotSchemaVersion = '2'

function Resolve-WorkbookPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        throw "$Label not found: $resolvedPath"
    }

    return $resolvedPath
}

function Get-CustomPropertyElements {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$Root
    )

    return @($Root.Elements($customPropsNamespace + 'property'))
}

function Get-CustomPropertyElement {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$Root,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    foreach ($propertyElement in (Get-CustomPropertyElements -Root $Root)) {
        $nameAttribute = $propertyElement.Attribute('name')
        if ($null -ne $nameAttribute -and [string]::Equals($nameAttribute.Value, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $propertyElement
        }
    }

    throw "Custom document property not found: $Name"
}

function Get-CustomPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$PropertyElement
    )

    $valueElement = $PropertyElement.Elements() | Select-Object -First 1
    if ($null -eq $valueElement) {
        $propertyName = $PropertyElement.Attribute('name').Value
        throw "Custom document property value element missing: $propertyName"
    }

    return $valueElement.Value
}

function Set-CustomPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$PropertyElement,

        [AllowEmptyString()]
        [string]$Value
    )

    $valueElement = $PropertyElement.Elements() | Select-Object -First 1
    if ($null -eq $valueElement) {
        $propertyName = $PropertyElement.Attribute('name').Value
        throw "Custom document property value element missing: $propertyName"
    }

    $valueElement.Value = if ($null -eq $Value) { '' } else { $Value }
}

function Get-CustomPropertyMap {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$Root
    )

    $map = @{}
    foreach ($propertyElement in (Get-CustomPropertyElements -Root $Root)) {
        $nameAttribute = $propertyElement.Attribute('name')
        if ($null -eq $nameAttribute) {
            continue
        }

        $map[$nameAttribute.Value] = $propertyElement
    }

    return $map
}

function Decode-DocumentPropertyText {
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return $Value.Replace('_x000d_', "`r")
}

function Encode-DocumentPropertyText {
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return $Value.Replace("`r", '_x000d_')
}

function Get-SnapshotPartName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prefix,

        [Parameter(Mandatory = $true)]
        [int]$Index
    )

    return '{0}{1:00}' -f $Prefix, $Index
}

function Read-SnapshotText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$Root,

        [Parameter(Mandatory = $true)]
        [string]$CountPropertyName,

        [Parameter(Mandatory = $true)]
        [string]$PartPropertyPrefix
    )

    $countPropertyElement = Get-CustomPropertyElement -Root $Root -Name $CountPropertyName
    $countText = Get-CustomPropertyValue -PropertyElement $countPropertyElement
    $partCount = 0
    if (-not [int]::TryParse($countText, [ref]$partCount)) {
        throw "Custom document property is not an integer: $CountPropertyName=$countText"
    }

    $builder = New-Object System.Text.StringBuilder
    for ($index = 1; $index -le $partCount; $index++) {
        $partPropertyName = Get-SnapshotPartName -Prefix $PartPropertyPrefix -Index $index
        $partPropertyElement = Get-CustomPropertyElement -Root $Root -Name $partPropertyName
        $partText = Get-CustomPropertyValue -PropertyElement $partPropertyElement
        [void]$builder.Append((Decode-DocumentPropertyText -Value $partText))
    }

    return [pscustomobject]@{
        Count = $partCount
        Text = $builder.ToString()
    }
}

function Clear-SnapshotCache {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$Root,

        [Parameter(Mandatory = $true)]
        [string]$CountPropertyName,

        [Parameter(Mandatory = $true)]
        [string]$PartPropertyPrefix
    )

    $propertyMap = Get-CustomPropertyMap -Root $Root
    $countPropertyElement = Get-CustomPropertyElement -Root $Root -Name $CountPropertyName
    Set-CustomPropertyValue -PropertyElement $countPropertyElement -Value '0'

    foreach ($entry in $propertyMap.GetEnumerator()) {
        if (
            $entry.Key.StartsWith($PartPropertyPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and
            -not [string]::Equals($entry.Key, $CountPropertyName, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            Set-CustomPropertyValue -PropertyElement $entry.Value -Value ''
        }
    }
}

function Split-SnapshotFields {
    param(
        [AllowEmptyString()]
        [string]$Line
    )

    if ($null -eq $Line) {
        return @()
    }

    $rawFields = $Line.Split("`t")
    $fields = New-Object System.Collections.Generic.List[string]
    foreach ($rawField in $rawFields) {
        $value = $rawField
        if ($null -eq $value) {
            $value = ''
        }

        $fields.Add(
            $value.
                Replace('\n', "`n").
                Replace('\t', "`t").
                Replace('\\', '\')
        )
    }

    return $fields.ToArray()
}

function Join-SnapshotFields {
    param(
        [string[]]$Fields
    )

    if ($null -eq $Fields) {
        return ''
    }

    $escapedFields = New-Object System.Collections.Generic.List[string]
    foreach ($field in $Fields) {
        $value = if ($null -eq $field) { '' } else { $field }
        $escapedFields.Add(
            $value.
                Replace('\', '\\').
                Replace("`t", '\t').
                Replace("`r`n", '\n').
                Replace("`r", '\n').
                Replace("`n", '\n')
        )
    }

    return [string]::Join("`t", $escapedFields)
}

function Normalize-BaseSnapshotText {
    param(
        [AllowEmptyString()]
        [string]$SnapshotText
    )

    if ([string]::IsNullOrWhiteSpace($SnapshotText)) {
        throw 'TASKPANE_BASE_SNAPSHOT_* is empty.'
    }

    $lines = $SnapshotText -split "`r?`n"
    $normalizedLines = New-Object System.Collections.Generic.List[string]
    $metaUpdated = $false

    foreach ($line in $lines) {
        if (-not $metaUpdated -and $line.StartsWith('META', [System.StringComparison]::Ordinal)) {
            $fields = Split-SnapshotFields -Line $line
            if ($fields.Length -lt 6) {
                throw 'TASKPANE_BASE_SNAPSHOT META line is malformed.'
            }

            if (-not [string]::Equals($fields[1], $snapshotSchemaVersion, [System.StringComparison]::Ordinal)) {
                throw "Unsupported TASKPANE_BASE_SNAPSHOT schema version: $($fields[1])"
            }

            $fields[3] = ''
            $normalizedLines.Add((Join-SnapshotFields -Fields $fields))
            $metaUpdated = $true
            continue
        }

        $normalizedLines.Add($line)
    }

    if (-not $metaUpdated) {
        throw 'TASKPANE_BASE_SNAPSHOT META line was not found.'
    }

    return [string]::Join("`r`n", $normalizedLines)
}

function Write-SnapshotTextPreservingCount {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XElement]$Root,

        [Parameter(Mandatory = $true)]
        [string]$CountPropertyName,

        [Parameter(Mandatory = $true)]
        [string]$PartPropertyPrefix,

        [Parameter(Mandatory = $true)]
        [string]$SnapshotText
    )

    $snapshotState = Read-SnapshotText -Root $Root -CountPropertyName $CountPropertyName -PartPropertyPrefix $PartPropertyPrefix
    $partCount = $snapshotState.Count
    $normalizedSnapshotText = if ($null -eq $SnapshotText) { '' } else { $SnapshotText }
    $requiredPartCount = if ([string]::IsNullOrEmpty($normalizedSnapshotText)) { 0 } else { [int][Math]::Ceiling($normalizedSnapshotText.Length / $snapshotChunkSize) }

    if ($requiredPartCount -gt $partCount) {
        throw "$CountPropertyName cannot preserve count. required=$requiredPartCount, existing=$partCount"
    }

    $countPropertyElement = Get-CustomPropertyElement -Root $Root -Name $CountPropertyName
    Set-CustomPropertyValue -PropertyElement $countPropertyElement -Value $partCount.ToString()

    for ($index = 1; $index -le $partCount; $index++) {
        $partPropertyName = Get-SnapshotPartName -Prefix $PartPropertyPrefix -Index $index
        $partPropertyElement = Get-CustomPropertyElement -Root $Root -Name $partPropertyName
        $offset = ($index - 1) * $snapshotChunkSize
        $chunkText = ''
        if ($offset -lt $normalizedSnapshotText.Length) {
            $length = [Math]::Min($snapshotChunkSize, $normalizedSnapshotText.Length - $offset)
            $chunkText = $normalizedSnapshotText.Substring($offset, $length)
        }

        Set-CustomPropertyValue -PropertyElement $partPropertyElement -Value (Encode-DocumentPropertyText -Value $chunkText)
    }
}

function Save-CustomPropertiesXml {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$ZipArchive,

        [Parameter(Mandatory = $true)]
        [System.Xml.Linq.XDocument]$Document
    )

    $entryPath = 'docProps/custom.xml'
    $existingEntry = $ZipArchive.GetEntry($entryPath)
    if ($null -eq $existingEntry) {
        throw "$entryPath entry not found."
    }

    $existingEntry.Delete()
    $newEntry = $ZipArchive.CreateEntry($entryPath)

    $declaration = $Document.Declaration
    if ($null -eq $declaration) {
        $Document.Declaration = [System.Xml.Linq.XDeclaration]::new('1.0', 'UTF-8', 'yes')
    }
    else {
        $declaration.Version = '1.0'
        $declaration.Encoding = 'UTF-8'
        $declaration.Standalone = 'yes'
    }

    $xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
    $xmlWriterSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $xmlWriterSettings.Indent = $false
    $xmlWriterSettings.OmitXmlDeclaration = $false

    $stream = $null
    $writer = $null
    try {
        $stream = $newEntry.Open()
        $writer = [System.Xml.XmlWriter]::Create($stream, $xmlWriterSettings)
        $Document.Save($writer)
    }
    finally {
        if ($null -ne $writer) {
            $writer.Dispose()
        }
        elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Update-WorkbookCustomProperties {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Mutator
    )

    $zipArchive = $null
    $streamReader = $null
    try {
        $zipArchive = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
        $entry = $zipArchive.GetEntry('docProps/custom.xml')
        if ($null -eq $entry) {
            throw 'docProps/custom.xml entry not found.'
        }

        $streamReader = New-Object System.IO.StreamReader($entry.Open())
        $document = [System.Xml.Linq.XDocument]::Load($streamReader)
        $streamReader.Dispose()
        $streamReader = $null

        & $Mutator $document.Root
        Save-CustomPropertiesXml -ZipArchive $zipArchive -Document $document
    }
    finally {
        if ($null -ne $streamReader) {
            $streamReader.Dispose()
        }
        if ($null -ne $zipArchive) {
            $zipArchive.Dispose()
        }
    }
}

$kernelPath = Resolve-WorkbookPath -Path $KernelWorkbookPath -Label 'Kernel workbook'
$basePath = Resolve-WorkbookPath -Path $BaseWorkbookPath -Label 'Base workbook'

$kernelMutator = {
    param([System.Xml.Linq.XElement]$root)

    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'SYSTEM_ROOT') -Value ''
    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'WORD_TEMPLATE_DIR') -Value ''
    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'DEFAULT_ROOT') -Value ''
    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'LAST_PICK_FOLDER') -Value ''
    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'SUPPRESS_VSTO_HOME_ON_ACTIVATE') -Value '0'
}

$baseMutator = {
    param([System.Xml.Linq.XElement]$root)

    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'SYSTEM_ROOT') -Value ''
    Set-CustomPropertyValue -PropertyElement (Get-CustomPropertyElement -Root $root -Name 'WORD_TEMPLATE_DIR') -Value ''
    Clear-SnapshotCache -Root $root -CountPropertyName 'TASKPANE_SNAPSHOT_CACHE_COUNT' -PartPropertyPrefix 'TASKPANE_SNAPSHOT_CACHE_'

    $baseSnapshotState = Read-SnapshotText -Root $root -CountPropertyName 'TASKPANE_BASE_SNAPSHOT_COUNT' -PartPropertyPrefix 'TASKPANE_BASE_SNAPSHOT_'
    $normalizedSnapshotText = Normalize-BaseSnapshotText -SnapshotText $baseSnapshotState.Text
    Write-SnapshotTextPreservingCount -Root $root -CountPropertyName 'TASKPANE_BASE_SNAPSHOT_COUNT' -PartPropertyPrefix 'TASKPANE_BASE_SNAPSHOT_' -SnapshotText $normalizedSnapshotText
}

if ($PSCmdlet.ShouldProcess($kernelPath, 'Normalize Kernel workbook docprops for distribution')) {
    Update-WorkbookCustomProperties -WorkbookPath $kernelPath -Mutator $kernelMutator
}

if ($PSCmdlet.ShouldProcess($basePath, 'Normalize Base workbook docprops for distribution')) {
    Update-WorkbookCustomProperties -WorkbookPath $basePath -Mutator $baseMutator
}

Write-Host "Kernel workbook: $kernelPath"
Write-Host '  - Cleared SYSTEM_ROOT / WORD_TEMPLATE_DIR / DEFAULT_ROOT / LAST_PICK_FOLDER'
Write-Host '  - Reset SUPPRESS_VSTO_HOME_ON_ACTIVATE to 0'
Write-Host "Base workbook: $basePath"
Write-Host '  - Cleared SYSTEM_ROOT / WORD_TEMPLATE_DIR'
Write-Host '  - Cleared TASKPANE_SNAPSHOT_CACHE_COUNT / TASKPANE_SNAPSHOT_CACHE_*'
Write-Host '  - Normalized TASKPANE_BASE_SNAPSHOT_* META.WorkbookPath to empty while preserving schema/version/count'
