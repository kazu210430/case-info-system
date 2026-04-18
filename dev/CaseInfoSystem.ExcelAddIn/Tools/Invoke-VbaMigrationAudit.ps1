param(
    [string]$WorkspaceRoot = (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))))
)

$ErrorActionPreference = 'Stop'

function Get-ModuleText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
}

function Get-ProcedureBlocks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $lines = Get-Content -LiteralPath $Path -Encoding UTF8
    $procedures = @()
    $current = $null

    for ($index = 0; $index -lt $lines.Count; $index++) {
        $line = [string]$lines[$index]
        if ($null -eq $current) {
            if ($line -match '^\s*(Private|Public)\s+Sub\s+([A-Za-z0-9_]+)\s*\(') {
                $current = [ordered]@{
                    Name = $matches[2]
                    StartLine = $index + 1
                    Lines = New-Object System.Collections.Generic.List[string]
                }
                [void]$current.Lines.Add($line)
            }

            continue
        }

        [void]$current.Lines.Add($line)
        if ($line -match '^\s*End\s+Sub\s*$') {
            $procedures += [pscustomobject]@{
                Path = $Path
                Name = $current.Name
                StartLine = $current.StartLine
                Lines = @($current.Lines)
            }
            $current = $null
        }
    }

    return $procedures
}

function Test-ProcedureActive {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Procedure
    )

    foreach ($line in $Procedure.Lines) {
        $trimmed = ([string]$line).Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.StartsWith("'")) { continue }
        if ($trimmed.StartsWith("Attribute ", [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($trimmed -match '^(Private|Public)\s+Sub\s+') { continue }
        if ($trimmed -match '^End\s+Sub$') { continue }
        if ($trimmed -match '^On\s+Error\s+Resume\s+Next$') { continue }
        return $true
    }

    return $false
}

function Get-PatternHits {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Paths,
        [Parameter(Mandatory = $true)]
        [string]$Pattern,
        [Parameter(Mandatory = $true)]
        [string]$Category
    )

    $hits = @()
    foreach ($path in $Paths) {
        $hits += Get-Content -LiteralPath $path -Encoding UTF8 |
            Select-String -Pattern $Pattern |
            ForEach-Object {
                [pscustomobject]@{
                    Category = $Category
                    Path = $_.Path
                    Line = $_.LineNumber
                    Text = $_.Line.Trim()
                }
            }
    }

    return $hits
}

$baseRoot = Join-Path $WorkspaceRoot 'Base'
$kernelRoot = Join-Path $WorkspaceRoot 'Kemel'

$allModules = @(Get-ChildItem -LiteralPath $baseRoot, $kernelRoot -File)
$modulePaths = @($allModules | Select-Object -ExpandProperty FullName)

$procedures = @()
foreach ($path in $modulePaths) {
    $procedures += Get-ProcedureBlocks -Path $path
}

$eventProcedures = $procedures | Where-Object {
    $_.Name -match '^(Workbook|Worksheet|mApp)_'
}

$activeEventProcedures = $eventProcedures | Where-Object {
    Test-ProcedureActive -Procedure $_
}

$bridgeHits = @()
$bridgeHits += Get-PatternHits -Paths $modulePaths -Pattern 'Application\.Run' -Category 'ApplicationRun'
$bridgeHits += Get-PatternHits -Paths $modulePaths -Pattern 'CallByName' -Category 'CallByName'
$bridgeHits += Get-PatternHits -Paths $modulePaths -Pattern 'COMAddIns' -Category 'ComAddInBridge'
$bridgeHits += Get-PatternHits -Paths $modulePaths -Pattern 'CreateObject\("Excel\.Application"\)' -Category 'HiddenExcelInstance'
$bridgeHits += Get-PatternHits -Paths $modulePaths -Pattern 'WithEvents\s+mApp\s+As\s+Application' -Category 'ApplicationEventMonitor'

$summary = [ordered]@{
    WorkspaceRoot = $WorkspaceRoot
    BaseRoot = $baseRoot
    KernelRoot = $kernelRoot
    TotalModuleCount = $modulePaths.Count
    EventProcedureCount = @($eventProcedures).Count
    ActiveEventProcedureCount = @($activeEventProcedures).Count
    ActiveEventProcedures = @(
        $activeEventProcedures | ForEach-Object {
            [pscustomobject]@{
                Path = $_.Path
                Name = $_.Name
                StartLine = $_.StartLine
            }
        }
    )
    ApplicationRunHitCount = @($bridgeHits | Where-Object { $_.Category -eq 'ApplicationRun' }).Count
    CallByNameHitCount = @($bridgeHits | Where-Object { $_.Category -eq 'CallByName' }).Count
    ComAddInBridgeHitCount = @($bridgeHits | Where-Object { $_.Category -eq 'ComAddInBridge' }).Count
    HiddenExcelInstanceHitCount = @($bridgeHits | Where-Object { $_.Category -eq 'HiddenExcelInstance' }).Count
    ApplicationEventMonitorHitCount = @($bridgeHits | Where-Object { $_.Category -eq 'ApplicationEventMonitor' }).Count
    BridgeHits = @($bridgeHits)
}

[pscustomobject]$summary
