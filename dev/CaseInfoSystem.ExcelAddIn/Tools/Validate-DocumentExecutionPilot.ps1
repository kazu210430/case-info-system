param(
    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

$pilotPath = Join-Path $PolicyDirectory 'DocumentExecutionPilot.txt'
if (-not (Test-Path -LiteralPath $pilotPath)) {
    throw "Document execution pilot file was not found: $pilotPath"
}

$entries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$lineNumber = 0
foreach ($rawLine in Get-Content -LiteralPath $pilotPath -Encoding UTF8) {
    $lineNumber++
    $line = [string]$rawLine
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $trimmed = $line.Trim()
    if ($trimmed.StartsWith('#')) { continue }

    $parts = $trimmed.Split('|')
    if ($parts.Length -ne 2) {
        throw "Pilot file format is invalid. path=$pilotPath line=$lineNumber"
    }

    $key = $parts[0].Trim()
    $templateFileName = $parts[1].Trim()
    if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($templateFileName)) {
        throw "Pilot file contains empty key or templateFileName. path=$pilotPath line=$lineNumber"
    }

    [void]$entries.Add($key + '|' + $templateFileName)
}

Write-Output ('Document execution pilot validated. path=' + $pilotPath + ', entries=' + $entries.Count)
