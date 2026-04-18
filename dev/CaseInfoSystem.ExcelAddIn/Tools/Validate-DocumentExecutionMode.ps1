param(
    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

$modePath = Join-Path $PolicyDirectory 'DocumentExecutionMode.txt'
$allowedModes = @('Disabled', 'PilotOnly', 'AllowlistedOnly')

if (-not (Test-Path -LiteralPath $modePath)) {
    throw "Document execution mode file was not found: $modePath"
}

$modeLine = $null
foreach ($rawLine in Get-Content -LiteralPath $modePath -Encoding UTF8) {
    $line = [string]$rawLine
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $trimmed = $line.Trim()
    if ($trimmed.StartsWith('#')) { continue }

    $modeLine = $trimmed
    break
}

if ([string]::IsNullOrWhiteSpace($modeLine)) {
    throw "Document execution mode file does not contain a mode value: $modePath"
}

if ($allowedModes -notcontains $modeLine) {
    throw "Document execution mode is invalid. path=$modePath value=$modeLine allowed=$($allowedModes -join ',')"
}

Write-Output ('Document execution mode validated. path=' + $modePath + ', mode=' + $modeLine)
