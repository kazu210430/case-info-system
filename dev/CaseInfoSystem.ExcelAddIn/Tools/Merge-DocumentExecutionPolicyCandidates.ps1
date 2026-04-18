param(
    [Parameter(Mandatory = $true)]
    [string]$CandidateDirectory,

    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory,

    [ValidateSet('PASS', 'HOLD', 'FAIL')]
    [string]$Status = 'HOLD',

    [string]$Reviewer = '',
    [string]$Notes = '',
    [string]$ReviewedOn = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ReviewedOn)) {
    $ReviewedOn = (Get-Date).ToString('yyyy-MM-dd')
}

if ([string]::IsNullOrWhiteSpace($Reviewer)) {
    throw 'Reviewer is required.'
}

if ([string]::IsNullOrWhiteSpace($Notes)) {
    throw 'Notes are required.'
}

$candidateAllowlistPath = Join-Path $CandidateDirectory 'DocumentExecutionAllowlist.candidates.txt'
$policyAllowlistPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.txt'
$policyReviewPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.review.txt'

if (-not (Test-Path -LiteralPath $candidateAllowlistPath)) {
    throw "Candidate allowlist file was not found: $candidateAllowlistPath"
}

if (-not (Test-Path -LiteralPath $PolicyDirectory)) {
    New-Item -ItemType Directory -Path $PolicyDirectory -Force | Out-Null
}

function Read-PolicyLines {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    return @(Get-Content -LiteralPath $Path -Encoding UTF8)
}

function Read-CandidateEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $entries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $line = [string]$rawLine
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $trimmed = $line.Trim()
        if ($trimmed.StartsWith('#')) { continue }
        $parts = $trimmed.Split('|')
        if ($parts.Length -ne 2) { continue }
        if ([string]::IsNullOrWhiteSpace($parts[0]) -or [string]::IsNullOrWhiteSpace($parts[1])) { continue }
        [void]$entries.Add($trimmed)
    }

    return $entries
}

function Split-HeaderAndEntries {
    param(
        [string[]]$Lines
    )

    $header = New-Object System.Collections.Generic.List[string]
    $entries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($line in $Lines) {
        $text = [string]$line
        if ([string]::IsNullOrWhiteSpace($text) -or $text.Trim().StartsWith('#')) {
            $header.Add($text)
            continue
        }

        [void]$entries.Add($text.Trim())
    }

    return [pscustomobject]@{
        Header = $header
        Entries = $entries
    }
}

$candidateEntries = Read-CandidateEntries -Path $candidateAllowlistPath
$allowlistInfo = Split-HeaderAndEntries -Lines (Read-PolicyLines -Path $policyAllowlistPath)
$reviewInfo = Split-HeaderAndEntries -Lines (Read-PolicyLines -Path $policyReviewPath)

$addedAllowlistEntries = New-Object System.Collections.Generic.List[string]
$addedReviewEntries = New-Object System.Collections.Generic.List[string]

foreach ($entry in ($candidateEntries | Sort-Object)) {
    if (-not $allowlistInfo.Entries.Contains($entry)) {
        [void]$allowlistInfo.Entries.Add($entry)
        $addedAllowlistEntries.Add($entry)
    }

    $reviewLine = $entry + '|' + $Status + '|' + $ReviewedOn + '|' + $Reviewer + '|' + $Notes
    if (-not $reviewInfo.Entries.Contains($reviewLine)) {
        [void]$reviewInfo.Entries.Add($reviewLine)
        $addedReviewEntries.Add($reviewLine)
    }
}

$allowlistOutput = @($allowlistInfo.Header) + @('') + @($allowlistInfo.Entries | Sort-Object)
$reviewOutput = @($reviewInfo.Header) + @('') + @($reviewInfo.Entries | Sort-Object)

Set-Content -LiteralPath $policyAllowlistPath -Value $allowlistOutput -Encoding UTF8
Set-Content -LiteralPath $policyReviewPath -Value $reviewOutput -Encoding UTF8

$message = 'Document execution policy candidates merged.' `
    + ' allowlistAdded=' + $addedAllowlistEntries.Count `
    + ', reviewAdded=' + $addedReviewEntries.Count `
    + ', policyDirectory=' + $PolicyDirectory
Write-Output $message
