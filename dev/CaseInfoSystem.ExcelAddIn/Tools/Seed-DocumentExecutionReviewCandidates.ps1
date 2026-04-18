param(
    [Parameter(Mandatory = $true)]
    [string]$CandidateDirectory,

    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory,

    [ValidateSet('PASS', 'HOLD', 'FAIL')]
    [string]$Status = 'HOLD',

    [Parameter(Mandatory = $true)]
    [string]$Reviewer,

    [Parameter(Mandatory = $true)]
    [string]$Notes,

    [string]$ReviewedOn = ''
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ReviewedOn)) {
    $ReviewedOn = (Get-Date).ToString('yyyy-MM-dd')
}

$reviewCandidatePath = Join-Path $CandidateDirectory 'DocumentExecutionAllowlist.review.candidates.txt'
$policyReviewPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.review.txt'

if (-not (Test-Path -LiteralPath $reviewCandidatePath)) {
    throw "Review candidate file was not found: $reviewCandidatePath"
}

if (-not (Test-Path -LiteralPath $PolicyDirectory)) {
    New-Item -ItemType Directory -Path $PolicyDirectory -Force | Out-Null
}

function Read-PolicyLines {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    return @(Get-Content -LiteralPath $Path -Encoding UTF8)
}

function Split-HeaderAndEntries {
    param([string[]]$Lines)

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

function Read-ReviewCandidateIdentities {
    param([Parameter(Mandatory = $true)][string]$Path)

    $identities = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $line = [string]$rawLine
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $trimmed = $line.Trim()
        if ($trimmed.StartsWith('#')) { continue }

        $parts = $trimmed.Split('|', 6)
        if ($parts.Length -lt 2) { continue }

        $key = $parts[0].Trim()
        $templateFileName = $parts[1].Trim()
        if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($templateFileName)) { continue }

        [void]$identities.Add($key + '|' + $templateFileName)
    }

    return ,$identities
}

$reviewInfo = Split-HeaderAndEntries -Lines (Read-PolicyLines -Path $policyReviewPath)
$candidateIdentities = Read-ReviewCandidateIdentities -Path $reviewCandidatePath
$addedEntries = New-Object System.Collections.Generic.List[string]

foreach ($identity in ($candidateIdentities | Sort-Object)) {
    $reviewLine = $identity + '|' + $Status + '|' + $ReviewedOn + '|' + $Reviewer + '|' + $Notes
    if (-not $reviewInfo.Entries.Contains($reviewLine)) {
        [void]$reviewInfo.Entries.Add($reviewLine)
        $addedEntries.Add($reviewLine)
    }
}

$reviewOutput = @($reviewInfo.Header) + @('') + @($reviewInfo.Entries | Sort-Object)
Set-Content -LiteralPath $policyReviewPath -Value $reviewOutput -Encoding UTF8

$message = 'Document execution review candidates seeded.' `
    + ' added=' + $addedEntries.Count `
    + ', status=' + $Status `
    + ', reviewedOn=' + $ReviewedOn `
    + ', policyDirectory=' + $PolicyDirectory
Write-Output $message
