param(
    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [string]$OutputDirectory = ''
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $LogPath)) {
    throw "Eligibility log was not found: $LogPath"
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Split-Path -Path $LogPath -Parent
}

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$allowlistOutputPath = Join-Path $OutputDirectory 'DocumentExecutionAllowlist.candidates.txt'
$reviewOutputPath = Join-Path $OutputDirectory 'DocumentExecutionAllowlist.review.candidates.txt'
$rolloutReadyOutputPath = Join-Path $OutputDirectory 'DocumentExecutionAllowlist.rollout-ready.txt'

$allowlistCandidates = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$reviewCandidates = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$rolloutReadyCandidates = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

foreach ($rawLine in Get-Content -LiteralPath $LogPath -Encoding UTF8) {
    $line = [string]$rawLine
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }

    $trimmed = $line.Trim()

    if ($trimmed.Contains('ALLOWLIST_FILE_CANDIDATE ')) {
        $value = $trimmed.Substring($trimmed.IndexOf('ALLOWLIST_FILE_CANDIDATE ', [System.StringComparison]::Ordinal) + 'ALLOWLIST_FILE_CANDIDATE '.Length).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            [void]$allowlistCandidates.Add($value)
        }
    }

    if ($trimmed.Contains('ALLOWLIST_REVIEW_CANDIDATE ')) {
        $value = $trimmed.Substring($trimmed.IndexOf('ALLOWLIST_REVIEW_CANDIDATE ', [System.StringComparison]::Ordinal) + 'ALLOWLIST_REVIEW_CANDIDATE '.Length).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            [void]$reviewCandidates.Add($value)
        }
    }

    if ($trimmed.Contains('Document eligibility PASS-reviewed candidates blocked only by allowlist. keys=')) {
        $prefix = 'Document eligibility PASS-reviewed candidates blocked only by allowlist. keys='
        $start = $trimmed.IndexOf($prefix, [System.StringComparison]::Ordinal)
        if ($start -ge 0) {
            $value = $trimmed.Substring($start + $prefix.Length)
            $allowlistPathIndex = $value.IndexOf(', allowlistPath=', [System.StringComparison]::Ordinal)
            if ($allowlistPathIndex -ge 0) {
                $value = $value.Substring(0, $allowlistPathIndex)
            }

            foreach ($entry in ($value -split ',')) {
                $normalized = $entry.Trim()
                if ([string]::IsNullOrWhiteSpace($normalized)) {
                    continue
                }

                $separatorIndex = $normalized.IndexOf(':')
                if ($separatorIndex -lt 1) {
                    continue
                }

                $rolloutReadyEntry = $normalized.Substring(0, $separatorIndex) + '|' + $normalized.Substring($separatorIndex + 1)
                [void]$rolloutReadyCandidates.Add($rolloutReadyEntry)
            }
        }
    }
}

$allowlistHeader = @(
    '# Generated from DocumentEligibilityDiagnosticsService log output',
    ('# Source log: ' + $LogPath),
    '# Copy reviewed entries into DocumentExecutionAllowlist.txt only after parity review is complete.',
    ''
)
$reviewHeader = @(
    '# Generated from DocumentEligibilityDiagnosticsService log output',
    ('# Source log: ' + $LogPath),
    '# Fill reviewer / notes / reviewedOn before copying into DocumentExecutionAllowlist.review.txt.',
    ''
)
$rolloutHeader = @(
    '# PASS-reviewed candidates blocked only by allowlist',
    ('# Source log: ' + $LogPath),
    '# These entries are already review-ready and only need allowlist registration.',
    ''
)

Set-Content -LiteralPath $allowlistOutputPath -Value ($allowlistHeader + ($allowlistCandidates | Sort-Object)) -Encoding UTF8
Set-Content -LiteralPath $reviewOutputPath -Value ($reviewHeader + ($reviewCandidates | Sort-Object)) -Encoding UTF8
Set-Content -LiteralPath $rolloutReadyOutputPath -Value ($rolloutHeader + ($rolloutReadyCandidates | Sort-Object)) -Encoding UTF8

$message = 'Document eligibility candidate files generated.' `
    + ' allowlistCandidates=' + $allowlistCandidates.Count `
    + ', reviewCandidates=' + $reviewCandidates.Count `
    + ', rolloutReadyCandidates=' + $rolloutReadyCandidates.Count `
    + ', outputDirectory=' + $OutputDirectory
Write-Output $message
