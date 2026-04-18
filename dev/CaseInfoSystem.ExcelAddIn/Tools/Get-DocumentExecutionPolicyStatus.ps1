param(
    [Parameter(Mandatory = $true)]
    [string]$CandidateDirectory,

    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

function Read-Entries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$ExpectedColumnCount = 0
    )

    $entries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path -LiteralPath $Path)) {
        return ,$entries
    }

    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $line = [string]$rawLine
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $trimmed = $line.Trim()
        if ($trimmed.StartsWith('#')) { continue }
        if ($ExpectedColumnCount -gt 0) {
            $parts = $trimmed.Split('|')
            if ($parts.Length -lt $ExpectedColumnCount) { continue }
        }

        [void]$entries.Add($trimmed)
    }

    return ,$entries
}

function Convert-ReviewEntriesToPassIdentities {
    param(
        $ReviewEntries
    )

    $identities = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -eq $ReviewEntries) {
        return ,$identities
    }

    foreach ($entry in $ReviewEntries) {
        $parts = ([string]$entry).Split('|', 6)
        if ($parts.Length -lt 6) { continue }
        if ($parts[2].Trim() -ieq 'PASS') {
            [void]$identities.Add($parts[0].Trim() + '|' + $parts[1].Trim())
        }
    }

    return ,$identities
}

function Get-ReviewStatusCounts {
    param(
        $ReviewEntries
    )

    $counts = [ordered]@{
        PASS = 0
        HOLD = 0
        FAIL = 0
        OTHER = 0
    }

    if ($null -eq $ReviewEntries) {
        return $counts
    }

    foreach ($entry in $ReviewEntries) {
        $parts = ([string]$entry).Split('|', 6)
        if ($parts.Length -lt 6) { continue }

        $status = ([string]$parts[2]).Trim().ToUpperInvariant()
        if ($counts.Contains($status)) {
            $counts[$status]++
        }
        else {
            $counts['OTHER']++
        }
    }

    return $counts
}

$candidateAllowlistPath = Join-Path $CandidateDirectory 'DocumentExecutionAllowlist.candidates.txt'
$policyAllowlistPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.txt'
$policyReviewPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.review.txt'

$candidateEntries = Read-Entries -Path $candidateAllowlistPath -ExpectedColumnCount 2
$allowlistEntries = Read-Entries -Path $policyAllowlistPath -ExpectedColumnCount 2
$reviewEntries = Read-Entries -Path $policyReviewPath -ExpectedColumnCount 6
$passReviewIdentities = Convert-ReviewEntriesToPassIdentities -ReviewEntries $reviewEntries
$reviewStatusCounts = Get-ReviewStatusCounts -ReviewEntries $reviewEntries

$candidateNotAllowlisted = @()
$allowlistedWithoutPass = @()
$rolloutReady = @()

foreach ($entry in $candidateEntries) {
    if (-not $allowlistEntries.Contains($entry)) {
        $candidateNotAllowlisted += $entry
        continue
    }

    if ($passReviewIdentities.Contains($entry)) {
        $rolloutReady += $entry
    }
}

foreach ($entry in $allowlistEntries) {
    if (-not $passReviewIdentities.Contains($entry)) {
        $allowlistedWithoutPass += $entry
    }
}

$message = 'Document execution policy status.' `
    + ' candidateEntries=' + $candidateEntries.Count `
    + ', allowlistEntries=' + $allowlistEntries.Count `
    + ', reviewEntries=' + $reviewEntries.Count `
    + ', passReviewEntries=' + $passReviewIdentities.Count `
    + ', holdReviewEntries=' + $reviewStatusCounts['HOLD'] `
    + ', failReviewEntries=' + $reviewStatusCounts['FAIL'] `
    + ', otherReviewEntries=' + $reviewStatusCounts['OTHER'] `
    + ', rolloutReady=' + $rolloutReady.Count `
    + ', candidateNotAllowlisted=' + $candidateNotAllowlisted.Count `
    + ', allowlistedWithoutPass=' + $allowlistedWithoutPass.Count
Write-Output $message

if ($candidateNotAllowlisted.Count -gt 0) {
    Write-Output ('PENDING_ALLOWLIST ' + (($candidateNotAllowlisted | Sort-Object) -join ','))
}

if ($allowlistedWithoutPass.Count -gt 0) {
    Write-Output ('PENDING_PASS_REVIEW ' + (($allowlistedWithoutPass | Sort-Object) -join ','))
}

if ($rolloutReady.Count -gt 0) {
    Write-Output ('ROLLOUT_READY ' + (($rolloutReady | Sort-Object) -join ','))
}
