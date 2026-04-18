param(
    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

$allowlistPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.txt'
$reviewNotesPath = Join-Path $PolicyDirectory 'DocumentExecutionAllowlist.review.txt'
$reviewDateFormat = 'yyyy-MM-dd'

function Read-AllowlistEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $entries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $entries
    }

    $lineNumber = 0
    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $lineNumber++
        $line = [string]$rawLine
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $line = $line.Trim()
        if ($line.StartsWith('#')) { continue }

        $parts = $line.Split('|')
        if ($parts.Length -ne 2) {
            throw "Allowlist format is invalid. path=$Path line=$lineNumber"
        }

        $key = $parts[0].Trim()
        $templateFileName = $parts[1].Trim()
        if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($templateFileName)) {
            throw "Allowlist contains empty key or templateFileName. path=$Path line=$lineNumber"
        }

        [void]$entries.Add($key + '|' + $templateFileName)
    }

    return $entries
}

function Read-PassReviewEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $entries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $statusCounts = @{}
    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            PassEntries = $entries
            StatusCounts = $statusCounts
        }
    }

    $lineNumber = 0
    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $lineNumber++
        $line = [string]$rawLine
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $line = $line.Trim()
        if ($line.StartsWith('#')) { continue }

        $parts = $line.Split('|', 6)
        if ($parts.Length -lt 6) {
            throw "Review notes format is invalid. path=$Path line=$lineNumber"
        }

        $key = $parts[0].Trim()
        $templateFileName = $parts[1].Trim()
        $status = $parts[2].Trim()
        $reviewedOn = $parts[3].Trim()
        $reviewer = $parts[4].Trim()
        $notes = $parts[5].Trim()
        if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($templateFileName)) {
            throw "Review notes contain empty key or templateFileName. path=$Path line=$lineNumber"
        }

        $identity = $key + '|' + $templateFileName
        if (-not $statusCounts.ContainsKey($identity)) {
            $statusCounts[$identity] = @{
                PASS = 0
                HOLD = 0
                FAIL = 0
                OTHER = 0
            }
        }

        if ($status -ieq 'PASS') {
            $statusCounts[$identity]['PASS']++
            [void]$entries.Add($identity)
        }
        elseif ($status -ieq 'HOLD') {
            $statusCounts[$identity]['HOLD']++
        }
        elseif ($status -ieq 'FAIL') {
            $statusCounts[$identity]['FAIL']++
        }
        else {
            $statusCounts[$identity]['OTHER']++
        }

        $parsedDate = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($reviewedOn, $reviewDateFormat, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
            throw "Review notes contain invalid reviewedOn. path=$Path line=$lineNumber expectedFormat=$reviewDateFormat"
        }

        if ([string]::IsNullOrWhiteSpace($reviewer) -or $reviewer -ieq 'reviewer') {
            throw "Review notes contain placeholder reviewer. path=$Path line=$lineNumber"
        }

        if ([string]::IsNullOrWhiteSpace($notes) -or $notes -ieq 'notes') {
            throw "Review notes contain placeholder notes. path=$Path line=$lineNumber"
        }
    }

    return [pscustomobject]@{
        PassEntries = $entries
        StatusCounts = $statusCounts
    }
}

$allowlistEntries = Read-AllowlistEntries -Path $allowlistPath
$reviewInfo = Read-PassReviewEntries -Path $reviewNotesPath
$passReviewEntries = $reviewInfo.PassEntries
$reviewStatusCounts = $reviewInfo.StatusCounts

$missingPassReviews = @()
foreach ($entry in $allowlistEntries) {
    if (-not $passReviewEntries.Contains($entry)) {
        $missingPassReviews += $entry
    }
}

if ($missingPassReviews.Count -gt 0) {
    throw ('Allowlist entries without PASS review notes were found: ' + ($missingPassReviews -join ', '))
}

$conflictingReviewEntries = @()
$duplicateReviewEntries = @()
foreach ($entry in $reviewStatusCounts.Keys) {
    $counts = $reviewStatusCounts[$entry]
    $knownKinds = 0
    if ($counts['PASS'] -gt 0) { $knownKinds++ }
    if ($counts['HOLD'] -gt 0) { $knownKinds++ }
    if ($counts['FAIL'] -gt 0) { $knownKinds++ }
    if ($knownKinds -gt 1) {
        $conflictingReviewEntries += ($entry + "(PASS=$($counts['PASS']),HOLD=$($counts['HOLD']),FAIL=$($counts['FAIL']),OTHER=$($counts['OTHER']))")
    }

    if ($counts['PASS'] -gt 1 -or $counts['HOLD'] -gt 1 -or $counts['FAIL'] -gt 1 -or $counts['OTHER'] -gt 1) {
        $duplicateReviewEntries += ($entry + "(PASS=$($counts['PASS']),HOLD=$($counts['HOLD']),FAIL=$($counts['FAIL']),OTHER=$($counts['OTHER']))")
    }
}

if ($conflictingReviewEntries.Count -gt 0) {
    throw ('Review notes contain conflicting statuses: ' + ($conflictingReviewEntries -join ', '))
}

if ($duplicateReviewEntries.Count -gt 0) {
    throw ('Review notes contain duplicate statuses: ' + ($duplicateReviewEntries -join ', '))
}

Write-Output ('Document execution policy validated. allowlistEntries=' + $allowlistEntries.Count + ', passReviewEntries=' + $passReviewEntries.Count)
