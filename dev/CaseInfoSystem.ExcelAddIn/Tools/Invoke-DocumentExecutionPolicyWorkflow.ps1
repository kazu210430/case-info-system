param(
    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [Parameter(Mandatory = $true)]
    [string]$CandidateDirectory,

    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory,

    [switch]$ApplyMerge,

    [ValidateSet('PASS', 'HOLD', 'FAIL')]
    [string]$Status = 'HOLD',

    [string]$Reviewer = '',
    [string]$Notes = '',
    [string]$ReviewedOn = '',

    [string]$ReportPath = ''
)

$ErrorActionPreference = 'Stop'

$toolRoot = $PSScriptRoot
$convertScriptPath = Join-Path $toolRoot 'Convert-DocumentEligibilityLogToPolicyCandidates.ps1'
$mergeScriptPath = Join-Path $toolRoot 'Merge-DocumentExecutionPolicyCandidates.ps1'
$statusScriptPath = Join-Path $toolRoot 'Get-DocumentExecutionPolicyStatus.ps1'
$validateScriptPath = Join-Path $toolRoot 'Validate-DocumentExecutionPolicy.ps1'

if (-not (Test-Path -LiteralPath $CandidateDirectory)) {
    New-Item -ItemType Directory -Path $CandidateDirectory -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $PolicyDirectory)) {
    New-Item -ItemType Directory -Path $PolicyDirectory -Force | Out-Null
}

& $convertScriptPath -LogPath $LogPath -OutputDirectory $CandidateDirectory

if ($ApplyMerge) {
    & $mergeScriptPath `
        -CandidateDirectory $CandidateDirectory `
        -PolicyDirectory $PolicyDirectory `
        -Status $Status `
        -Reviewer $Reviewer `
        -Notes $Notes `
        -ReviewedOn $ReviewedOn
}

$statusLines = @(& $statusScriptPath -CandidateDirectory $CandidateDirectory -PolicyDirectory $PolicyDirectory)
$statusLines | ForEach-Object { Write-Output $_ }

$validateLines = @(& $validateScriptPath -PolicyDirectory $PolicyDirectory)
$validateLines | ForEach-Object { Write-Output $_ }

if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $reportDirectory = Split-Path -Path $ReportPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($reportDirectory) -and -not (Test-Path -LiteralPath $reportDirectory)) {
        New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
    }

    $reportLines = @(
        'Document execution policy workflow report',
        "GeneratedAt=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
        "LogPath=$LogPath",
        "CandidateDirectory=$CandidateDirectory",
        "PolicyDirectory=$PolicyDirectory",
        "ApplyMerge=$($ApplyMerge.ToString())",
        ''
    ) + $statusLines + @('') + $validateLines

    Set-Content -LiteralPath $ReportPath -Value $reportLines -Encoding UTF8
    Write-Output ('Document execution policy workflow report written. path=' + $ReportPath)
}
