param(
    [ValidateSet('WordAddIn', 'ExcelAddIn', 'All')]
    [string]$Project = 'All',

    [string]$MsBuildPath,

    [string]$ReleaseCertificateKeyFile,

    [string]$ReleaseCertificateThumbprint,

    [string]$ManifestCertificatePassword
)

$ErrorActionPreference = 'Stop'

function Resolve-MsBuildPath {
    param([string]$PreferredPath)

    if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) {
        if (-not (Test-Path -LiteralPath $PreferredPath)) {
            throw "MSBuild was not found: $PreferredPath"
        }

        return (Resolve-Path -LiteralPath $PreferredPath).Path
    }

    $candidates = @(
        'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw 'MSBuild.exe was not found. Install Visual Studio or pass -MsBuildPath explicitly.'
}

function Assert-PathIsOutsideRepo {
    param(
        [string]$Path,
        [string]$RepoRoot
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $resolvedRepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)

    if ($resolvedPath.StartsWith($resolvedRepoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Release signing certificate must stay outside the repository. Configured path: $resolvedPath"
    }

    return $resolvedPath
}

function Assert-PathIsNotCodexTemp {
    param(
        [string]$Path,
        [string]$Label
    )

    if ($Path -match '(?i)(?:^|[\\/])\.codex-temp(?:[\\/]|$)') {
        throw "Invalid $Label (.codex-temp detected). Refusing to build from a transient clone: $Path"
    }
}

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$resolvedMsBuildPath = Resolve-MsBuildPath -PreferredPath $MsBuildPath

Assert-PathIsNotCodexTemp -Path $repoRoot -Label 'repository root'

if ([string]::IsNullOrWhiteSpace($ReleaseCertificateKeyFile)) {
    throw 'Release build requires -ReleaseCertificateKeyFile. Use a repository-external .pfx file.'
}

if ([string]::IsNullOrWhiteSpace($ReleaseCertificateThumbprint)) {
    throw 'Release build requires -ReleaseCertificateThumbprint so the intended signing certificate is selected explicitly.'
}

$resolvedCertificatePath = $null
if (-not (Test-Path -LiteralPath $ReleaseCertificateKeyFile)) {
    throw "Release signing certificate file was not found: $ReleaseCertificateKeyFile"
}

$resolvedCertificatePath = Assert-PathIsOutsideRepo -Path $ReleaseCertificateKeyFile -RepoRoot $repoRoot
if ($resolvedCertificatePath -match '(?i)_TemporaryKey\.pfx$') {
    throw "TemporaryKey.pfx must not be used for Release signing: $resolvedCertificatePath"
}

$projectFiles = @{
    WordAddIn = Join-Path $repoRoot 'dev\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.csproj'
    ExcelAddIn = Join-Path $repoRoot 'dev\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.csproj'
}

$targets = switch ($Project) {
    'All' { @('ExcelAddIn', 'WordAddIn') }
    default { @($Project) }
}

foreach ($target in $targets) {
    $projectFile = $projectFiles[$target]
    if (-not (Test-Path -LiteralPath $projectFile)) {
        throw "Project file was not found: $projectFile"
    }

    Assert-PathIsNotCodexTemp -Path $projectFile -Label 'project path'

    $arguments = @(
        $projectFile,
        '/restore',
        '/t:DeployReleaseAddIn',
        '/p:Configuration=Release',
        '/p:RuntimeIdentifier=win',
        '/nologo',
        '/v:m'
    )

    $arguments += "/p:ReleaseCertificateKeyFile=$resolvedCertificatePath"
    $arguments += "/p:ReleaseCertificateThumbprint=$ReleaseCertificateThumbprint"

    if (-not [string]::IsNullOrWhiteSpace($ManifestCertificatePassword)) {
        $arguments += "/p:ManifestCertificatePassword=$ManifestCertificatePassword"
    }

    Write-Host "Building signed Release package for $target"
    Write-Host "  project: $projectFile"
    Write-Host "  certificate: $resolvedCertificatePath"
    Write-Host "  thumbprint: $ReleaseCertificateThumbprint"

    & $resolvedMsBuildPath @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "DeployReleaseAddIn failed for $target."
    }
}
