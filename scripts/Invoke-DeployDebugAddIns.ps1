param(
    [ValidateSet('WordAddIn', 'ExcelAddIn', 'All')]
    [string]$Project = 'WordAddIn',

    [string]$MsBuildPath
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

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$resolvedMsBuildPath = Resolve-MsBuildPath -PreferredPath $MsBuildPath

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

    Write-Host "Deploying debug runtime for $target..."
    & $resolvedMsBuildPath $projectFile /restore /t:DeployDebugAddIn /p:Configuration=Debug /p:RuntimeIdentifier=win /nologo /v:m
    if ($LASTEXITCODE -ne 0) {
        throw "DeployDebugAddIn failed for $target."
    }
}
