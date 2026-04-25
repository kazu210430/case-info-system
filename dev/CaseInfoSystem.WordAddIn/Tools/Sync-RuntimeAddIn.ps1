param(
    [Parameter(Mandatory = $true)]
    [string]$PackageDir,

    [Parameter(Mandatory = $true)]
    [string]$RuntimeAddInDir,

    [Parameter(Mandatory = $true)]
    [string]$RuntimeManifestPath
)

$ErrorActionPreference = 'Stop'

$RuntimeAddInDir = [System.IO.Path]::GetFullPath($RuntimeAddInDir)
if ($RuntimeAddInDir -match '(?i)(?:^|[\\/])\.codex-temp(?:[\\/]|$)') {
    throw "Invalid runtime add-in directory (.codex-temp detected). Aborting because this is an incorrect execution environment and would risk VSTO misregistration: $RuntimeAddInDir"
}

function Sync-FolderContents {
    param(
        [Parameter(Mandatory = $true)][string]$SourceDir,
        [Parameter(Mandatory = $true)][string]$DestinationDir
    )

    if (-not (Test-Path -LiteralPath $DestinationDir)) {
        New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    }

    Get-ChildItem -LiteralPath $DestinationDir -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Force
    }

    Get-ChildItem -LiteralPath $SourceDir -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $DestinationDir $_.Name) -Force
    }
}

Sync-FolderContents -SourceDir $PackageDir -DestinationDir $RuntimeAddInDir
$repairScriptPath = Join-Path $PSScriptRoot 'Repair-VstoRegistration.ps1'
& $repairScriptPath -RuntimeManifestPath $RuntimeManifestPath
Write-Output "Runtime add-in synced and registered: $RuntimeAddInDir"
