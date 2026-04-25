param(
    [Parameter(Mandatory = $true)]
    [string]$PackageDir,

    [Parameter(Mandatory = $true)]
    [string]$RuntimeAddInDir,

    [Parameter(Mandatory = $true)]
    [string]$RuntimeManifestPath,

    [string[]]$TransientArtifacts = @()
)

$ErrorActionPreference = 'Stop'

$RuntimeAddInDir = [System.IO.Path]::GetFullPath($RuntimeAddInDir)
if ($RuntimeAddInDir -match '(?i)(?:^|[\\/])\.codex-temp(?:[\\/]|$)') {
    throw "Invalid runtime add-in directory (.codex-temp detected). Aborting because this is an incorrect execution environment and would risk VSTO misregistration: $RuntimeAddInDir"
}

function Get-ManifestVersionPair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VstoPath,
        [Parameter(Mandatory = $true)]
        [string]$ApplicationManifestPath
    )

    [xml]$vstoXml = Get-Content -LiteralPath $VstoPath -Raw -Encoding UTF8
    [xml]$appXml = Get-Content -LiteralPath $ApplicationManifestPath -Raw -Encoding UTF8

    $vstoDependency = $vstoXml.SelectSingleNode("//*[local-name()='dependency']/*[local-name()='dependentAssembly']/*[local-name()='assemblyIdentity']")
    $appIdentity = $appXml.SelectSingleNode("//*[local-name()='assemblyIdentity'][1]")

    if ($null -eq $vstoDependency) {
        throw "VSTO deployment manifest dependency identity was not found: $VstoPath"
    }

    if ($null -eq $appIdentity) {
        throw "Application manifest identity was not found: $ApplicationManifestPath"
    }

    return [pscustomobject]@{
        DeploymentReferencedName = [string]$vstoDependency.name
        DeploymentReferencedVersion = [string]$vstoDependency.version
        ApplicationManifestName = [string]$appIdentity.name
        ApplicationManifestVersion = [string]$appIdentity.version
    }
}

function Assert-ManifestVersionsMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VstoPath,
        [Parameter(Mandatory = $true)]
        [string]$ApplicationManifestPath
    )

    $pair = Get-ManifestVersionPair -VstoPath $VstoPath -ApplicationManifestPath $ApplicationManifestPath

    if ($pair.DeploymentReferencedName -ne $pair.ApplicationManifestName) {
        throw "Manifest identity name mismatch. deployment=$($pair.DeploymentReferencedName) app=$($pair.ApplicationManifestName)"
    }

    if ($pair.DeploymentReferencedVersion -ne $pair.ApplicationManifestVersion) {
        throw "Manifest version mismatch. deployment=$($pair.DeploymentReferencedVersion) app=$($pair.ApplicationManifestVersion)"
    }
}

function Restore-PackageVstoIfMissing {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageVstoPath,
        [Parameter(Mandatory = $true)]
        [string]$PackageManifestPath,
        [Parameter(Mandatory = $true)]
        [string]$RuntimeManifestPath
    )

    if (Test-Path -LiteralPath $PackageVstoPath) {
        return
    }

    if (-not (Test-Path -LiteralPath $RuntimeManifestPath)) {
        throw "Package VSTO manifest was not found and runtime fallback was unavailable: $PackageVstoPath"
    }

    Assert-ManifestVersionsMatch -VstoPath $RuntimeManifestPath -ApplicationManifestPath $PackageManifestPath
    Copy-Item -LiteralPath $RuntimeManifestPath -Destination $PackageVstoPath -Force
}

function Sync-FolderContents {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,
        [Parameter(Mandatory = $true)]
        [string]$DestinationDir
    )

    if (-not (Test-Path -LiteralPath $SourceDir)) {
        throw "Package directory was not found: $SourceDir"
    }

    if (-not (Test-Path -LiteralPath $DestinationDir)) {
        New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    }

    $sourceFileNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    Get-ChildItem -LiteralPath $SourceDir -File | ForEach-Object {
        [void]$sourceFileNames.Add($_.Name)
    }

    Get-ChildItem -LiteralPath $DestinationDir -File -Force | ForEach-Object {
        if (-not $sourceFileNames.Contains($_.Name)) {
            Remove-Item -LiteralPath $_.FullName -Force
        }
    }

    Get-ChildItem -LiteralPath $SourceDir -File | ForEach-Object {
        $destinationPath = Join-Path $DestinationDir $_.Name
        Copy-Item -LiteralPath $_.FullName -Destination $destinationPath -Force
    }
}

function Assert-FolderContentsMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,
        [Parameter(Mandatory = $true)]
        [string]$DestinationDir
    )

    $sourceFiles = Get-ChildItem -LiteralPath $SourceDir -File | Select-Object -ExpandProperty Name
    foreach ($name in $sourceFiles) {
        $destinationPath = Join-Path $DestinationDir $name
        if (-not (Test-Path -LiteralPath $destinationPath)) {
            throw "Runtime add-in sync is missing copied file: $destinationPath"
        }
    }
}

function Normalize-TransientArtifacts {
    param(
        [string[]]$Artifacts
    )

    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($artifact in ($Artifacts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        foreach ($entry in ($artifact -split ',')) {
            $trimmed = $entry.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                $normalized.Add($trimmed)
            }
        }
    }

    return $normalized.ToArray()
}

$packageVstoPath = Join-Path $PackageDir 'CaseInfoSystem.ExcelAddIn.vsto'
$packageManifestPath = Join-Path $PackageDir 'CaseInfoSystem.ExcelAddIn.dll.manifest'
$runtimeManifestFilePath = Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.dll.manifest'

if (-not (Test-Path -LiteralPath $packageManifestPath)) {
    throw "Package application manifest was not found: $packageManifestPath"
}

Restore-PackageVstoIfMissing -PackageVstoPath $packageVstoPath -PackageManifestPath $packageManifestPath -RuntimeManifestPath $RuntimeManifestPath
Assert-ManifestVersionsMatch -VstoPath $packageVstoPath -ApplicationManifestPath $packageManifestPath
Sync-FolderContents -SourceDir $PackageDir -DestinationDir $RuntimeAddInDir
Assert-FolderContentsMatch -SourceDir $PackageDir -DestinationDir $RuntimeAddInDir
Assert-ManifestVersionsMatch -VstoPath $RuntimeManifestPath -ApplicationManifestPath $runtimeManifestFilePath

$repairScriptPath = Join-Path $PSScriptRoot 'Repair-VstoRegistration.ps1'
& $repairScriptPath -RuntimeManifestPath $RuntimeManifestPath

$validatePolicyScriptPath = Join-Path $PSScriptRoot 'Validate-DocumentExecutionPolicy.ps1'
& $validatePolicyScriptPath -PolicyDirectory $RuntimeAddInDir

$validateModeScriptPath = Join-Path $PSScriptRoot 'Validate-DocumentExecutionMode.ps1'
& $validateModeScriptPath -PolicyDirectory $RuntimeAddInDir

$validatePilotScriptPath = Join-Path $PSScriptRoot 'Validate-DocumentExecutionPilot.ps1'
& $validatePilotScriptPath -PolicyDirectory $RuntimeAddInDir

foreach ($artifact in (Normalize-TransientArtifacts -Artifacts $TransientArtifacts)) {
    if (Test-Path -LiteralPath $artifact) {
        try {
            Remove-Item -LiteralPath $artifact -Force
        }
        catch {
            Write-Warning "Transient artifact cleanup skipped because the file is in use: $artifact"
        }
    }
}

Write-Output "Runtime add-in synced and validated: $RuntimeAddInDir"
