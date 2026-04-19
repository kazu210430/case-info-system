[CmdletBinding()]
param(
    [ValidateSet('Backup', 'Verify')]
    [string]$Mode = 'Backup',

    [ValidateSet('ExcelAddIn', 'WordAddIn', 'All')]
    [string]$Project = 'All',

    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',

    [string]$BackupRoot
)

$ErrorActionPreference = 'Stop'

function Get-RepositoryRoot {
    return Split-Path -Parent $PSScriptRoot
}

function Get-RuntimeRoot {
    return Split-Path -Parent (Get-RepositoryRoot)
}

function New-ProjectDefinition {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$RuntimeDir,

        [Parameter(Mandatory = $true)]
        [string]$PackageDir,

        [Parameter(Mandatory = $true)]
        [string[]]$RequiredFiles,

        [Parameter(Mandatory = $true)]
        [object[]]$BackupItems,

        [Parameter(Mandatory = $true)]
        [string]$RuntimeManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$AddInRegistryPath
    )

    [pscustomobject]@{
        Name = $Name
        RuntimeDir = $RuntimeDir
        PackageDir = $PackageDir
        RequiredFiles = $RequiredFiles
        BackupItems = $BackupItems
        RuntimeManifestPath = $RuntimeManifestPath
        AddInRegistryPath = $AddInRegistryPath
    }
}

function Convert-ToFileUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    return ([System.Uri]$fullPath).AbsoluteUri
}

function Resolve-FirstChildFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$Filter
    )

    $item = Get-ChildItem -LiteralPath $RootPath -File -Filter $Filter -ErrorAction SilentlyContinue |
        Sort-Object Name |
        Select-Object -First 1

    if ($null -eq $item) {
        return $null
    }

    return $item.FullName
}

function Get-ExtraRuntimeDirectoryBackupItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuntimeRoot,

        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $excludedDirectoryNames = @(
        '.git',
        '.github',
        'Addins',
        'build',
        'dev',
        'docs',
        'scripts',
        'tools'
    )

    $repositoryRootFolderName = Split-Path -Leaf $RepositoryRoot
    if (-not [string]::IsNullOrWhiteSpace($repositoryRootFolderName)) {
        $excludedDirectoryNames += $repositoryRootFolderName
    }

    $items = New-Object System.Collections.Generic.List[object]
    Get-ChildItem -LiteralPath $RuntimeRoot -Directory | Where-Object {
        $excludedDirectoryNames -notcontains $_.Name
    } | ForEach-Object {
        $items.Add([pscustomobject]@{
                Label = ('Runtime directory: {0}' -f $_.Name)
                SourcePath = $_.FullName
                SnapshotRelativePath = ('ExcelAddIn\RuntimeFiles\{0}' -f $_.Name)
            })
    }

    return $items.ToArray()
}

function Get-ProjectDefinitions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$RuntimeRoot,

        [Parameter(Mandatory = $true)]
        [string]$SelectedProject,

        [Parameter(Mandatory = $true)]
        [string]$BuildConfiguration
    )

    $excelPackageDir = if ($BuildConfiguration -eq 'Release') {
        Join-Path $RepoRoot 'dev\Deploy\Package\CaseInfoSystem.ExcelAddIn'
    }
    else {
        Join-Path $RepoRoot 'dev\Deploy\DebugPackage\CaseInfoSystem.ExcelAddIn'
    }

    $wordPackageDir = if ($BuildConfiguration -eq 'Release') {
        Join-Path $RepoRoot 'dev\Deploy\Package\CaseInfoSystem.WordAddIn'
    }
    else {
        Join-Path $RepoRoot 'dev\Deploy\DebugPackage\CaseInfoSystem.WordAddIn'
    }
    $kernelWorkbookPath = Resolve-FirstChildFile -RootPath $RuntimeRoot -Filter '*Kernel.xlsx'
    $baseWorkbookPath = Resolve-FirstChildFile -RootPath $RuntimeRoot -Filter '*Base.xlsx'
    $excelExtraRuntimeItems = Get-ExtraRuntimeDirectoryBackupItems -RuntimeRoot $RuntimeRoot -RepositoryRoot $RepoRoot
    $excelBackupItems = @(
        [pscustomobject]@{
            Label = 'Excel runtime add-in'
            SourcePath = Join-Path $RuntimeRoot 'Addins\CaseInfoSystem.ExcelAddIn'
            SnapshotRelativePath = 'ExcelAddIn\Addins\CaseInfoSystem.ExcelAddIn'
        },
        [pscustomobject]@{
            Label = 'Kernel workbook'
            SourcePath = $kernelWorkbookPath
            SnapshotRelativePath = 'ExcelAddIn\RuntimeFiles\KernelWorkbook.xlsx'
        },
        [pscustomobject]@{
            Label = 'Base workbook'
            SourcePath = $baseWorkbookPath
            SnapshotRelativePath = 'ExcelAddIn\RuntimeFiles\BaseWorkbook.xlsx'
        }
    ) + $excelExtraRuntimeItems

    $excelDefinition = New-ProjectDefinition `
        -Name 'ExcelAddIn' `
        -RuntimeDir (Join-Path $RuntimeRoot 'Addins\CaseInfoSystem.ExcelAddIn') `
        -PackageDir $excelPackageDir `
        -RequiredFiles @(
            'CaseInfoSystem.ExcelAddIn.dll',
            'CaseInfoSystem.ExcelAddIn.dll.manifest',
            'CaseInfoSystem.ExcelAddIn.vsto',
            'DocumentExecutionMode.txt',
            'DocumentExecutionPilot.txt',
            'DocumentExecutionAllowlist.txt',
            'DocumentExecutionAllowlist.review.txt'
        ) `
        -BackupItems $excelBackupItems `
        -RuntimeManifestPath (Join-Path $RuntimeRoot 'Addins\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.vsto') `
        -AddInRegistryPath 'HKCU:\Software\Microsoft\Office\Excel\Addins\CaseInfoSystem.ExcelAddIn'

    $wordDefinition = New-ProjectDefinition `
        -Name 'WordAddIn' `
        -RuntimeDir (Join-Path $RuntimeRoot 'Addins\CaseInfoSystem.WordAddIn') `
        -PackageDir $wordPackageDir `
        -RequiredFiles @(
            'CaseInfoSystem.WordAddIn.dll',
            'CaseInfoSystem.WordAddIn.dll.manifest',
            'CaseInfoSystem.WordAddIn.vsto'
        ) `
        -BackupItems @(
            [pscustomobject]@{
                Label = 'Word runtime add-in'
                SourcePath = Join-Path $RuntimeRoot 'Addins\CaseInfoSystem.WordAddIn'
                SnapshotRelativePath = 'WordAddIn\Addins\CaseInfoSystem.WordAddIn'
            }
        ) `
        -RuntimeManifestPath (Join-Path $RuntimeRoot 'Addins\CaseInfoSystem.WordAddIn\CaseInfoSystem.WordAddIn.vsto') `
        -AddInRegistryPath 'HKCU:\Software\Microsoft\Office\Word\Addins\CaseInfoSystem.WordAddIn'

    if ($SelectedProject -eq 'ExcelAddIn') {
        return @($excelDefinition)
    }

    if ($SelectedProject -eq 'WordAddIn') {
        return @($wordDefinition)
    }

    return @($excelDefinition, $wordDefinition)
}

function Get-FileMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    $map = @{}
    Get-ChildItem -LiteralPath $DirectoryPath -File | ForEach-Object {
        $map[$_.Name] = $_.FullName
    }

    return $map
}

function Copy-BackupItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,

        [Parameter(Mandatory = $true)]
        [string]$SessionRoot
    )

    $sourcePath = $Item.SourcePath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        Write-Warning ("Backup skipped because the source was not found. label={0}, path={1}" -f $Item.Label, $sourcePath)
        return [pscustomobject]@{
            Label = $Item.Label
            SourcePath = $sourcePath
            SnapshotPath = $null
            Status = 'Missing'
        }
    }

    $snapshotPath = Join-Path $SessionRoot $Item.SnapshotRelativePath
    $snapshotParent = Split-Path -Parent $snapshotPath
    if (-not (Test-Path -LiteralPath $snapshotParent)) {
        New-Item -ItemType Directory -Path $snapshotParent -Force | Out-Null
    }

    Copy-Item -LiteralPath $sourcePath -Destination $snapshotPath -Recurse -Force
    Write-Output ("Backed up: {0}" -f $Item.Label)

    return [pscustomobject]@{
        Label = $Item.Label
        SourcePath = $sourcePath
        SnapshotPath = $snapshotPath
        Status = 'Copied'
    }
}

function Invoke-BackupMode {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ProjectDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedBackupRoot,

        [Parameter(Mandatory = $true)]
        [string]$BuildConfiguration
    )

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $sessionRoot = Join-Path $ResolvedBackupRoot $timestamp
    New-Item -ItemType Directory -Path $sessionRoot -Force | Out-Null

    $copiedItems = New-Object System.Collections.Generic.List[object]
    foreach ($definition in $ProjectDefinitions) {
        foreach ($item in $definition.BackupItems) {
            $copiedItems.Add((Copy-BackupItem -Item $item -SessionRoot $sessionRoot))
        }
    }

    $manifestPath = Join-Path $sessionRoot 'manifest.json'
    $manifest = [pscustomobject]@{
        CreatedAt = (Get-Date).ToString('o')
        Configuration = $BuildConfiguration
        Projects = ($ProjectDefinitions | Select-Object -ExpandProperty Name)
        Items = $copiedItems
    }
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    Write-Output ("Backup completed: {0}" -f $sessionRoot)
}

function Assert-RequiredFilesExist {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Definition,

        [Parameter(Mandatory = $true)]
        [hashtable]$PackageFiles,

        [Parameter(Mandatory = $true)]
        [hashtable]$RuntimeFiles
    )

    foreach ($name in $Definition.RequiredFiles) {
        if (-not $PackageFiles.ContainsKey($name)) {
            throw ("Package output is missing required file. project={0}, file={1}, packageDir={2}" -f $Definition.Name, $name, $Definition.PackageDir)
        }

        if (-not $RuntimeFiles.ContainsKey($name)) {
            throw ("Runtime add-in is missing required file. project={0}, file={1}, runtimeDir={2}" -f $Definition.Name, $name, $Definition.RuntimeDir)
        }
    }
}

function Assert-DirectoriesInSync {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Definition
    )

    if (-not (Test-Path -LiteralPath $Definition.PackageDir)) {
        throw ("Package directory was not found. project={0}, path={1}" -f $Definition.Name, $Definition.PackageDir)
    }

    if (-not (Test-Path -LiteralPath $Definition.RuntimeDir)) {
        throw ("Runtime add-in directory was not found. project={0}, path={1}" -f $Definition.Name, $Definition.RuntimeDir)
    }

    $packageFiles = Get-FileMap -DirectoryPath $Definition.PackageDir
    $runtimeFiles = Get-FileMap -DirectoryPath $Definition.RuntimeDir

    Assert-RequiredFilesExist -Definition $Definition -PackageFiles $packageFiles -RuntimeFiles $runtimeFiles

    foreach ($fileName in ($packageFiles.Keys | Sort-Object)) {
        if (-not $runtimeFiles.ContainsKey($fileName)) {
            throw ("Runtime add-in did not receive a copied file. project={0}, file={1}" -f $Definition.Name, $fileName)
        }

        $packageHash = (Get-FileHash -LiteralPath $packageFiles[$fileName] -Algorithm SHA256).Hash
        $runtimeHash = (Get-FileHash -LiteralPath $runtimeFiles[$fileName] -Algorithm SHA256).Hash
        if ($packageHash -ne $runtimeHash) {
            throw ("Runtime add-in file hash mismatch. project={0}, file={1}" -f $Definition.Name, $fileName)
        }
    }

    foreach ($fileName in ($runtimeFiles.Keys | Sort-Object)) {
        if (-not $packageFiles.ContainsKey($fileName)) {
            throw ("Runtime add-in contains an unexpected extra file. project={0}, file={1}" -f $Definition.Name, $fileName)
        }
    }

    Write-Output ("Verify completed: {0}" -f $Definition.Name)
}

function Assert-AddInRegistrationMatches {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Definition
    )

    if (-not (Test-Path -LiteralPath $Definition.RuntimeManifestPath)) {
        throw ("Runtime manifest was not found. project={0}, path={1}" -f $Definition.Name, $Definition.RuntimeManifestPath)
    }

    if (-not (Test-Path -LiteralPath $Definition.AddInRegistryPath)) {
        throw ("Add-in registry key was not found. project={0}, key={1}" -f $Definition.Name, $Definition.AddInRegistryPath)
    }

    $registration = Get-ItemProperty -LiteralPath $Definition.AddInRegistryPath
    $expectedManifest = (Convert-ToFileUri -Path $Definition.RuntimeManifestPath) + '|vstolocal'
    $actualManifest = [string]$registration.Manifest
    if ($actualManifest -ne $expectedManifest) {
        throw ("Add-in registry manifest mismatch. project={0}, expected={1}, actual={2}" -f $Definition.Name, $expectedManifest, $actualManifest)
    }

    $actualLoadBehavior = [int]$registration.LoadBehavior
    if ($actualLoadBehavior -ne 3) {
        throw ("Add-in LoadBehavior mismatch. project={0}, expected=3, actual={1}" -f $Definition.Name, $actualLoadBehavior)
    }
}

function Invoke-VerifyMode {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ProjectDefinitions
    )

    foreach ($definition in $ProjectDefinitions) {
        Assert-DirectoriesInSync -Definition $definition
        Assert-AddInRegistrationMatches -Definition $definition
    }
}

$repoRoot = Get-RepositoryRoot
$runtimeRoot = Get-RuntimeRoot
if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
    $BackupRoot = Join-Path $repoRoot 'build\hot-reload-backups'
}

$definitions = Get-ProjectDefinitions -RepoRoot $repoRoot -RuntimeRoot $runtimeRoot -SelectedProject $Project -BuildConfiguration $Configuration

if ($Mode -eq 'Backup') {
    Invoke-BackupMode -ProjectDefinitions $definitions -ResolvedBackupRoot $BackupRoot -BuildConfiguration $Configuration
    return
}

Invoke-VerifyMode -ProjectDefinitions $definitions
