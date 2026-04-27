[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Write-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host ''
    Write-Host "==> $Message"
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label was not found: $Path"
    }
}

function Assert-DirectoryExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "$Label was not found: $Path"
    }
}

function Assert-PathIsUnderRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($Path)
    $resolvedRoot = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\')
    if (-not $resolvedPath.StartsWith($resolvedRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label must be under runtime root. path=$resolvedPath, runtimeRoot=$resolvedRoot"
    }
}

function Copy-DirectoryContents {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [string[]]$ExcludeNamePatterns = @()
    )

    Assert-DirectoryExists -Path $SourcePath -Label 'Source directory'
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null

    $sourceRoot = [System.IO.Path]::GetFullPath($SourcePath).TrimEnd('\')
    foreach ($item in (Get-ChildItem -LiteralPath $sourceRoot -Force -Recurse)) {
        $excluded = $false
        foreach ($pattern in $ExcludeNamePatterns) {
            if ($item.Name -like $pattern) {
                $excluded = $true
                break
            }
        }

        if ($excluded) {
            continue
        }

        $relativePath = $item.FullName.Substring($sourceRoot.Length).TrimStart('\')
        $targetPath = Join-Path $DestinationPath $relativePath

        if ($item.PSIsContainer) {
            New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
            continue
        }

        $targetParent = Split-Path -Parent $targetPath
        if (-not (Test-Path -LiteralPath $targetParent)) {
            New-Item -ItemType Directory -Path $targetParent -Force | Out-Null
        }

        Copy-Item -LiteralPath $item.FullName -Destination $targetPath -Force
    }
}

function Invoke-NormalizeDistributionWorkbookDocProps {
    param(
        [Parameter(Mandatory = $true)]
        [string]$NormalizeScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$KernelWorkbookPath,

        [Parameter(Mandatory = $true)]
        [string]$BaseWorkbookPath
    )

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $NormalizeScriptPath,
        '-KernelWorkbookPath', $KernelWorkbookPath,
        '-BaseWorkbookPath', $BaseWorkbookPath
    )

    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Docprops normalization failed with exit code $LASTEXITCODE."
    }
}

function Get-CertificateBytesFromVstoManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    Assert-FileExists -Path $ManifestPath -Label 'VSTO manifest'

    [xml]$manifestXml = Get-Content -LiteralPath $ManifestPath -Raw
    $certificateNode = $manifestXml.SelectSingleNode("//*[local-name()='X509Certificate']")
    if ($null -eq $certificateNode -or [string]::IsNullOrWhiteSpace($certificateNode.InnerText)) {
        throw "X509Certificate element was not found in VSTO manifest: $ManifestPath"
    }

    return [System.Convert]::FromBase64String($certificateNode.InnerText)
}

function Export-CertificateFromVstoManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $certificateBytes = Get-CertificateBytesFromVstoManifest -ManifestPath $ManifestPath
    [System.IO.File]::WriteAllBytes($OutputPath, $certificateBytes)
}

function New-ZipFromDirectoryWithRootName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$ZipPath,

        [Parameter(Mandatory = $true)]
        [string]$RootEntryName
    )

    $sourceRoot = [System.IO.Path]::GetFullPath($SourceDirectory).TrimEnd('\')
    $zipArchive = $null
    try {
        $zipArchive = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
        [void]$zipArchive.CreateEntry($RootEntryName.TrimEnd('/') + '/')

        Get-ChildItem -LiteralPath $sourceRoot -Force -Recurse -Directory | ForEach-Object {
            $relativePath = $_.FullName.Substring($sourceRoot.Length).TrimStart('\').Replace('\', '/')
            [void]$zipArchive.CreateEntry(('{0}/{1}/' -f $RootEntryName.TrimEnd('/'), $relativePath))
        }

        Get-ChildItem -LiteralPath $sourceRoot -Force -Recurse -File | ForEach-Object {
            $relativePath = $_.FullName.Substring($sourceRoot.Length).TrimStart('\').Replace('\', '/')
            $entryName = '{0}/{1}' -f $RootEntryName.TrimEnd('/'), $relativePath
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zipArchive,
                $_.FullName,
                $entryName,
                [System.IO.Compression.CompressionLevel]::Optimal
            ) | Out-Null
        }
    }
    finally {
        if ($null -ne $zipArchive) {
            $zipArchive.Dispose()
        }
    }
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$runtimeRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot '..'))
$distributionRoot = Join-Path $runtimeRoot '配布用'
$zipPath = Join-Path $runtimeRoot '案件情報System.zip'

$releaseLauncherExe = Join-Path $repoRoot 'dev\CaseInfoSystem.ExcelLauncher\bin\Release\CaseInfoSystem.ExcelLauncher.exe'
$releasePackageRoot = Join-Path $repoRoot 'dev\Deploy\Package'
$releaseExcelAddIn = Join-Path $releasePackageRoot 'CaseInfoSystem.ExcelAddIn'
$releaseWordAddIn = Join-Path $releasePackageRoot 'CaseInfoSystem.WordAddIn'

$kernelWorkbook = Join-Path $runtimeRoot '案件情報System_Kernel.xlsx'
$baseWorkbook = Join-Path $runtimeRoot '案件情報System_Base.xlsx'
$sourceGuidePdf = Join-Path $runtimeRoot '案件情報System_利用開始ガイド.pdf'
$sourceTemplates = Join-Path $runtimeRoot '雛形'
$normalizeScript = Join-Path $repoRoot 'scripts\Normalize-DistributionWorkbookDocProps.ps1'
$setupBatchTemplate = Join-Path $repoRoot 'distribution-assets\初回セットアップ.bat'

try {
    Write-Step 'Checking fixed paths'
    if ((Split-Path -Leaf $repoRoot) -ne '開発用') {
        throw "Repository root must be the fixed development folder: $repoRoot"
    }

    if ((Split-Path -Leaf $runtimeRoot) -ne '案件情報System') {
        throw "Runtime root must be the fixed system folder: $runtimeRoot"
    }

    Assert-PathIsUnderRoot -Path $distributionRoot -RootPath $runtimeRoot -Label 'Distribution folder'
    Assert-PathIsUnderRoot -Path $zipPath -RootPath $runtimeRoot -Label 'Distribution ZIP'
    if ((Split-Path -Leaf $distributionRoot) -ne '配布用') {
        throw "Unexpected distribution folder path: $distributionRoot"
    }

    Write-Host "Repository root : $repoRoot"
    Write-Host "Runtime root    : $runtimeRoot"
    Write-Host "Distribution    : $distributionRoot"
    Write-Host "ZIP             : $zipPath"

    Write-Step 'Checking required Release outputs'
    Assert-FileExists -Path $releaseLauncherExe -Label 'Release launcher executable'
    Assert-DirectoryExists -Path $releaseExcelAddIn -Label 'Release Excel Add-in package'
    Assert-DirectoryExists -Path $releaseWordAddIn -Label 'Release Word Add-in package'
    Assert-FileExists -Path (Join-Path $releaseExcelAddIn 'CaseInfoSystem.ExcelAddIn.dll') -Label 'Release Excel Add-in DLL'
    Assert-FileExists -Path (Join-Path $releaseExcelAddIn 'CaseInfoSystem.ExcelAddIn.vsto') -Label 'Release Excel Add-in VSTO manifest'
    Assert-FileExists -Path (Join-Path $releaseWordAddIn 'CaseInfoSystem.WordAddIn.dll') -Label 'Release Word Add-in DLL'
    Assert-FileExists -Path (Join-Path $releaseWordAddIn 'CaseInfoSystem.WordAddIn.vsto') -Label 'Release Word Add-in VSTO manifest'

    Write-Step 'Checking runtime source assets'
    Assert-FileExists -Path $kernelWorkbook -Label 'Runtime Kernel workbook'
    Assert-FileExists -Path $baseWorkbook -Label 'Runtime Base workbook'
    Assert-FileExists -Path $sourceGuidePdf -Label 'Runtime user guide PDF'
    Assert-DirectoryExists -Path $sourceTemplates -Label 'Runtime template folder'
    Assert-FileExists -Path $normalizeScript -Label 'Docprops normalization script'
    Assert-FileExists -Path $setupBatchTemplate -Label 'Initial setup batch template'

    Write-Step 'Removing existing distribution ZIP'
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }

    Write-Step 'Recreating distribution folder'
    if (Test-Path -LiteralPath $distributionRoot) {
        Remove-Item -LiteralPath $distributionRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $distributionRoot -Force | Out-Null

    Write-Step 'Copying Release build outputs'
    Copy-Item -LiteralPath $releaseLauncherExe -Destination (Join-Path $distributionRoot '案件情報System.exe') -Force
    $distributionAddins = Join-Path $distributionRoot 'Addins'
    New-Item -ItemType Directory -Path $distributionAddins -Force | Out-Null
    # Excel Add-in は現状 .config を持たないが、将来追加されてもこのコピーで拾われる前提
    Copy-DirectoryContents -SourcePath $releaseExcelAddIn -DestinationPath (Join-Path $distributionAddins 'CaseInfoSystem.ExcelAddIn')
    Copy-DirectoryContents -SourcePath $releaseWordAddIn -DestinationPath (Join-Path $distributionAddins 'CaseInfoSystem.WordAddIn')

    Write-Step 'Copying runtime source assets'
    Copy-Item -LiteralPath $kernelWorkbook -Destination (Join-Path $distributionRoot '案件情報System_Kernel.xlsx') -Force
    Copy-Item -LiteralPath $baseWorkbook -Destination (Join-Path $distributionRoot '案件情報System_Base.xlsx') -Force
    Copy-Item -LiteralPath $sourceGuidePdf -Destination (Join-Path $distributionRoot '利用開始ガイド.pdf') -Force
    Copy-Item -LiteralPath $setupBatchTemplate -Destination (Join-Path $distributionRoot '初回セットアップ.bat') -Force
    Copy-DirectoryContents -SourcePath $sourceTemplates -DestinationPath (Join-Path $distributionRoot '雛形') -ExcludeNamePatterns @('~$*')
    New-Item -ItemType Directory -Path (Join-Path $distributionRoot 'logs') -Force | Out-Null

    Write-Step 'Exporting distribution certificate from Release VSTO manifest'
    $excelCertificateBytes = Get-CertificateBytesFromVstoManifest -ManifestPath (Join-Path $releaseExcelAddIn 'CaseInfoSystem.ExcelAddIn.vsto')
    $wordCertificateBytes = Get-CertificateBytesFromVstoManifest -ManifestPath (Join-Path $releaseWordAddIn 'CaseInfoSystem.WordAddIn.vsto')
    if ([System.BitConverter]::ToString($excelCertificateBytes) -ne [System.BitConverter]::ToString($wordCertificateBytes)) {
        throw 'Release Excel/Word VSTO manifests are signed with different certificates.'
    }
    Export-CertificateFromVstoManifest `
        -ManifestPath (Join-Path $releaseExcelAddIn 'CaseInfoSystem.ExcelAddIn.vsto') `
        -OutputPath (Join-Path $distributionRoot 'CaseInfoSystem.Internal.cer')

    Write-Step 'Normalizing copied Kernel/Base docprops'
    $distributionKernel = Join-Path $distributionRoot '案件情報System_Kernel.xlsx'
    $distributionBase = Join-Path $distributionRoot '案件情報System_Base.xlsx'
    Invoke-NormalizeDistributionWorkbookDocProps -NormalizeScriptPath $normalizeScript -KernelWorkbookPath $distributionKernel -BaseWorkbookPath $distributionBase

    Write-Step 'Recreating distribution ZIP'
    Assert-DirectoryExists -Path $distributionRoot -Label 'Distribution folder before ZIP'
    New-ZipFromDirectoryWithRootName -SourceDirectory $distributionRoot -ZipPath $zipPath -RootEntryName '案件情報System'

    Write-Step 'Completed'
    Write-Host "Distribution folder: $distributionRoot"
    Write-Host "Distribution ZIP   : $zipPath"
}
catch {
    Write-Error $_
    exit 1
}
