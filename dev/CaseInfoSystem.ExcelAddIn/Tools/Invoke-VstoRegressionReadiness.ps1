param(
    [string]$ProjectDirectory = (Split-Path -Path $PSScriptRoot -Parent),
    [string]$WorkspaceRoot = (Split-Path -Path (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent) -Parent),
    [string]$RuntimeAddInDir = (Join-Path (Split-Path -Path (Split-Path -Path (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent) -Parent) -Parent) 'Addins\CaseInfoSystem.ExcelAddIn'),
    [string]$PackageDir = (Join-Path (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent) 'Deploy\Package\CaseInfoSystem.ExcelAddIn')
)

$ErrorActionPreference = 'Stop'

function Get-CodeHits {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,
        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    return @(Get-ChildItem -Path $RootPath -Recurse -Include *.cs |
        Select-String -Pattern $Pattern |
        ForEach-Object {
            [pscustomobject]@{
                Path = $_.Path
                Line = $_.LineNumber
                Text = $_.Line.Trim()
            }
        })
}

function Invoke-ValidationScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    try {
        $output = & $ScriptPath @Parameters 2>&1
    }
    catch {
        throw ($_ | Out-String)
    }

    return @($output | ForEach-Object { [string]$_ })
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label was not found: $Path"
    }
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

    return $pair
}

function Assert-ManagedFilesMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageDir,
        [Parameter(Mandatory = $true)]
        [string]$RuntimeDir
    )

    $managedFiles = @(
        'DocumentExecutionMode.txt'
    )

    foreach ($fileName in $managedFiles) {
        $packagePath = Join-Path $PackageDir $fileName
        $runtimePath = Join-Path $RuntimeDir $fileName
        Assert-FileExists -Path $packagePath -Label "Package managed file"
        Assert-FileExists -Path $runtimePath -Label "Runtime managed file"

        $packageText = (Get-Content -LiteralPath $packagePath -Raw -Encoding UTF8).Replace("`r`n", "`n")
        $runtimeText = (Get-Content -LiteralPath $runtimePath -Raw -Encoding UTF8).Replace("`r`n", "`n")
        if ($packageText -ne $runtimeText) {
            throw "Managed policy file mismatch. file=$fileName"
        }
    }
}

function Get-RegistryValueOrDefault {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName,
        [Parameter(Mandatory = $true)]
        [object]$DefaultValue
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        return $DefaultValue
    }

    try {
        $item = Get-ItemProperty -LiteralPath $RegistryPath
        $value = $item.PSObject.Properties[$PropertyName]
        if ($null -eq $value) {
            return $DefaultValue
        }

        return $value.Value
    }
    catch {
        return $DefaultValue
    }
}

$sourceCsRoot = $ProjectDirectory
$validateModeScript = Join-Path $PSScriptRoot 'Validate-DocumentExecutionMode.ps1'
$xlsxAuditScript = Join-Path $PSScriptRoot 'Invoke-XlsxSwitchAudit.ps1'
$trustAuditScript = Join-Path $PSScriptRoot 'Invoke-VstoTrustAudit.ps1'
$workbookEventAuditScript = Join-Path $PSScriptRoot 'Invoke-WorkbookEventMigrationAudit.ps1'
$vbaMigrationAuditScript = Join-Path $PSScriptRoot 'Invoke-VbaMigrationAudit.ps1'
$workbookFormatAuditScript = Join-Path $PSScriptRoot 'Invoke-WorkbookFormatDependencyAudit.ps1'

Assert-FileExists -Path $sourceCsRoot -Label 'Source root'
Assert-FileExists -Path $RuntimeAddInDir -Label 'Runtime add-in directory'
Assert-FileExists -Path $PackageDir -Label 'Deploy package directory'
Assert-FileExists -Path (Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.vsto') -Label 'Runtime deployment manifest'
Assert-FileExists -Path (Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.dll.manifest') -Label 'Runtime application manifest'
Assert-FileExists -Path (Join-Path $PackageDir 'CaseInfoSystem.ExcelAddIn.vsto') -Label 'Package deployment manifest'
Assert-FileExists -Path (Join-Path $PackageDir 'CaseInfoSystem.ExcelAddIn.dll.manifest') -Label 'Package application manifest'
Assert-FileExists -Path $xlsxAuditScript -Label 'Xlsx switch audit script'
Assert-FileExists -Path $trustAuditScript -Label 'VSTO trust audit script'
Assert-FileExists -Path $workbookEventAuditScript -Label 'Workbook event migration audit script'
Assert-FileExists -Path $vbaMigrationAuditScript -Label 'VBA migration audit script'
Assert-FileExists -Path $workbookFormatAuditScript -Label 'Workbook format dependency audit script'

$packageManifestPair = Assert-ManifestVersionsMatch -VstoPath (Join-Path $PackageDir 'CaseInfoSystem.ExcelAddIn.vsto') -ApplicationManifestPath (Join-Path $PackageDir 'CaseInfoSystem.ExcelAddIn.dll.manifest')
$runtimeManifestPair = Assert-ManifestVersionsMatch -VstoPath (Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.vsto') -ApplicationManifestPath (Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.dll.manifest')
Assert-ManagedFilesMatch -PackageDir $PackageDir -RuntimeDir $RuntimeAddInDir

$applicationRunHits = Get-CodeHits -RootPath $sourceCsRoot -Pattern 'Application\.Run\('
$vbaBridgeHits = Get-CodeHits -RootPath $sourceCsRoot -Pattern 'VbaBridgeService'
$deprecatedFallbackHits = Get-CodeHits -RootPath $sourceCsRoot -Pattern 'VBA fallback|vbaFallback|TryGetTaskPaneSnapshotText'

if ($applicationRunHits.Count -gt 0) {
    throw ('Application.Run references remain in add-in source: ' + (($applicationRunHits | ForEach-Object { $_.Path + ':' + $_.Line }) -join ', '))
}

if ($vbaBridgeHits.Count -gt 0) {
    throw ('VbaBridgeService references remain in add-in source: ' + (($vbaBridgeHits | ForEach-Object { $_.Path + ':' + $_.Line }) -join ', '))
}

if ($deprecatedFallbackHits.Count -gt 0) {
    throw ('Deprecated fallback terminology remains in add-in source: ' + (($deprecatedFallbackHits | ForEach-Object { $_.Path + ':' + $_.Line }) -join ', '))
}

$modeOutput = Invoke-ValidationScript -ScriptPath $validateModeScript -Parameters @{ PolicyDirectory = $RuntimeAddInDir }
$xlsxAuditOutput = & $xlsxAuditScript
$trustAuditOutput = & $trustAuditScript -RuntimeManifestPath (Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.vsto')
$workbookEventAuditOutput = & $workbookEventAuditScript
$vbaMigrationAuditOutput = & $vbaMigrationAuditScript
$workbookFormatAuditOutput = & $workbookFormatAuditScript -WorkspaceRoot $WorkspaceRoot

$kernelFormatAudit = @($workbookFormatAuditOutput.Results | Where-Object { $_.Role -eq 'Kernel' }) | Select-Object -First 1
$baseFormatAudit = @($workbookFormatAuditOutput.Results | Where-Object { $_.Role -eq 'Base' }) | Select-Object -First 1
$accountingTemplateFormatAudit = @($workbookFormatAuditOutput.Results | Where-Object { $_.Role -eq 'AccountingTemplate' }) | Select-Object -First 1

$accountingTemplateWorksheetSummaries = @()
if ($null -ne $accountingTemplateFormatAudit -and $null -ne $accountingTemplateFormatAudit.WorksheetSummaries) {
    $accountingTemplateWorksheetSummaries = @($accountingTemplateFormatAudit.WorksheetSummaries)
}

$accountingTemplateFormControlCount = @($accountingTemplateWorksheetSummaries | Measure-Object -Property FormControlCount -Sum).Sum
if ($null -eq $accountingTemplateFormControlCount) {
    $accountingTemplateFormControlCount = 0
}

$accountingTemplateActiveXControlCount = @($accountingTemplateWorksheetSummaries | Measure-Object -Property ActiveXControlCount -Sum).Sum
if ($null -eq $accountingTemplateActiveXControlCount) {
    $accountingTemplateActiveXControlCount = 0
}

$legacyKernelHomeAddInRegistryPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\AnkenInfoSystem.ExcelAddIn'
$legacyKernelHomeManifest = [string](Get-RegistryValueOrDefault -RegistryPath $legacyKernelHomeAddInRegistryPath -PropertyName 'Manifest' -DefaultValue '')
$legacyKernelHomeManifestPath = if ([string]::IsNullOrWhiteSpace($legacyKernelHomeManifest)) {
    ''
}
else {
    $normalizedManifest = $legacyKernelHomeManifest -replace '\|vstolocal$', ''
    try {
        [System.Uri]::UnescapeDataString(([System.Uri]$normalizedManifest).LocalPath)
    }
    catch {
        ''
    }
}

$summary = [ordered]@{
    ProjectDirectory = $sourceCsRoot
    WorkspaceRoot = $WorkspaceRoot
    RuntimeAddInDir = $RuntimeAddInDir
    PackageDir = $PackageDir
    ApplicationRunHits = $applicationRunHits.Count
    VbaBridgeHits = $vbaBridgeHits.Count
    DeprecatedFallbackHits = $deprecatedFallbackHits.Count
    PackageManifestVersion = $packageManifestPair.ApplicationManifestVersion
    RuntimeManifestVersion = $runtimeManifestPair.ApplicationManifestVersion
    ModeValidation = ($modeOutput -join ' / ')
    XlsxAuditGeneratedDeployVersion = $xlsxAuditOutput.GeneratedDeployVersion
    XlsxAuditRuntimeManifestVersion = $xlsxAuditOutput.RuntimeManifestVersion
    XlsxAuditUnexpectedXlsmLiteralHitCount = $xlsxAuditOutput.UnexpectedXlsmLiteralHitCount
    TrustAuditComManifestMatches = $trustAuditOutput.ComManifestMatches
    TrustAuditComLoadBehaviorIs3 = $trustAuditOutput.ComLoadBehaviorIs3
    TrustAuditSecurityInclusionExists = $trustAuditOutput.SecurityInclusionExists
    TrustAuditSecurityInclusionHasPublicKey = $trustAuditOutput.SecurityInclusionHasPublicKey
    WorkbookEventAuditKernelEvents = $workbookEventAuditOutput.KernelWorkbookEvents
    WorkbookEventAuditBaseEvents = $workbookEventAuditOutput.BaseWorkbookEvents
    WorkbookEventAuditHasVstoWorkbookOpenHook = $workbookEventAuditOutput.HasVstoWorkbookOpenHook
    WorkbookEventAuditHasVstoWorkbookBeforeSaveHook = $workbookEventAuditOutput.HasVstoWorkbookBeforeSaveHook
    WorkbookEventAuditHasVstoWorkbookBeforeCloseHook = $workbookEventAuditOutput.HasVstoWorkbookBeforeCloseHook
    WorkbookEventAuditHasCaseLifecycleService = $workbookEventAuditOutput.HasCaseLifecycleService
    WorkbookEventAuditHasKernelLifecycleService = $workbookEventAuditOutput.HasKernelLifecycleService
    VbaAuditActiveEventProcedureCount = $vbaMigrationAuditOutput.ActiveEventProcedureCount
    VbaAuditApplicationRunHitCount = $vbaMigrationAuditOutput.ApplicationRunHitCount
    VbaAuditCallByNameHitCount = $vbaMigrationAuditOutput.CallByNameHitCount
    VbaAuditComAddInBridgeHitCount = $vbaMigrationAuditOutput.ComAddInBridgeHitCount
    VbaAuditApplicationEventMonitorHitCount = $vbaMigrationAuditOutput.ApplicationEventMonitorHitCount
    WorkbookFormatKernelFileFormat = if ($null -eq $kernelFormatAudit) { 0 } else { $kernelFormatAudit.FileFormat }
    WorkbookFormatKernelHasVBProject = if ($null -eq $kernelFormatAudit) { $false } else { [bool]$kernelFormatAudit.HasVBProject }
    WorkbookFormatBaseFileFormat = if ($null -eq $baseFormatAudit) { 0 } else { $baseFormatAudit.FileFormat }
    WorkbookFormatBaseHasVBProject = if ($null -eq $baseFormatAudit) { $false } else { [bool]$baseFormatAudit.HasVBProject }
    WorkbookFormatAccountingTemplateFileFormat = if ($null -eq $accountingTemplateFormatAudit) { 0 } else { $accountingTemplateFormatAudit.FileFormat }
    WorkbookFormatAccountingTemplateHasVBProject = if ($null -eq $accountingTemplateFormatAudit) { $false } else { [bool]$accountingTemplateFormatAudit.HasVBProject }
    WorkbookFormatAccountingTemplateVBComponentCount = if ($null -eq $accountingTemplateFormatAudit) { 0 } else { $accountingTemplateFormatAudit.VBComponentCount }
    WorkbookFormatAccountingTemplateFormControlCount = $accountingTemplateFormControlCount
    WorkbookFormatAccountingTemplateActiveXControlCount = $accountingTemplateActiveXControlCount
    Phase5KernelBaseOperationallyVstoReady = (($vbaMigrationAuditOutput.ActiveEventProcedureCount -eq 0) -and ($vbaMigrationAuditOutput.ComAddInBridgeHitCount -eq 0))
    Phase5AccountingTemplateBlockedByWorkbookAssets = (
        ($null -ne $accountingTemplateFormatAudit) -and
        (
            [bool]$accountingTemplateFormatAudit.HasVBProject -or
            ($accountingTemplateFormControlCount -gt 0) -or
            ($accountingTemplateActiveXControlCount -gt 0)
        )
    )
    LegacyKernelHomeAddInRegistered = (Test-Path -LiteralPath $legacyKernelHomeAddInRegistryPath)
    LegacyKernelHomeManifest = $legacyKernelHomeManifest
    LegacyKernelHomeManifestExists = if ([string]::IsNullOrWhiteSpace($legacyKernelHomeManifestPath)) { $false } else { Test-Path -LiteralPath $legacyKernelHomeManifestPath }
}

[pscustomobject]$summary
