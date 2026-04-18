param(
    [string]$ProjectFile = (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'CaseInfoSystem.ExcelAddIn.csproj'),
    [string]$MsBuildPath = 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe',
    [string]$ReadinessScriptPath = (Join-Path $PSScriptRoot 'Invoke-VstoRegressionReadiness.ps1'),
    [string]$XlsxAuditScriptPath = (Join-Path $PSScriptRoot 'Invoke-XlsxSwitchAudit.ps1'),
    [string]$TrustAuditScriptPath = (Join-Path $PSScriptRoot 'Invoke-VstoTrustAudit.ps1')
)

$ErrorActionPreference = 'Stop'

function Assert-PathExists {
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

function Invoke-Build {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Configuration
    )

    $output = & $MsBuildPath $ProjectFile /t:Build /p:Configuration=$Configuration /nologo 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ($output | Out-String)
    }

    $summaryLine = $output | Where-Object { $_ -match 'ビルドに成功しました。|Build succeeded\.' } | Select-Object -Last 1
    if ($null -eq $summaryLine) {
        $summaryLine = "Build completed: $Configuration"
    }

    return [pscustomobject]@{
        Configuration = $Configuration
        Summary = [string]$summaryLine
    }
}

Assert-PathExists -Path $ProjectFile -Label 'Project file'
Assert-PathExists -Path $MsBuildPath -Label 'MSBuild'
Assert-PathExists -Path $ReadinessScriptPath -Label 'Readiness script'
Assert-PathExists -Path $XlsxAuditScriptPath -Label 'Xlsx switch audit script'
Assert-PathExists -Path $TrustAuditScriptPath -Label 'VSTO trust audit script'

$debugBuild = Invoke-Build -Configuration 'Debug'
$releaseBuild = Invoke-Build -Configuration 'Release'
$readiness = & $ReadinessScriptPath
$xlsxAudit = & $XlsxAuditScriptPath
$trustAudit = & $TrustAuditScriptPath

[pscustomobject]@{
    DebugBuild = $debugBuild.Summary
    ReleaseBuild = $releaseBuild.Summary
    ApplicationRunHits = $readiness.ApplicationRunHits
    VbaBridgeHits = $readiness.VbaBridgeHits
    DeprecatedFallbackHits = $readiness.DeprecatedFallbackHits
    PackageManifestVersion = $readiness.PackageManifestVersion
    RuntimeManifestVersion = $readiness.RuntimeManifestVersion
    PolicyValidation = $readiness.PolicyValidation
    ModeValidation = $readiness.ModeValidation
    PilotValidation = $readiness.PilotValidation
    XlsxAuditGeneratedDeployVersion = $xlsxAudit.GeneratedDeployVersion
    XlsxAuditRuntimeManifestVersion = $xlsxAudit.RuntimeManifestVersion
    XlsxAuditUnexpectedXlsmLiteralHitCount = $xlsxAudit.UnexpectedXlsmLiteralHitCount
    TrustAuditComManifestMatches = $trustAudit.ComManifestMatches
    TrustAuditComLoadBehaviorIs3 = $trustAudit.ComLoadBehaviorIs3
    TrustAuditSecurityInclusionExists = $trustAudit.SecurityInclusionExists
    TrustAuditSecurityInclusionHasPublicKey = $trustAudit.SecurityInclusionHasPublicKey
}
