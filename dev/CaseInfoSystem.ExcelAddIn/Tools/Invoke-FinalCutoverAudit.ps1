param(
    [string]$ReadinessScriptPath = (Join-Path $PSScriptRoot 'Invoke-VstoRegressionReadiness.ps1')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ReadinessScriptPath)) {
    throw "Readiness script was not found: $ReadinessScriptPath"
}

$readiness = & $ReadinessScriptPath

$blockers = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]
$rollbackPoints = New-Object System.Collections.Generic.List[string]

if (-not $readiness.TrustAuditComManifestMatches) {
    $blockers.Add('VSTO COM manifest registration is not aligned with the runtime manifest.')
}

if (-not $readiness.TrustAuditSecurityInclusionExists) {
    $blockers.Add('VSTO trust inclusion is missing.')
}

if ($readiness.ApplicationRunHits -ne 0) {
    $blockers.Add('Application.Run references remain in the add-in source.')
}

if ($readiness.VbaBridgeHits -ne 0) {
    $blockers.Add('VBA bridge references remain in the add-in source.')
}

if ($readiness.VbaAuditActiveEventProcedureCount -ne 0) {
    $blockers.Add('Active VBA event procedures remain in exported workbook code.')
}

if ($readiness.VbaAuditComAddInBridgeHitCount -ne 0) {
    $blockers.Add('COM add-in bridge code remains in VBA exports.')
}

if (-not $readiness.Phase5KernelBaseOperationallyVstoReady) {
    $blockers.Add('Kernel/Base are not yet operationally VSTO-ready.')
}

if ($readiness.Phase5AccountingTemplateBlockedByWorkbookAssets) {
    $blockers.Add('Accounting workbook assets still require .xlsm-level workbook contents (VBProject and/or form controls).')
}

if ($readiness.WorkbookFormatKernelHasVBProject) {
    $warnings.Add('Kernel workbook still contains a VBProject, although active VBA event responsibility is zero.')
}

if ($readiness.WorkbookFormatBaseHasVBProject) {
    $warnings.Add('Base workbook still contains a VBProject, although active VBA event responsibility is zero.')
}

if ($readiness.WorkbookFormatAccountingTemplateVBComponentCount -gt 0) {
    $warnings.Add('Accounting template still contains VBA components: ' + $readiness.WorkbookFormatAccountingTemplateVBComponentCount)
}

if ($readiness.WorkbookFormatAccountingTemplateFormControlCount -gt 0) {
    $warnings.Add('Accounting template still contains form controls: ' + $readiness.WorkbookFormatAccountingTemplateFormControlCount)
}

$rollbackPoints.Add('Runtime add-in package can be restored from Addins\\CaseInfoSystem.ExcelAddIn if a deployment regression is found.')
$rollbackPoints.Add('Kernel/Base continue to run as .xlsm workbooks, so workbook format rollback is still available.')
$rollbackPoints.Add('Accounting workbook remains blocked from .xlsx cutover until workbook assets are fully removed or replaced.')

[pscustomobject]@{
    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    FinalCutoverReady = ($blockers.Count -eq 0)
    BlockingIssueCount = $blockers.Count
    WarningCount = $warnings.Count
    BlockingIssues = @($blockers)
    Warnings = @($warnings)
    RollbackPoints = @($rollbackPoints)
    Readiness = $readiness
}
