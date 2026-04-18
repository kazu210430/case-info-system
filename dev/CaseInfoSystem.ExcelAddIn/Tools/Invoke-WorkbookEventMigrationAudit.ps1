param(
    [string]$WorkspaceRoot = (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))))
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

function Get-WorkbookEventNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookModulePath
    )

    $pattern = 'Private Sub (Workbook_[A-Za-z0-9_]+)\('
    $matches = Select-String -Path $WorkbookModulePath -Pattern $pattern -Encoding Default
    return @($matches | ForEach-Object { $_.Matches[0].Groups[1].Value } | Sort-Object -Unique)
}

function Test-SourceContains {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,
        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    return $null -ne (Get-ChildItem -Path $RootPath -Recurse -Filter *.cs |
        Select-String -Pattern $Pattern |
        Select-Object -First 1)
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$kernelWorkbookModule = Join-Path $WorkspaceRoot 'Kemel\ThisWorkbook.cls'
$baseWorkbookModule = Join-Path $WorkspaceRoot 'Base\ThisWorkbook.cls'

Assert-PathExists -Path $projectRoot -Label 'Project root'
Assert-PathExists -Path $kernelWorkbookModule -Label 'Kernel workbook module'
Assert-PathExists -Path $baseWorkbookModule -Label 'Base workbook module'

$kernelEvents = Get-WorkbookEventNames -WorkbookModulePath $kernelWorkbookModule
$baseEvents = Get-WorkbookEventNames -WorkbookModulePath $baseWorkbookModule

[pscustomobject]@{
    KernelWorkbookEvents = ($kernelEvents -join ', ')
    BaseWorkbookEvents = ($baseEvents -join ', ')
    HasVstoWorkbookOpenHook = (Test-SourceContains -RootPath $projectRoot -Pattern 'WorkbookOpen \+=')
    HasVstoWorkbookActivateHook = (Test-SourceContains -RootPath $projectRoot -Pattern 'WorkbookActivate \+=')
    HasVstoWorkbookBeforeSaveHook = (Test-SourceContains -RootPath $projectRoot -Pattern 'WorkbookBeforeSave \+=')
    HasVstoWorkbookBeforeCloseHook = (Test-SourceContains -RootPath $projectRoot -Pattern 'WorkbookBeforeClose \+=')
    HasCaseLifecycleService = (Test-SourceContains -RootPath $projectRoot -Pattern 'class CaseWorkbookLifecycleService|sealed class CaseWorkbookLifecycleService')
    HasKernelLifecycleService = (Test-SourceContains -RootPath $projectRoot -Pattern 'class KernelWorkbookLifecycleService|sealed class KernelWorkbookLifecycleService')
}
