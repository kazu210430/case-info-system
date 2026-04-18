param(
    [string]$WorkbookPath = '',
    [string]$SourceFolderPath = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $baseWorkbook = Get-ChildItem -Path $repoRoot -Filter '*Base.xlsm' | Select-Object -First 1
    if ($null -eq $baseWorkbook) {
        throw "Could not locate *Base.xlsm under $repoRoot"
    }

    $WorkbookPath = $baseWorkbook.FullName
}

if ([string]::IsNullOrWhiteSpace($SourceFolderPath)) {
    $SourceFolderPath = Join-Path $repoRoot 'Base'
}

function Get-VbaCodeBody {
    param([string]$Path)

    $lines = Get-Content -Path $Path -Encoding Default
    $bodyLines = New-Object System.Collections.Generic.List[string]
    $inBody = $false

    foreach ($line in $lines) {
        if (-not $inBody) {
            if ($line -match '^Option Explicit\b') {
                $inBody = $true
                $bodyLines.Add($line)
            }

            continue
        }

        $bodyLines.Add($line)
    }

    if ($bodyLines.Count -eq 0) {
        return ''
    }

    return ($bodyLines -join "`r`n")
}

function Remove-ComponentIfExists {
    param($Components, [string]$Name)

    try {
        $component = $Components.Item($Name)
    }
    catch {
        $component = $null
    }

    if ($null -ne $component) {
        $Components.Remove($component)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
    }
}

function Replace-DocumentModuleCode {
    param($Components, [string]$ComponentName, [string]$SourcePath)

    $component = $Components.Item($ComponentName)
    $codeModule = $component.CodeModule
    $body = Get-VbaCodeBody -Path $SourcePath

    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }

    if (-not [string]::IsNullOrWhiteSpace($body)) {
        $codeModule.AddFromString($body)
    }

    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($codeModule)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

if (-not (Test-Path -LiteralPath $SourceFolderPath)) {
    throw "Source folder not found: $SourceFolderPath"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    $components = $workbook.VBProject.VBComponents

    $standardModules = Get-ChildItem -Path $SourceFolderPath -Filter '*.bas' | Sort-Object Name
    foreach ($moduleFile in $standardModules) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($moduleFile.Name)
        Remove-ComponentIfExists -Components $components -Name $moduleName
        [void]$components.Import($moduleFile.FullName)
    }

    $documentModules = @(
        @{ Name = 'ThisWorkbook'; Source = (Join-Path $SourceFolderPath 'ThisWorkbook.cls') },
        @{ Name = 'Sheet1'; Source = (Join-Path $SourceFolderPath 'Sheet1.cls') },
        @{ Name = 'shHOME'; Source = (Join-Path $SourceFolderPath 'shHOME.cls') }
    )

    foreach ($module in $documentModules) {
        if (Test-Path -LiteralPath $module.Source) {
            Replace-DocumentModuleCode -Components $components -ComponentName $module.Name -SourcePath $module.Source
        }
    }

    $workbook.Save()
    Write-Host "Synced VBA project in $WorkbookPath from $SourceFolderPath"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
