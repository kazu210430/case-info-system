param(
    [string]$WorkspaceRoot = (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))))
)

$ErrorActionPreference = 'Stop'

function Get-WorkbookAuditTarget {
    param(
        [string]$Path = '',
        [Parameter(Mandatory = $true)]
        [string]$Role
    )

    [pscustomobject]@{
        Path = $Path
        Role = $Role
    }
}

function Resolve-WorkbookPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,
        [Parameter(Mandatory = $true)]
        [string]$Filter
    )

    $item = Get-ChildItem -LiteralPath $DirectoryPath -Filter $Filter -File -ErrorAction SilentlyContinue |
        Sort-Object Name |
        Select-Object -First 1

    if ($null -eq $item) {
        return ''
    }

    return [string]$item.FullName
}

function Resolve-AccountingTemplatePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceRootPath
    )

    $rootPath = [System.IO.Path]::GetFullPath($WorkspaceRootPath).TrimEnd('\')
    $item = Get-ChildItem -LiteralPath $WorkspaceRootPath -Recurse -File -Filter '*.xlsm' -ErrorAction SilentlyContinue |
        Where-Object {
            $directoryPath = [System.IO.Path]::GetFullPath($_.DirectoryName).TrimEnd('\')
            $directoryName = [System.IO.Path]::GetFileName($directoryPath)
            -not [string]::Equals($directoryPath, $rootPath, [System.StringComparison]::OrdinalIgnoreCase) -and
            -not $directoryName.StartsWith('_codex_', [System.StringComparison]::OrdinalIgnoreCase)
        } |
        Sort-Object Name |
        Select-Object -First 1

    if ($null -eq $item) {
        return ''
    }

    return [string]$item.FullName
}

function Get-ShapeSummary {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet
    )

    $formControls = 0
    $activeXControls = 0

    foreach ($shape in @($Worksheet.Shapes)) {
        try {
            if ($shape.Type -eq 8) {
                $formControls++
            }
            elseif ($shape.Type -eq 12) {
                $activeXControls++
            }
        }
        finally {
            if ($null -ne $shape) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shape)
            }
        }
    }

    [pscustomobject]@{
        SheetName = [string]$Worksheet.Name
        ShapeCount = [int]$Worksheet.Shapes.Count
        FormControlCount = $formControls
        ActiveXControlCount = $activeXControls
    }
}

function Test-HasVbaProject {
    param(
        [Parameter(Mandatory = $true)]
        $Workbook
    )

    try {
        return [bool]$Workbook.HasVBProject
    }
    catch {
        return $false
    }
}

$targets = @(
    Get-WorkbookAuditTarget -Role 'Kernel' -Path (Resolve-WorkbookPath -DirectoryPath $WorkspaceRoot -Filter '*_Kernel.xlsm')
    Get-WorkbookAuditTarget -Role 'Base' -Path (Resolve-WorkbookPath -DirectoryPath $WorkspaceRoot -Filter '*_Base.xlsm')
    Get-WorkbookAuditTarget -Role 'AccountingTemplate' -Path (Resolve-AccountingTemplatePath -WorkspaceRootPath $WorkspaceRoot)
)

$excel = $null
$results = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    foreach ($target in $targets) {
        if ([string]::IsNullOrWhiteSpace($target.Path)) {
            $results += [pscustomobject]@{
                Role = $target.Role
                Path = $target.Path
                Exists = $false
            }
            continue
        }

        if (-not (Test-Path -LiteralPath $target.Path)) {
            $results += [pscustomobject]@{
                Role = $target.Role
                Path = $target.Path
                Exists = $false
            }
            continue
        }

        $workbook = $null
        try {
            $workbook = $excel.Workbooks.Open($target.Path, 0, $true)

            $sheetSummaries = @()
            foreach ($worksheet in @($workbook.Worksheets)) {
                try {
                    $sheetSummaries += Get-ShapeSummary -Worksheet $worksheet
                }
                finally {
                    if ($null -ne $worksheet) {
                        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
                    }
                }
            }

            $hasVbaProject = Test-HasVbaProject -Workbook $workbook
            $results += [pscustomobject]@{
                Role = $target.Role
                Path = $target.Path
                Exists = $true
                FileFormat = [int]$workbook.FileFormat
                HasVBProject = $hasVbaProject
                VBComponentCount = if ($hasVbaProject) { [int]$workbook.VBProject.VBComponents.Count } else { 0 }
                WorksheetSummaries = $sheetSummaries
            }
        }
        finally {
            if ($null -ne $workbook) {
                try {
                    $workbook.Close($false)
                }
                catch {
                }

                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            }
        }
    }
}
finally {
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }

        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}

[pscustomobject]@{
    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Results = $results
}
