param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$RuntimeAddInDir = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))) 'Addins\CaseInfoSystem.ExcelAddIn')
)

$ErrorActionPreference = 'Stop'

function Get-XlsmLiteralHits {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $allowedPatternCategories = [ordered]@{
        'AccountingSaveAsService.cs'   = 'AccountingSaveFormat'
        'AccountingSetNamingService.cs' = 'AccountingSaveFormat'
        'AccountingTemplateResolver.cs' = 'AccountingTemplateDiscovery'
        'WorkbookRoleResolver.cs'       = 'WorkbookRoleDetection'
        'CaseWorkbookLifecycleService.cs' = 'WorkbookLifecycleDetection'
        'KernelNamingService.cs'        = 'CaseDefaultExtension'
        'KernelHomeForm.cs'             = 'CaseDefaultExtension'
        'WorkbookFileNameResolver.cs'   = 'MainWorkbookExtensionResolution'
    }

    $hits = Get-ChildItem -Path $RootPath -Recurse -Filter *.cs |
        Select-String -Pattern '"\.xlsm"' |
        ForEach-Object {
            $isAllowed = $false
            $category = ''
            foreach ($pattern in $allowedPatternCategories.Keys) {
                if ($_.Path -like "*$pattern") {
                    $isAllowed = $true
                    $category = [string]$allowedPatternCategories[$pattern]
                    break
                }
            }

            [pscustomobject]@{
                Path = $_.Path
                LineNumber = $_.LineNumber
                Line = $_.Line.Trim()
                IsAllowed = $isAllowed
                Category = $category
            }
        }

    return $hits
}

function Get-ManifestVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    if (-not (Test-Path -LiteralPath $ManifestPath)) {
        return ''
    }

    [xml]$manifestXml = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8
    $identity = $manifestXml.SelectSingleNode("//*[local-name()='assemblyIdentity'][1]")
    if ($null -eq $identity) {
        return ''
    }

    return [string]$identity.version
}

$runtimeManifestPath = Join-Path $RuntimeAddInDir 'CaseInfoSystem.ExcelAddIn.vsto'
$projectFilePath = Join-Path $ProjectRoot 'CaseInfoSystem.ExcelAddIn.csproj'
$hits = Get-XlsmLiteralHits -RootPath $ProjectRoot

$generatedDeployVersion = ''
if (Test-Path -LiteralPath $projectFilePath) {
    [xml]$projectXml = Get-Content -LiteralPath $projectFilePath -Raw -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($projectXml.NameTable)
    $ns.AddNamespace('msb', 'http://schemas.microsoft.com/developer/msbuild/2003')
    $node = $projectXml.SelectSingleNode('//msb:GeneratedDeployVersion', $ns)
    if ($node -ne $null) {
        $generatedDeployVersion = [string]$node.InnerText
        if ($null -eq $generatedDeployVersion) {
            $generatedDeployVersion = ''
        }

        $generatedDeployVersion = $generatedDeployVersion.Trim()
    }
}

[pscustomobject]@{
    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    GeneratedDeployVersion = $generatedDeployVersion
    RuntimeManifestVersion = Get-ManifestVersion -ManifestPath $runtimeManifestPath
    XlsmLiteralHitCount = @($hits).Count
    UnexpectedXlsmLiteralHitCount = @($hits | Where-Object { -not $_.IsAllowed }).Count
    AllowedXlsmLiteralCategories = @(
        $hits |
            Where-Object { $_.IsAllowed } |
            Group-Object -Property Category |
            Sort-Object Name |
            ForEach-Object {
                [pscustomobject]@{
                    Category = [string]$_.Name
                    Count = @($_.Group).Count
                    Hits = @(
                        $_.Group |
                            ForEach-Object {
                                '{0}:{1}: {2}' -f $_.Path, $_.LineNumber, $_.Line
                            }
                    )
                }
            }
    )
    UnexpectedXlsmLiteralHits = @($hits | Where-Object { -not $_.IsAllowed } | ForEach-Object {
        '{0}:{1}: {2}' -f $_.Path, $_.LineNumber, $_.Line
    })
}
