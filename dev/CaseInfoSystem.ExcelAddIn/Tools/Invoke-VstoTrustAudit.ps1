param(
    [string]$RuntimeManifestPath = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))) 'Addins\CaseInfoSystem.ExcelAddIn\CaseInfoSystem.ExcelAddIn.vsto')
)

$ErrorActionPreference = 'Stop'

function Convert-ToFileUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    return ([System.Uri]$fullPath).AbsoluteUri
}

function Test-VstoSecurityInclusion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUrl
    )

    $root = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    if (-not (Test-Path -LiteralPath $root)) {
        return $null
    }

    foreach ($child in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
        try {
            $item = Get-ItemProperty -LiteralPath $child.PSPath
            if ([string]$item.Url -eq $ExpectedUrl) {
                return [pscustomobject]@{
                    Key = $child.PSChildName
                    Url = [string]$item.Url
                    HasPublicKey = -not [string]::IsNullOrWhiteSpace([string]$item.PublicKey)
                }
            }
        }
        catch {
        }
    }

    return $null
}

function Get-ComAddInRegistration {
    $path = 'HKCU:\Software\Microsoft\Office\Excel\Addins\CaseInfoSystem.ExcelAddIn'
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    $item = Get-ItemProperty -LiteralPath $path
    return [pscustomobject]@{
        FriendlyName = [string]$item.FriendlyName
        Description = [string]$item.Description
        Manifest = [string]$item.Manifest
        LoadBehavior = [int]$item.LoadBehavior
    }
}

$expectedUrl = Convert-ToFileUri -Path $RuntimeManifestPath
$expectedManifestValue = $expectedUrl + '|vstolocal'
$comRegistration = Get-ComAddInRegistration
$securityInclusion = Test-VstoSecurityInclusion -ExpectedUrl $expectedUrl

[pscustomobject]@{
    RuntimeManifestPath = $RuntimeManifestPath
    ExpectedUrl = $expectedUrl
    ComRegistrationExists = $null -ne $comRegistration
    ComManifestMatches = $null -ne $comRegistration -and [string]$comRegistration.Manifest -eq $expectedManifestValue
    ComLoadBehaviorIs3 = $null -ne $comRegistration -and [int]$comRegistration.LoadBehavior -eq 3
    SecurityInclusionExists = $null -ne $securityInclusion
    SecurityInclusionHasPublicKey = $null -ne $securityInclusion -and [bool]$securityInclusion.HasPublicKey
}
