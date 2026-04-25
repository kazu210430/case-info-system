param(
    [Parameter(Mandatory = $true)]
    [string]$RuntimeManifestPath
)

$ErrorActionPreference = 'Stop'

$RuntimeManifestPath = [System.IO.Path]::GetFullPath($RuntimeManifestPath)
$resolvedRuntimeAddInDir = Split-Path -Parent $RuntimeManifestPath
if ($resolvedRuntimeAddInDir -match '(?i)(?:^|[\\/])\.codex-temp(?:[\\/]|$)') {
    throw "Invalid runtime add-in directory (.codex-temp detected). Aborting because this is an incorrect execution environment and would risk VSTO misregistration: $resolvedRuntimeAddInDir"
}

$CurrentAddInName = 'CaseInfoSystem.WordAddIn'
$LegacyAddInName = 'WordStyleRightAddIn'

function Convert-ToFileUri {
    param([Parameter(Mandatory = $true)][string]$Path)
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    return ([System.Uri]$fullPath).AbsoluteUri
}

function Remove-RegistryKeyIfPresent {
    param([Parameter(Mandatory = $true)][string]$RegistryPath)
    if (Test-Path -LiteralPath $RegistryPath) {
        Remove-Item -LiteralPath $RegistryPath -Recurse -Force
    }
}

function Get-VstoManifestPublicKey {
    param([Parameter(Mandatory = $true)][string]$ManifestPath)
    [xml]$manifestXml = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8
    $publicKeyNode = $manifestXml.SelectSingleNode("//*[local-name()='RSAKeyValue']")
    if ($null -eq $publicKeyNode) {
        throw "VSTO manifest public key was not found: $ManifestPath"
    }

    return $publicKeyNode.OuterXml
}

function Ensure-VstoSecurityInclusion {
    param(
        [Parameter(Mandatory = $true)][string]$ExpectedUrl,
        [Parameter(Mandatory = $true)][string]$ManifestPath
    )

    $root = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    if (-not (Test-Path -LiteralPath $root)) {
        New-Item -Path $root -Force | Out-Null
    }

    $expectedPublicKey = Get-VstoManifestPublicKey -ManifestPath $ManifestPath
    foreach ($child in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
        try {
            $item = Get-ItemProperty -LiteralPath $child.PSPath
            if ([string]$item.Url -eq $ExpectedUrl) {
                Set-ItemProperty -LiteralPath $child.PSPath -Name Url -Value $ExpectedUrl
                Set-ItemProperty -LiteralPath $child.PSPath -Name PublicKey -Value $expectedPublicKey
                return
            }
        }
        catch {
        }
    }

    $childPath = Join-Path $root ([System.Guid]::NewGuid().ToString())
    New-Item -Path $childPath -Force | Out-Null
    New-ItemProperty -LiteralPath $childPath -Name Url -PropertyType String -Value $ExpectedUrl -Force | Out-Null
    New-ItemProperty -LiteralPath $childPath -Name PublicKey -PropertyType String -Value $expectedPublicKey -Force | Out-Null
}

function Remove-WordStyleRegistrationCaches {
    param([Parameter(Mandatory = $true)][string]$ExpectedUrl)

    $expectedUrlLower = $ExpectedUrl.ToLowerInvariant()

    $solutionMetadataRoot = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
    if (Test-Path -LiteralPath $solutionMetadataRoot) {
        Get-Item -LiteralPath $solutionMetadataRoot | Get-ItemProperty | ForEach-Object {
            foreach ($property in $_.PSObject.Properties) {
                if ($property.MemberType -ne 'NoteProperty') {
                    continue
                }

                $name = [string]$property.Name
                $value = [string]$property.Value
                if ($name -notlike "file:///*$CurrentAddInName*" -and $name -notlike "file:///*$LegacyAddInName*") {
                    continue
                }

                if ($name.ToLowerInvariant() -eq $expectedUrlLower) {
                    continue
                }

                Remove-ItemProperty -LiteralPath $solutionMetadataRoot -Name $name -Force
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    Remove-RegistryKeyIfPresent -RegistryPath (Join-Path $solutionMetadataRoot $value)
                }
            }
        }
    }

    $vstaRoot = 'HKCU:\Software\Microsoft\VSTA\Solutions'
    if (Test-Path -LiteralPath $vstaRoot) {
        Get-ChildItem -LiteralPath $vstaRoot | ForEach-Object {
            try {
                $item = Get-ItemProperty -LiteralPath $_.PSPath
                if (([string]$item.ProductName -eq $CurrentAddInName -or [string]$item.ProductName -eq $LegacyAddInName) -and [string]$item.Url -ne $ExpectedUrl) {
                    Remove-RegistryKeyIfPresent -RegistryPath $_.PSPath
                }
            }
            catch {
            }
        }
    }
}

function Set-ComAddInRegistration {
    param([Parameter(Mandatory = $true)][string]$ExpectedUrl)

    $legacyAddInPath = "HKCU:\Software\Microsoft\Office\Word\Addins\$LegacyAddInName"
    Remove-RegistryKeyIfPresent -RegistryPath $legacyAddInPath

    $addInPath = "HKCU:\Software\Microsoft\Office\Word\Addins\$CurrentAddInName"
    if (-not (Test-Path -LiteralPath $addInPath)) {
        New-Item -Path $addInPath -Force | Out-Null
    }

    Set-ItemProperty -LiteralPath $addInPath -Name FriendlyName -Value $CurrentAddInName
    Set-ItemProperty -LiteralPath $addInPath -Name Description -Value $CurrentAddInName
    Set-ItemProperty -LiteralPath $addInPath -Name Manifest -Value ($ExpectedUrl + '|vstolocal')
    New-ItemProperty -LiteralPath $addInPath -Name LoadBehavior -PropertyType DWord -Value 3 -Force | Out-Null
}

$expectedUrl = Convert-ToFileUri -Path $RuntimeManifestPath
Ensure-VstoSecurityInclusion -ExpectedUrl $expectedUrl -ManifestPath $RuntimeManifestPath
Remove-WordStyleRegistrationCaches -ExpectedUrl $expectedUrl
Set-ComAddInRegistration -ExpectedUrl $expectedUrl
