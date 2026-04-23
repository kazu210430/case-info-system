param(
    [Parameter(Mandatory = $true)]
    [string]$RuntimeManifestPath
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

function Remove-RegistryKeyIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    if (Test-Path -LiteralPath $RegistryPath) {
        Remove-Item -LiteralPath $RegistryPath -Recurse -Force
    }
}

function Remove-VstoSecurityInclusions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUrl
    )

    $root = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    if (-not (Test-Path -LiteralPath $root)) {
        return
    }

    Get-ChildItem -LiteralPath $root | ForEach-Object {
        try {
            $item = Get-ItemProperty -LiteralPath $_.PSPath
            $url = [string]$item.Url
            if ([string]::IsNullOrWhiteSpace($url)) {
                return
            }

            if ($url -like '*CaseInfoSystem.ExcelAddIn*' -and $url -ne $ExpectedUrl) {
                Remove-RegistryKeyIfPresent -RegistryPath $_.PSPath
            }
        }
        catch {
            # Ignore unreadable keys and continue with the authoritative registration.
        }
    }
}

function Get-VstoManifestPublicKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    [xml]$manifestXml = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8
    $publicKeyNode = $manifestXml.SelectSingleNode("//*[local-name()='RSAKeyValue']")
    if ($null -eq $publicKeyNode) {
        throw "VSTO manifest public key was not found: $ManifestPath"
    }

    return $publicKeyNode.OuterXml
}

function Ensure-VstoSecurityInclusion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUrl,
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
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
                if ([string]$item.PublicKey -eq $expectedPublicKey) {
                    return
                }

                Set-ItemProperty -LiteralPath $child.PSPath -Name Url -Value $ExpectedUrl
                Set-ItemProperty -LiteralPath $child.PSPath -Name PublicKey -Value $expectedPublicKey
                return
            }
        }
        catch {
            # Ignore unreadable keys and continue with authoritative registration.
        }
    }

    $childPath = Join-Path $root ([System.Guid]::NewGuid().ToString())
    New-Item -Path $childPath -Force | Out-Null
    New-ItemProperty -LiteralPath $childPath -Name Url -PropertyType String -Value $ExpectedUrl -Force | Out-Null
    New-ItemProperty -LiteralPath $childPath -Name PublicKey -PropertyType String -Value $expectedPublicKey -Force | Out-Null
}

function Remove-VstoSolutionMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUrl
    )

    $root = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
    if (-not (Test-Path -LiteralPath $root)) {
        return
    }

    $expectedUrlLower = $ExpectedUrl.ToLowerInvariant()

    Get-Item -LiteralPath $root | Get-ItemProperty | ForEach-Object {
        foreach ($property in $_.PSObject.Properties) {
            if ($property.MemberType -ne 'NoteProperty') {
                continue
            }

            $name = [string]$property.Name
            $value = [string]$property.Value
            if ($name -notlike 'file:///*CaseInfoSystem.ExcelAddIn*') {
                continue
            }

            $urlLower = $name.ToLowerInvariant()
            $guid = $value
            if ($urlLower -eq $expectedUrlLower) {
                continue
            }

            Remove-ItemProperty -LiteralPath $root -Name $name -Force
            if (-not [string]::IsNullOrWhiteSpace($guid)) {
                Remove-RegistryKeyIfPresent -RegistryPath (Join-Path $root $guid)
            }
        }
    }
}

function Remove-VstaSolutions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUrl
    )

    $root = 'HKCU:\Software\Microsoft\VSTA\Solutions'
    if (-not (Test-Path -LiteralPath $root)) {
        return
    }

    Get-ChildItem -LiteralPath $root | ForEach-Object {
        try {
            $item = Get-ItemProperty -LiteralPath $_.PSPath
            $productName = [string]$item.ProductName
            $url = [string]$item.Url
            if ($productName -eq 'CaseInfoSystem.ExcelAddIn' -and $url -ne $ExpectedUrl) {
                Remove-RegistryKeyIfPresent -RegistryPath $_.PSPath
            }
        }
        catch {
            # Ignore unreadable cache items and keep processing the remaining keys.
        }
    }
}

function Set-ComAddInRegistration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUrl
    )

    $addInPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\CaseInfoSystem.ExcelAddIn'
    if (-not (Test-Path -LiteralPath $addInPath)) {
        New-Item -Path $addInPath -Force | Out-Null
    }

    $expectedFriendlyName = 'CaseInfoSystem.ExcelAddIn'
    $expectedDescription = 'CaseInfoSystem.ExcelAddIn'
    $expectedManifest = $ExpectedUrl + '|vstolocal'
    $expectedLoadBehavior = 3

    $current = Get-ItemProperty -LiteralPath $addInPath
    if ([string]$current.FriendlyName -eq $expectedFriendlyName -and
        [string]$current.Description -eq $expectedDescription -and
        [string]$current.Manifest -eq $expectedManifest -and
        [int]$current.LoadBehavior -eq $expectedLoadBehavior) {
        return
    }

    if ([string]$current.FriendlyName -ne $expectedFriendlyName) {
        Set-ItemProperty -LiteralPath $addInPath -Name FriendlyName -Value $expectedFriendlyName
    }

    if ([string]$current.Description -ne $expectedDescription) {
        Set-ItemProperty -LiteralPath $addInPath -Name Description -Value $expectedDescription
    }

    if ([string]$current.Manifest -ne $expectedManifest) {
        Set-ItemProperty -LiteralPath $addInPath -Name Manifest -Value $expectedManifest
    }

    if ([int]$current.LoadBehavior -ne $expectedLoadBehavior) {
        New-ItemProperty -LiteralPath $addInPath -Name LoadBehavior -PropertyType DWord -Value $expectedLoadBehavior -Force | Out-Null
    }
}

$expectedUrl = Convert-ToFileUri -Path $RuntimeManifestPath

Remove-VstoSecurityInclusions -ExpectedUrl $expectedUrl
Ensure-VstoSecurityInclusion -ExpectedUrl $expectedUrl -ManifestPath $RuntimeManifestPath
Remove-VstoSolutionMetadata -ExpectedUrl $expectedUrl
Remove-VstaSolutions -ExpectedUrl $expectedUrl
Set-ComAddInRegistration -ExpectedUrl $expectedUrl
