param(
    [ValidateSet('Test', 'Compile', 'DeployDebugAddIn', 'DeployReleaseAddIn', 'Help')]
    [string]$Mode = 'Help',

    [ValidateSet('ExcelAddIn', 'WordAddIn', 'All')]
    [string]$Project = 'All',

    [string]$Configuration,

    [string]$MsBuildPath,

    [string]$ReleaseCertificateKeyFile,

    [string]$ReleaseCertificateThumbprint,

    [string]$ManifestCertificatePassword,

    [switch]$NoBuild,

    [switch]$NoRestore
)

$ErrorActionPreference = 'Stop'

function Invoke-ExternalCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw ("Command failed with exit code {0}: {1} {2}" -f $LASTEXITCODE, $FilePath, ($Arguments -join ' '))
    }
}

function Show-Help {
    Write-Host 'Usage:'
    Write-Host '  .\build.ps1 -Mode <Test|Compile|DeployDebugAddIn|DeployReleaseAddIn|Help>'
    Write-Host ''
    Write-Host 'Available modes:'
    Write-Host '  Test'
    Write-Host '    Standard test entrypoint. Runs dotnet test for dev\CaseInfoSystem.slnx.'
    Write-Host '  Compile'
    Write-Host '    CI-equivalent safe build check. This is compile-only and does not deploy add-ins to the runtime Addins directory.'
    Write-Host '  DeployDebugAddIn'
    Write-Host '    Wraps scripts\Invoke-DeployDebugAddIns.ps1 and reflects Debug add-ins into the runtime Addins directory.'
    Write-Host '  DeployReleaseAddIn'
    Write-Host '    Wraps scripts\Invoke-DeployReleaseAddIns.ps1 and builds signed Release VSTO packages from dev\*.csproj only.'
    Write-Host '  Help'
    Write-Host '    Shows this help without running build, test, or deploy work.'
    Write-Host ''
    Write-Host 'Examples:'
    Write-Host '  .\build.ps1'
    Write-Host '  .\build.ps1 -Mode Test'
    Write-Host '  .\build.ps1 -Mode Compile'
    Write-Host '  .\build.ps1 -Mode DeployDebugAddIn'
    Write-Host '  .\build.ps1 -Mode DeployReleaseAddIn -Project All -ReleaseCertificateKeyFile C:\certs\CaseInfoSystem.InternalRelease.pfx -ReleaseCertificateThumbprint <thumbprint>'
}

try {
    $repoRoot = Split-Path -Path $PSCommandPath -Parent

    switch ($Mode) {
        'Help' {
            Show-Help
            exit 0
        }

        'Test' {
            $solutionPath = Join-Path $repoRoot 'dev\CaseInfoSystem.slnx'
            if (-not (Test-Path -LiteralPath $solutionPath)) {
                throw "Solution file was not found: $solutionPath"
            }

            $resolvedConfiguration = if ([string]::IsNullOrWhiteSpace($Configuration)) { 'Debug' } else { $Configuration }
            $arguments = @('test', $solutionPath, '-c', $resolvedConfiguration)

            if ($NoBuild) {
                $arguments += '--no-build'
            }

            if ($NoRestore) {
                $arguments += '--no-restore'
            }

            Invoke-ExternalCommand -FilePath 'dotnet' -Arguments $arguments
        }

        'Compile' {
            $solutionPath = Join-Path $repoRoot 'dev\CaseInfoSystem.slnx'
            if (-not (Test-Path -LiteralPath $solutionPath)) {
                throw "Solution file was not found: $solutionPath"
            }

            $resolvedConfiguration = if ([string]::IsNullOrWhiteSpace($Configuration)) { 'Release' } else { $Configuration }

            # Compile mode is intentionally packaging-free so it can run under dotnet/MSBuild Core.
            $arguments = @(
                'build',
                $solutionPath,
                '-c', $resolvedConfiguration,
                '/p:AllowCoreBuildWithoutVstoPackaging=true',
                '/p:SignManifests=false',
                '/p:ManifestCertificateThumbprint='
            )

            if ($NoRestore) {
                $arguments += '--no-restore'
            }

            Invoke-ExternalCommand -FilePath 'dotnet' -Arguments $arguments
        }

        'DeployDebugAddIn' {
            $deployScriptPath = Join-Path $repoRoot 'scripts\Invoke-DeployDebugAddIns.ps1'
            if (-not (Test-Path -LiteralPath $deployScriptPath)) {
                throw "Deploy script was not found: $deployScriptPath"
            }

            $arguments = @(
                '-ExecutionPolicy', 'Bypass',
                '-File', $deployScriptPath,
                '-Project', $Project
            )

            if (-not [string]::IsNullOrWhiteSpace($MsBuildPath)) {
                $arguments += @('-MsBuildPath', $MsBuildPath)
            }

            Invoke-ExternalCommand -FilePath 'powershell.exe' -Arguments $arguments
        }

        'DeployReleaseAddIn' {
            $deployScriptPath = Join-Path $repoRoot 'scripts\Invoke-DeployReleaseAddIns.ps1'
            if (-not (Test-Path -LiteralPath $deployScriptPath)) {
                throw "Release deploy script was not found: $deployScriptPath"
            }

            $arguments = @(
                '-ExecutionPolicy', 'Bypass',
                '-File', $deployScriptPath,
                '-Project', $Project
            )

            if (-not [string]::IsNullOrWhiteSpace($MsBuildPath)) {
                $arguments += @('-MsBuildPath', $MsBuildPath)
            }

            if (-not [string]::IsNullOrWhiteSpace($ReleaseCertificateKeyFile)) {
                $arguments += @('-ReleaseCertificateKeyFile', $ReleaseCertificateKeyFile)
            }

            if (-not [string]::IsNullOrWhiteSpace($ReleaseCertificateThumbprint)) {
                $arguments += @('-ReleaseCertificateThumbprint', $ReleaseCertificateThumbprint)
            }

            if (-not [string]::IsNullOrWhiteSpace($ManifestCertificatePassword)) {
                $arguments += @('-ManifestCertificatePassword', $ManifestCertificatePassword)
            }

            Invoke-ExternalCommand -FilePath 'powershell.exe' -Arguments $arguments
        }
    }
}
catch {
    Write-Error $_
    exit 1
}
