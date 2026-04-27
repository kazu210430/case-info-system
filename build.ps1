param(
    [ValidateSet('Test', 'Compile', 'DeployDebugAddIn', 'Help')]
    [string]$Mode = 'Test',

    [ValidateSet('ExcelAddIn', 'WordAddIn', 'All')]
    [string]$Project = 'All',

    [string]$Configuration,

    [string]$MsBuildPath,

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
    Write-Host 'Supported modes:'
    Write-Host '  .\build.ps1 -Mode Test'
    Write-Host '  .\build.ps1 -Mode Compile'
    Write-Host '  .\build.ps1 -Mode DeployDebugAddIn'
    Write-Host ''
    Write-Host 'Notes:'
    Write-Host '  Test                Runs dotnet test for dev\CaseInfoSystem.slnx.'
    Write-Host '  Compile             Runs compile-only validation without VSTO packaging.'
    Write-Host '  DeployDebugAddIn    Wraps scripts\Invoke-DeployDebugAddIns.ps1 and reflects Debug add-ins into runtime Addins.'
}

try {
    $repoRoot = Split-Path -Path $PSCommandPath -Parent
    $solutionPath = Join-Path $repoRoot 'dev\CaseInfoSystem.slnx'
    $deployScriptPath = Join-Path $repoRoot 'scripts\Invoke-DeployDebugAddIns.ps1'

    if (-not (Test-Path -LiteralPath $solutionPath)) {
        throw "Solution file was not found: $solutionPath"
    }

    switch ($Mode) {
        'Help' {
            Show-Help
            exit 0
        }

        'Test' {
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
    }
}
catch {
    Write-Error $_
    exit 1
}
