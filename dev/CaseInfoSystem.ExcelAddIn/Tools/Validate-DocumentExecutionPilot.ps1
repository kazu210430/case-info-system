param(
    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

Write-Output ('Document execution pilot validation is disabled. policyDirectory=' + $PolicyDirectory)
