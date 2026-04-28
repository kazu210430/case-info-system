param(
    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

Write-Output ('Document execution policy validation is disabled. policyDirectory=' + $PolicyDirectory)
