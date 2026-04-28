param(
    [Parameter(Mandatory = $true)]
    [string]$PolicyDirectory
)

$ErrorActionPreference = 'Stop'

Write-Output ('Document execution mode validation is disabled. policyDirectory=' + $PolicyDirectory)
