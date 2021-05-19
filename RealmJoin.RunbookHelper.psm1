$ErrorActionPreference = "Stop" # takes effect only inside module functions

$Global:RjRbRunningInAzure = [bool]$env:AUTOMATION_ASSET_ACCOUNTID
$Global:RjRbLogPrefix = "$([IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)):"

if ($RjRbRunningInAzure) {
    Write-Warning "$RjRbLogPrefix Running in Azure Automation account"
}
else {
    Write-Warning "$RjRbLogPrefix Not running in Azure - probably development environment"
    . $PSScriptRoot\DevCredentials.ps1
}

. $PSScriptRoot\RJInterface.ps1

. $PSScriptRoot\ConnectionAzureAD.ps1
. $PSScriptRoot\ConnectionGraph.ps1
