# $VerbosePreference automatically is set to "Continue" by turning on "Log verbose records"
# But we only want to use the verbose stream for our own log output and not for verbose output from any other cmdlets
# that are getting called.
$Global:VerbosePreference = "SilentlyContinue"

# Default should be to terminate on any errors when using our module
$Global:ErrorActionPreference = "Stop"
# We still want errors occuring inside this module to be terminating even if ErrorActionPreference is being changed again
# globally, so we also set this locally since then it will still take effect inside this module's functions
$ErrorActionPreference = "Stop"

$Global:RjRbRunningInAzure = [bool]$env:AUTOMATION_ASSET_ACCOUNTID

. $PSScriptRoot\Logging.ps1

$logPrefix = "$([IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)):"
if ($RjRbRunningInAzure) {
    Write-RjRbLog "$logPrefix Running in Azure Automation account"
}
else {
    Write-RjRbLog "$logPrefix Not running in Azure - probably development environment"
    . $PSScriptRoot\DevCertificates.ps1
}

. $PSScriptRoot\RJInterface.ps1

. $PSScriptRoot\Rest.ps1
. $PSScriptRoot\Connection.ps1
. $PSScriptRoot\ConnectionAz.ps1
. $PSScriptRoot\ConnectionAzureAD.ps1
. $PSScriptRoot\ConnectionExchangeOnline.ps1
. $PSScriptRoot\ConnectionGraph.ps1
