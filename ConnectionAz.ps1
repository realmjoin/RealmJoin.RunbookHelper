function Connect-RjRbAzAccount {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $connectArgs = getConnectArgs 'AZ' $true $AutomationConnectionName

    Write-RjRbLog "Connecting with Az module" $connectArgs
    Connect-AzAccount @connectArgs | Out-Null
}
