function Connect-RjRbAzAccount {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $connectParams = @{}
    if (checkIfManagedIdentityShouldBeUsed 'AZ' $true) {
        $connectParams += @{ Identity = $true }
    }
    else {
        $connectParams += getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName
    }

    Write-RjRbLog "Connecting with Az module" $connectParams
    Connect-AzAccount @connectParams | Out-Null
}
