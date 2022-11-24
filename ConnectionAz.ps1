function Connect-RjRbAzAccount {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName
    $connectParams = @{}
    if ((-not $autoCon) -or (checkIfManagedIdentityShouldBeUsed 'AZ' $true)) {
        $connectParams += @{ Identity = $true }
    }
    else {
        $connectParams += $autoCon
    }

    Write-RjRbLog "Connecting with Az module" $connectParams
    Connect-AzAccount @connectParams | Out-Null
}
