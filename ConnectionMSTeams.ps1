function Connect-RjRbMSTeams {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    Write-RjRbLog "Connecting with MS Teams module" $autoCon
    Connect-MicrosoftTeams -TenantId $autoCon.TenantId -ApplicationId $autoCon.ApplicationId `
        -CertificateThumbprint $autoCon.CertificateThumbprint | Out-Null
}
