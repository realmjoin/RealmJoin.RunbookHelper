function Connect-RjRbAzAccount {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    Write-RjRbLog "Connecting with Az module" $autoCon
    Connect-AzAccount -ServicePrincipal -TenantId $autoCon.TenantId -ApplicationId $autoCon.ApplicationId `
        -CertificateThumbprint $autoCon.CertificateThumbprint | Out-Null
}
