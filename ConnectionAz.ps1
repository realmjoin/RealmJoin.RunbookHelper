function Connect-RjRbAzAccount {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    Write-RjRbLog "Connecting with Az module" $autoCon
    Connect-AzAccount -ServicePrincipal -TenantId $autoCon.TenantId -ApplicationId $autoCon.ApplicationId `
        -CertificateThumbprint $autoCon.CertificateThumbprint | Out-Null
}
