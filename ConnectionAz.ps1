function Connect-RjRbAzAccount {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    Write-RjRbLog "Connecting with Az module" $autoCon
    Connect-AzAccount -CertificateThumbprint $autoCon.CertificateThumbprint -ApplicationId $autoCon.ApplicationId `
        -TenantId $autoCon.TenantId | Out-Null
}
