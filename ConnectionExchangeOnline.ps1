function Connect-RjRbExchangeOnline {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    Write-RjRbLog "Connecting with ExchangeOnline module" $autoCon
    Connect-ExchangeOnline -CertificateThumbprint $autoCon.CertificateThumbprint -AppId $autoCon.ApplicationId `
        -Organization $autoCon.TenantId -ShowBanner:$false
}
