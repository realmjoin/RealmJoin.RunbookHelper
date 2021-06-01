function Connect-RjRbExchangeOnline {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    if ($autoCon.TenantId -inotlike "*.onmicrosoft.com") {
        Write-RjRbLog "Trying to determine initial domain name (*.onmicrosoft.com) for tenant Guid '$($autoCon.TenantId)' using Graph"
        Connect-RjRbGraph
        $autoCon.TenantId = Invoke-RjRbRestMethodGraph /organization | Select-Object -ExpandProperty verifiedDomains | `
            Where-Object { $_.isInitial } | Select-Object -First 1 -ExpandProperty name
    }

    Write-RjRbLog "Connecting with ExchangeOnline module" $autoCon
    Connect-ExchangeOnline -Organization $autoCon.TenantId -AppId $autoCon.ApplicationId `
        -CertificateThumbprint $autoCon.CertificateThumbprint -ShowBanner:$false
}
