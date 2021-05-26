function Connect-RjRbAzureAD {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

    Write-RjRbLog "Connecting with AzureAD module" $autoCon
    Connect-AzureAD -TenantId $autoCon.TenantId -ApplicationId $autoCon.ApplicationId `
        -CertificateThumbprint $autoCon.CertificateThumbprint | Out-Null
}

function Get-RjRbAzureADTenantDetail {
    Write-RjRbLog "Getting Azure AD tenant details"
    $aadTenantDetail = Get-AzureADTenantDetail
    return [PSCustomObject]@{
        UpnSuffix   = $aadTenantDetail.VerifiedDomains | Where-Object { $_._Default } | Select-Object -ExpandProperty Name
        DisplayName = $aadTenantDetail.DisplayName
        RawValues   = $aadTenantDetail
    }
}
