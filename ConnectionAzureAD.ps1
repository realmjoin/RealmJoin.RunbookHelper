function Connect-RjRbAzureAD {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    if ($RjRbRunningInAzure) {
        Write-RjRbLog "Getting automation connection '$AutomationConnectionName'"
        $autoCon = Get-AutomationConnection -Name $AutomationConnectionName
    }
    else {
        $autoCon = devGetAutomationConnectionFromLocalCertificate -Name $AutomationConnectionName
    }

    Write-RjRbLog "Connecting with AzureAD module" $autoCon
    Connect-AzureAD -CertificateThumbprint $autoCon.CertificateThumbprint -ApplicationId $autoCon.ApplicationId -TenantId $autoCon.TenantId -EA Stop | Out-Null
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
