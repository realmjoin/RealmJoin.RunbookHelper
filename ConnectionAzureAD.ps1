function Connect-RjRbAzureAD {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    # Azure AD PowerShell does not support authentication by managed identity out of the box and is planned for deprecation.
    # see https://learn.microsoft.com/en-us/powershell/azure/active-directory/overview?view=azureadps-2.0
    $connectArgs = getConnectArgs 'AAD' $false $AutomationConnectionName -automationConnectionOnly

    Write-RjRbLog "Connecting with AzureAD module" $connectArgs
    Connect-AzureAD -TenantId $connectArgs.TenantId -ApplicationId $connectArgs.ApplicationId `
        -CertificateThumbprint $connectArgs.CertificateThumbprint | Out-Null
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
