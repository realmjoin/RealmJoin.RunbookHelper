function Connect-RjRbAzureAD {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    if ($RjRbRunningInAzure) {
        Write-Warning "Getting automation connection '$AutomationConnectionName'"
        $araCon = Get-AutomationConnection -Name $AutomationConnectionName -EA Stop

        Write-Warning "Connecting with AzureAD module"
        Connect-AzureAD -CertificateThumbprint $araCon.CertificateThumbprint -ApplicationId $araCon.ApplicationId -TenantId $araCon.TenantId -EA Stop | Out-Null
    }
    else {
        $storedCredential = Get-RjRbStoredCredential
        if ($storedCredential) {
            Connect-AzureAD -Credential $storedCredential -EA Stop | Out-Null
        }
        else {
            Connect-AzureAD -EA Stop | Out-Null
        }
    }
}

function Get-RjRbAzureADTenantDetail {
    Write-Warning "Getting Azure AD tenant details"
    $aadTenantDetail = Get-AzureADTenantDetail
    return [PSCustomObject]@{
        UpnSuffix   = $aadTenantDetail.VerifiedDomains | Where-Object { $_._Default } | Select-Object -ExpandProperty Name
        DisplayName = $aadTenantDetail.DisplayName
        RawValues   = $aadTenantDetail
    }
}
