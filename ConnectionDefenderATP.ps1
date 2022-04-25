function Connect-RjRbDefenderATP {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection",
        [switch] $Force,
        [switch] $ReturnAuthHeaders
    )

    if ($Force -or -not (Test-Path Variable:Script:RjRbDefenderATPAuthHeaders)) {

        # see RealmJoin.RunbookHelper.psm1
        $Global:VerbosePreference = "SilentlyContinue"

        $autoCon = getAutomationConnectionOrFromLocalCertificate $AutomationConnectionName

        $certPsPath = "Cert:\CurrentUser\My\$($autoCon.CertificateThumbprint)"
        Write-RjRbLog "Getting certificate (and key) from '$certPsPath'"
        $cert = Get-Item $certPsPath

        $getAuthTokenParams = [ordered]@{
            TenantId    = $autoCon.TenantId
            AppClientId = $autoCon.ApplicationId
            CertWithKey = $cert
        }
        $tokenResult = authenticateToDefenderATPWithCert @getAuthTokenParams

        $Script:RjRbDefenderATPAuthHeaders = @{
            'Authorization' = "Bearer " + $tokenResult.access_token
        }
    }

    if ($ReturnAuthHeaders) {
        return $Script:RjRbDefenderATPAuthHeaders
    }
}

function authenticateToDefenderATPWithCert([string] $tenantId, [string] $appClientId,
    [System.Security.Cryptography.X509Certificates.X509Certificate2] $certWithKey) {

    $oauthUri = "https://login.microsoftonline.com/$tenantId/oauth2/token"
    Write-RjRbLog -Data ([ordered]@{ tenantId = $tenantId; oauthUri = $oauthUri; appClientId = $appClientId; certSubject = $certWithKey.Subject; certThumbp = $certWithKey.Thumbprint })

    $jwt = createSignedGraphLogonJwt $oauthUri $appClientId $certWithKey

    $invokeRestParams = [ordered]@{
        Method = "POST"
        Uri    = $oauthUri
        Body   = [ordered]@{ 
            resource              = "https://api.securitycenter.microsoft.com" # needs to be resource for the old AAD endpoint (without /v2.0/)
            client_id             = $appClientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $jwt
            grant_type            = "client_credentials"
        }
    }
    return Invoke-RjRbRestMethod @invokeRestParams
}
