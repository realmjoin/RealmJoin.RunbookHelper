function Connect-RjRbGraph {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection",
        [switch] $Force
    )

    if (-not $Force -and (Test-Path Variable:Script:RjRbGraphAuthHeaders)) {
        # already have an access token
        return
    }

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
    $tokenResult = authenticateToGraphWithCert @getAuthTokenParams

    $Script:RjRbGraphAuthHeaders = @{
        'Authorization' = "Bearer " + $tokenResult.access_token
    }
}

function authenticateToGraphWithCert([string] $tenantId, [string] $appClientId,
    [System.Security.Cryptography.X509Certificates.X509Certificate2] $certWithKey) {

    $oauthUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    Write-RjRbLog -Data ([ordered]@{ tenantId = $tenantId; oauthUri = $oauthUri; appClientId = $appClientId; certSubject = $certWithKey.Subject; certThumbp = $certWithKey.Thumbprint })

    $jwt = createSignedGraphLogonJwt $oauthUri $appClientId $certWithKey

    $invokeRestParams = [ordered]@{
        Method        = "POST"
        Uri           = $oauthUri
        Body          = [ordered]@{ 
            scope                 = "https://graph.microsoft.com/.default" # needs to be resource for the old AAD endpoint (without /v2.0/)
            client_id             = $appClientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $jwt
            grant_type            = "client_credentials"
        }
    }
    return Invoke-RjRbRestMethod @invokeRestParams
}

# based on https://github.com/SP3269/posh-jwt
function createSignedGraphLogonJwt([string] $oauthTokenUri, [string] $appClientId,
    [System.Security.Cryptography.X509Certificates.X509Certificate2] $certWithKey) {

    $rsaPrivateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certWithKey)
    if (-not $rsaPrivateKey) { throw [ArgumentNullException]::new("rsaPrivateKey") }

    function convertToBase64UrlString ($in) {
        if ($in -is [string]) {
            $in = [System.Text.Encoding]::UTF8.GetBytes($in)
        }
        if ($in -is [byte[]]) {
            return [Convert]::ToBase64String($in) -replace '\+', '-' -replace '/', '_' -replace '='
        }
        else {
            throw [InvalidOperationException]::new($in.GetType())
        }
    }

    $thumbprintBytes = [byte[]] ($certWithKey.Thumbprint -replace '..', '0x$&,' -split ',' -ne '')
    $header = [ordered]@{
        typ = "JWT"
        alg = "RS256"
        x5t = [Convert]::ToBase64String($thumbprintBytes)
    }

    $nowEpoch = [int64](([datetime]::UtcNow) - (Get-Date "1970-01-01")).TotalSeconds
    $payload = [ordered]@{
        jti = (New-Guid)
        aud = $oauthTokenUri
        iss = $appClientId
        sub = $appClientId
        nbf = $nowEpoch - 10     # account for minor clock deviation
        exp = $nowEpoch + 60 * 5 # 5 minutes
    }

    $headerJson = $header | ConvertTo-Json -Compress
    $payloadJson = $payload | ConvertTo-Json -Compress
    $jwt = "$(convertToBase64UrlString $headerJson).$(convertToBase64UrlString $payloadJson)"
    Write-RjRbDebug -Data ([ordered]@{ header = $header; payload = $payload; headerJson = $headerJson; payloadJson = $payloadJson; jwtWoSig = $jwt })

    $sigBytes = $rsaPrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($jwt), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $jwt += ".$(convertToBase64UrlString $sigBytes)"

    Write-RjRbDebug -Data @{ JWT = $jwt }

    return $jwt
}
