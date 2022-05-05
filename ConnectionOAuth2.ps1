function Connect-RjRbGraph {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection",
        [switch] $Force,
        [switch] $ReturnAuthHeaders
    )

    connectOAuth2Impl "RjRbGraphAuthHeaders" "https://graph.microsoft.com/.default" @PSBoundParameters
}

function Connect-RjRbDefenderATP {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection",
        [switch] $Force,
        [switch] $ReturnAuthHeaders
    )

    connectOAuth2Impl "RjRbDefenderATPAuthHeaders" "https://securitycenter.onmicrosoft.com/windowsatpservice/.default" @PSBoundParameters
}


function connectOAuth2Impl
(
    [string] $tokenVariableName,
    [string] $scope,
    [string] $AutomationConnectionName = "AzureRunAsConnection",
    [switch] $Force,
    [switch] $ReturnAuthHeaders
) {

    $requestNewToken = $false

    # check for expiration
    if ($Force -or -not (Test-Path "Variable:Script:$tokenVariableName")) {
        $requestNewToken = $true
    }

    elseif (Test-Path "Variable:Script:$tokenVariableName") {
        $existingTokenExpiration = (Get-Variable -Scope Script -Name "${tokenVariableName}Expiration" -ValueOnly)
        if ([DateTimeOffset]::UtcNow.AddMinutes(15) -gt $existingTokenExpiration) {
            Write-RjRbLog "Found existing token which will expire soon ($($existingTokenExpiration.ToString('u'))), requesting new token."
            $requestNewToken = $true
        }
        else {
            Write-RjRbLog "Found existing token which is still valid (until $($existingTokenExpiration.ToString('u'))), skipping new request."
        }
    }

    if ($requestNewToken) {

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
            scope       = $scope
        }
        $tokenResult = requestOAuth2AccessToken @getAuthTokenParams

        Set-Variable -Scope Script -Name $tokenVariableName -Value @{ Authorization = $tokenResult.Authorization }
        Set-Variable -Scope Script -Name "${tokenVariableName}Expiration" -Value $tokenResult.Expiration
    }

    if ($ReturnAuthHeaders) {
        return (Get-Variable -Scope Script -Name $tokenVariableName -ValueOnly)
    }
}

function requestOAuth2AccessToken(
    [string] $tenantId,
    [string] $appClientId,
    [System.Security.Cryptography.X509Certificates.X509Certificate2] $certWithKey,
    [string] $scope
) {

    $oauthUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    Write-RjRbLog -Data ([ordered]@{ tenantId = $tenantId; oauthUri = $oauthUri; appClientId = $appClientId; certSubject = $certWithKey.Subject; certThumbp = $certWithKey.Thumbprint; scope = $scope })

    $jwt = createSignedJwt $oauthUri $appClientId $certWithKey

    $invokeRestParams = [ordered]@{
        Method = "POST"
        Uri    = $oauthUri
        Body   = [ordered]@{ 
            scope                 = $scope # property name would need to be 'resource' instead of 'scope' for the old AAD endpoint (without /v2.0/)
            client_id             = $appClientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $jwt
            grant_type            = "client_credentials"
        }
    }
    $result = Invoke-RjRbRestMethod @invokeRestParams

    $tokenParts = $result.access_token.Split('.')
    $tokenHeader = convertFromBase64UrlString $tokenParts[0] | ConvertFrom-Json
    Write-RjRbDebug "access_token header" $tokenHeader
    $tokenPayload = convertFromBase64UrlString $tokenParts[1] | ConvertFrom-Json
    Write-RjRbDebug "access_token payload" $tokenPayload

    $expiration = [DateTimeOffset]::FromUnixTimeSeconds($tokenPayload.exp)
    $roles = $tokenPayload.roles
    Write-RjRbLog "Received token, expiration: $($expiration.ToString('u')), roles: $($roles -join ',')"

    return @{
        Authorization = "$($result.token_type) $($result.access_token)"
        Expiration    = $expiration
    }
}

# based on https://github.com/SP3269/posh-jwt
function createSignedJwt(
    [string] $oauthTokenUri,
    [string] $appClientId,
    [System.Security.Cryptography.X509Certificates.X509Certificate2] $certWithKey
) {

    $rsaPrivateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certWithKey)
    if (-not $rsaPrivateKey) { throw [ArgumentNullException]::new("rsaPrivateKey") }

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

function convertToBase64UrlString ([object] $in) {

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

function convertFromBase64UrlString ([string] $in) {

    $in = $in -replace '-', '+' -replace '_', '/'
    if ($in.Length % 4) {
        $in += ('=' * (4 - $in.Length % 4)) 
    }

    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($in))
}
