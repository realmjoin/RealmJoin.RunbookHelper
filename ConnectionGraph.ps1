function Connect-RjRbGraph {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    if ($RjRbRunningInAzure) {
        Write-Warning "Getting automation connection '$AutomationConnectionName'"
        $araCon = Get-AutomationConnection -Name $AutomationConnectionName

        $certPsPath = "Cert:\CurrentUser\My\$($araCon.CertificateThumbprint)"
        Write-Warning "Getting certificate (and key) from '$certPsPath'"
        $cert = Get-Item $certPsPath

        $getAuthTokenParams = [ordered]@{
            TenantId    = $araCon.TenantId
            AppClientId = $araCon.ApplicationId
            CertWithKey = $cert
        }
    }
    else {
        <# Create local certificate for development purposes (values are Automation Connection Name, TenantId, Application Client Id)
        New-SelfSignedCertificate -Subject 'CN=AzureRunAsConnection, DC=gkcorellia.onmicrosoft.com, OU=d7bd21d4-27ca-4f35-b108-284b283a4754' `
            -CertStoreLocation "cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(10) -KeySpec Signature | `
            Export-Certificate  -FilePath "AzureRunAsConnection.cer"
        #>
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -imatch "\bCN\s*=\s*$AutomationConnectionName," } | Select-Object -First 1
        $subjectParsed = @{}; ($cert.Subject | Select-String "\b(CN|DC|OU)\s*=\s*([^,]*)" -AllMatches).Matches | ForEach-Object { $subjectParsed[$_.Groups[1].Value] = $_.Groups[2].Value }
        $getAuthTokenParams = [ordered]@{
            TenantId    = $subjectParsed["DC"]
            AppClientId = $subjectParsed["OU"]
            CertWithKey = $cert
        }
    }

    Write-Warning "Getting Graph authentication token"
    Write-Warning "$($MyInvocation.InvocationName): $([PSCustomObject]$getAuthTokenParams | Format-List | Out-String)"
    return getGraphAuthToken @getAuthTokenParams
}

function getGraphAuthToken([string] $tenantId, [string] $appClientId,
    [System.Security.Cryptography.X509Certificates.X509Certificate2] $certWithKey) {

    $oauthUri = "https://login.microsoftonline.com/$tenantId/oauth2/token"
    $jwt = createAzureGraphLogonJwtWithCert $oauthUri $appClientId $certWithKey

    $invokeParams = [ordered]@{
        Method = "POST"
        Uri    = $oauthUri
        Body   = [ordered]@{ 
            scope                 = "https://graph.microsoft.com/.default"
            client_id             = $appClientId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $jwt
            grant_type            = "client_credentials"
        }
    }
    Write-Warning "$($MyInvocation.InvocationName): $([PSCustomObject]$invokeParams | ConvertTo-Json)"
    $result = Invoke-RestMethod @invokeParams

    @{
        'Content-Type'  = 'application/json'
        'Authorization' = "Bearer " + $result.access_token
        'ExpiresOn'     = $result.expires_on
    }
}

# based on https://github.com/SP3269/posh-jwt
function createAzureGraphLogonJwtWithCert([string] $oauthTokenUri, [string] $appClientId,
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
    Write-Warning "$($MyInvocation.InvocationName): $([PSCustomObject]@{ Cert = $certWithKey.Subject; Header = $headerJson; Payload = $payloadJson } | Format-List | Out-String)"

    $jwt = "$(convertToBase64UrlString $headerJson).$(convertToBase64UrlString $payloadJson)"

    $sigBytes = $rsaPrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($jwt), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $jwt += ".$(convertToBase64UrlString $sigBytes)"

    return $jwt
}
