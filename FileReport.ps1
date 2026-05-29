function Get-RjRbStorageSharedKeyAuthHeader {
    <#
        .SYNOPSIS
        Build a SharedKey Authorization header for an Azure Storage REST request.
    #>
    param(
        [Parameter(Mandatory = $true)][string] $StorageAccountName,
        [Parameter(Mandatory = $true)][byte[]] $KeyBytes,
        [Parameter(Mandatory = $true)][string] $Method,
        [Parameter(Mandatory = $true)][string] $CanonicalizedResource,
        [Parameter(Mandatory = $true)][hashtable] $Headers,
        [string] $ContentType = "",
        [int] $ContentLength = 0
    )

    $msHeaders = ($Headers.GetEnumerator() | Where-Object { $_.Key -like "x-ms-*" } | Sort-Object Key | ForEach-Object { "$($_.Key):$($_.Value)" }) -join "`n"
    $contentLengthStr = if ($ContentLength -gt 0) { "$ContentLength" } else { "" }
    $stringToSign = "$Method`n`n`n$contentLengthStr`n`n$ContentType`n`n`n`n`n`n`n$msHeaders`n$CanonicalizedResource"

    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    try {
        $hmac.Key = $KeyBytes
        $sig = [Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($stringToSign)))
    }
    finally {
        $hmac.Dispose()
    }
    return "SharedKey ${StorageAccountName}:$sig"
}

function New-RjRbBlobSasToken {
    <#
        .SYNOPSIS
        Generate a read-only Blob SAS URL signed with the storage account key.
    #>
    param(
        [Parameter(Mandatory = $true)][string] $StorageAccountName,
        [Parameter(Mandatory = $true)][byte[]] $KeyBytes,
        [Parameter(Mandatory = $true)][string] $Container,
        [Parameter(Mandatory = $true)][string] $Blob,
        [Parameter(Mandatory = $true)][datetime] $ExpiryTime
    )

    $startTime = (Get-Date).ToUniversalTime().AddMinutes(-5).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $expiryStr = $ExpiryTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $permissions = "r"
    $signedVersion = "2023-11-03"
    $signedResource = "b"
    $signedProtocol = "https"

    $canonicalizedResource = "/blob/$StorageAccountName/$Container/$Blob"
    $stringToSign = "$permissions`n$startTime`n$expiryStr`n$canonicalizedResource`n`n`n$signedProtocol`n$signedVersion`n$signedResource`n`n`n`n`n`n`n"

    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    try {
        $hmac.Key = $KeyBytes
        $sig = [Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($stringToSign)))
    }
    finally {
        $hmac.Dispose()
    }

    $sasToken = "sp=$permissions&st=$startTime&se=$expiryStr&spr=$signedProtocol&sv=$signedVersion&sr=$signedResource&sig=$([Uri]::EscapeDataString($sig))"
    return "https://$StorageAccountName.blob.core.windows.net/$Container/${Blob}?$sasToken"
}

function Publish-RjRbFilesToStorageContainer {
    <#
        .SYNOPSIS
        Upload one or more local files to an Azure Storage container, returning SAS
        download links.

        .DESCRIPTION
        Performs blob upload and SAS token generation using the Azure Storage REST API
        directly, avoiding the Az.Storage module entirely. This eliminates the well-known
        assembly conflict between Az.Storage and ExchangeOnlineManagement.

        Storage account keys are retrieved via ARM REST API (Invoke-AzRestMethod from
        Az.Accounts). Blob operations (container creation, upload, SAS generation) use
        the Azure Storage REST API with SharedKey authentication.

        Required Azure RBAC on the storage account:
        - Microsoft.Storage/storageAccounts/read
        - Microsoft.Storage/storageAccounts/listKeys/action
        Built-in role: 'Storage Account Contributor'.

        .PARAMETER FilePaths
        Array of local file paths to upload.

        .PARAMETER ContainerName
        Target blob container. Created automatically if missing.

        .PARAMETER ResourceGroupName
        Resource group containing the storage account.

        .PARAMETER StorageAccountName
        Target storage account name.

        .PARAMETER SubscriptionId
        Optional Azure subscription ID. Sets context before storage operations.

        .PARAMETER LinkExpiryDays
        SAS link validity in days (default 6, range 1-3650).

        .PARAMETER AddBlobNamePrefix
        When $true, prefixes blob names with yyyyMMdd-HHmmss (default $false).

        .OUTPUTS
        Array of PSCustomObject with BlobName, EndTime, SASLink for each uploaded file.

        .NOTES
        Dependencies:
        - Requires the Az.Accounts module in the runbook environment (cmdlets Get-AzContext,
          Set-AzContext, Connect-AzAccount, Invoke-AzRestMethod). Declare it explicitly in
          the consuming runbook, e.g.:
            #Requires -Modules @{ModuleName = "Az.Accounts"; ModuleVersion = "5.3.4"}
        - No dependency on Az.Storage: blob operations are performed via the Azure Storage
          REST API directly to avoid the Az.Storage / ExchangeOnlineManagement assembly
          conflict.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string[]] $FilePaths,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $ContainerName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $ResourceGroupName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $StorageAccountName,
        [Parameter(Mandatory = $false)][string] $SubscriptionId,
        [Parameter(Mandatory = $false)][ValidateRange(1, 3650)][int] $LinkExpiryDays = 6,
        [Parameter(Mandatory = $false)][bool] $AddBlobNamePrefix = $false
    )

    if (-not (Get-Command -Name Get-AzContext -ErrorAction SilentlyContinue)) {
        throw "Publish-RjRbFilesToStorageContainer requires the 'Az.Accounts' module. Add `#Requires -Modules @{ModuleName = 'Az.Accounts'; ModuleVersion = '5.3.4'}` to the calling runbook."
    }

    foreach ($p in $FilePaths) {
        if (-not (Test-Path -Path $p -PathType Leaf)) {
            throw "File '$p' was not found."
        }
    }

    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if ((-not $azContext) -or (-not $azContext.Account)) {
        Connect-RjRbAzAccount
    }
    if ($SubscriptionId) { Set-AzContext -Subscription $SubscriptionId | Out-Null }

    $effectiveSubscriptionId = (Get-AzContext).Subscription.Id
    $armPath = "/subscriptions/$effectiveSubscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Storage/storageAccounts/$StorageAccountName/listKeys?api-version=2023-05-01"
    $keysResponse = Invoke-AzRestMethod -Path $armPath -Method POST
    if ($keysResponse.StatusCode -ne 200) {
        throw "Failed to retrieve storage account keys for '$StorageAccountName' in resource group '$ResourceGroupName'. Status: $($keysResponse.StatusCode)"
    }
    $storageKey = ($keysResponse.Content | ConvertFrom-Json).keys[0].value
    $keyBytes = [Convert]::FromBase64String($storageKey)

    $baseUri = "https://$StorageAccountName.blob.core.windows.net"

    # Create container if it does not exist (using HttpClient to bypass Azure Automation's
    # Invoke-RestMethod interceptor that strips required headers)
    $dateStr = [DateTime]::UtcNow.ToString("R")
    $containerHeaders = @{
        "x-ms-date"    = $dateStr
        "x-ms-version" = "2023-11-03"
    }
    $canonResource = "/$StorageAccountName/$ContainerName`nrestype:container"
    $containerHeaders["Authorization"] = Get-RjRbStorageSharedKeyAuthHeader -StorageAccountName $StorageAccountName -KeyBytes $keyBytes -Method "PUT" -CanonicalizedResource $canonResource -Headers $containerHeaders

    $httpClient = [System.Net.Http.HttpClient]::new()
    try {
        $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, "$baseUri/$ContainerName`?restype=container")
        foreach ($h in $containerHeaders.GetEnumerator()) {
            $request.Headers.TryAddWithoutValidation($h.Key, $h.Value) | Out-Null
        }
        $response = $httpClient.SendAsync($request).GetAwaiter().GetResult()
        $statusCode = [int]$response.StatusCode
        if ($statusCode -ne 201 -and $statusCode -ne 409) {
            $errBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            throw "Container creation failed ($statusCode): $errBody"
        }
    }
    finally {
        $httpClient.Dispose()
    }

    $endTime = (Get-Date).AddDays($LinkExpiryDays)
    $results = @()
    foreach ($filePath in $FilePaths) {
        $blobName = Split-Path -Path $filePath -Leaf
        if ($AddBlobNamePrefix) {
            $prefix = (Get-Date).ToString("yyyyMMdd-HHmmss")
            $blobName = "$prefix-$blobName"
        }

        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
        $contentLength = $fileBytes.Length

        $dateStr = [DateTime]::UtcNow.ToString("R")
        $blobHeaders = @{
            "x-ms-date"      = $dateStr
            "x-ms-version"   = "2023-11-03"
            "x-ms-blob-type" = "BlockBlob"
        }
        $canonResource = "/$StorageAccountName/$ContainerName/$blobName"
        $blobHeaders["Authorization"] = Get-RjRbStorageSharedKeyAuthHeader -StorageAccountName $StorageAccountName -KeyBytes $keyBytes -Method "PUT" -CanonicalizedResource $canonResource -Headers $blobHeaders -ContentType "application/octet-stream" -ContentLength $contentLength

        # Use HttpClient directly to ensure all custom headers (x-ms-blob-type) are sent.
        # Invoke-RestMethod in hosted PowerShell can strip custom headers with binary bodies.
        $httpClient = [System.Net.Http.HttpClient]::new()
        try {
            $content = [System.Net.Http.ByteArrayContent]::new($fileBytes)
            $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
            $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, "$baseUri/$ContainerName/$blobName")
            $request.Content = $content
            foreach ($h in $blobHeaders.GetEnumerator()) {
                $request.Headers.TryAddWithoutValidation($h.Key, $h.Value) | Out-Null
            }
            $response = $httpClient.SendAsync($request).GetAwaiter().GetResult()
            if (-not $response.IsSuccessStatusCode) {
                $errBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                throw "Blob upload failed ($($response.StatusCode)): $errBody"
            }
        }
        finally {
            $httpClient.Dispose()
        }

        $sasLink = New-RjRbBlobSasToken -StorageAccountName $StorageAccountName -KeyBytes $keyBytes -Container $ContainerName -Blob $blobName -ExpiryTime $endTime
        $results += [PSCustomObject]@{ BlobName = $blobName; EndTime = $endTime; SASLink = $sasLink }
    }
    return $results
}
