function getAutomationConnectionOrFromLocalCertificate([string] $AutomationConnectionName) {
    if ($RjRbRunningInAzure) {
        Write-RjRbLog "Getting automation connection '$AutomationConnectionName'"
        return Get-AutomationConnection -Name $AutomationConnectionName
    }
    else {
        return devGetAutomationConnectionFromLocalCertificate -Name $AutomationConnectionName
    }
}

function checkIfManagedIdentityShouldBeUsed() {

    if ($RjRbRunningInAzure) {
        $ignoreManagedIdentityValue = Get-AutomationVariable 'RJRB_IGNORE_MANAGED_IDENTITY' -EA 0
        Write-RjRbDebug -Data @{ ignoreManagedIdentityValue = $ignoreManagedIdentityValue }
        if ([bool][int]$ignoreManagedIdentityValue) {
            return $false
        }
    }

    try {
        Invoke-RestMethod -Headers @{ 'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True' } `
            -Uri "$($env:IDENTITY_ENDPOINT)?resource=https://graph.microsoft.com/" -UseBasicParsing | Out-Null
        $managedIdentityAvailable = $true
    }
    catch {
        $managedIdentityAvailable = $false
    }
    Write-RjRbDebug -Data @{ managedIdentityAvailable = $managedIdentityAvailable }

    if ($managedIdentityAvailable) {
        Write-RjRbLog "Using Azure managed identity"
    }

    return $managedIdentityAvailable
}
