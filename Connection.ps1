function getAutomationConnectionOrFromLocalCertificate([string] $AutomationConnectionName) {
    if ($RjRbRunningInAzure) {
        Write-RjRbLog "Getting automation connection '$AutomationConnectionName'"
        return Get-AutomationConnection -Name $AutomationConnectionName
    }
    else {
        return devGetAutomationConnectionFromLocalCertificate -Name $AutomationConnectionName
    }
}

function checkIfManagedIdentityShouldBeUsed([string] $serviceNameStub, [bool] $default) {

    $tryManagedIdentity = $default
    if ($RjRbRunningInAzure) {
        $enforceManagedIdentityValue = Get-AutomationVariable "RJRB_ENFORCE_MANAGED_IDENTITY" -EA 0
        $enforceManagedIdentityServiceValue = Get-AutomationVariable "RJRB_ENFORCE_MANAGED_IDENTITY_${serviceNameStub}" -EA 0
        $ignoreManagedIdentityValue = Get-AutomationVariable "RJRB_IGNORE_MANAGED_IDENTITY" -EA 0
        $ignoreManagedIdentityServiceValue = Get-AutomationVariable "RJRB_IGNORE_MANAGED_IDENTITY_${serviceNameStub}" -EA 0
        Write-RjRbDebug -Data @{ 
            enforceManagedIdentityValue        = $enforceManagedIdentityValue
            enforceManagedIdentityServiceValue = $enforceManagedIdentityServiceValue
            ignoreManagedIdentityValue         = $ignoreManagedIdentityValue
            ignoreManagedIdentityServiceValue  = $ignoreManagedIdentityServiceValue
        }
        if ([bool][int]$enforceManagedIdentityValue -or [bool][int]$enforceManagedIdentityServiceValue) {
            $tryManagedIdentity = $true
        }
        elseif ([bool][int]$ignoreManagedIdentityValue -or [bool][int]$ignoreManagedIdentityServiceValue) {
            $tryManagedIdentity = $false
        }
    }
    if (-not $tryManagedIdentity) {
        Write-RjRbLog "Not trying to use Azure managed identity (service ${serviceNameStub})"
        return $false
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
        Write-RjRbLog "Found Azure managed identity and using it (service ${serviceNameStub})"
    }
    else {
        Write-RjRbLog "Did not find Azure managed identity (service ${serviceNameStub})"
    }

    return $managedIdentityAvailable
}
