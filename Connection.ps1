function getConnectArgs([string] $serviceNameStub, [bool] $preferManagedIdentity, [string] $AutomationConnectionName, [switch] $automationConnectionOnly) {

    if ($RjRbRunningInAzure) {

        ## check if automation connection is available - Feature obsolote as of 2023-09-30
        #$automationConnection = Get-AutomationConnection -Name $AutomationConnectionName -EA 0
        $automationConnectionAvailable = $false

        if ($automationConnectionOnly) {
            if ($automationConnectionAvailable) {
                Write-RjRbLog "Using automation connection '$AutomationConnectionName' (service ${serviceNameStub})"
                return $automationConnection
            }
            else {
                throw "Automation connection '$AutomationConnectionName' was not found and an Azure managed identity is not supported by service ${serviceNameStub}"
            }
        }

        if ($automationConnectionAvailable) {
            # still check if managed identity is preferred
            $enforceManagedIdentityValue = Get-AutomationVariable "RJRB_ENFORCE_MANAGED_IDENTITY" -EA 0
            $enforceManagedIdentityServiceValue = Get-AutomationVariable "RJRB_ENFORCE_MANAGED_IDENTITY_${serviceNameStub}" -EA 0
            $ignoreManagedIdentityValue = Get-AutomationVariable "RJRB_IGNORE_MANAGED_IDENTITY" -EA 0
            $ignoreManagedIdentityServiceValue = Get-AutomationVariable "RJRB_IGNORE_MANAGED_IDENTITY_${serviceNameStub}" -EA 0
            Write-RjRbDebug -Data @{
                automationConnectionName           = $AutomationConnectionName
                automationConnectionThumbprint     = $automationConnection.CertificateThumbprint
                enforceManagedIdentityValue        = $enforceManagedIdentityValue
                enforceManagedIdentityServiceValue = $enforceManagedIdentityServiceValue
                ignoreManagedIdentityValue         = $ignoreManagedIdentityValue
                ignoreManagedIdentityServiceValue  = $ignoreManagedIdentityServiceValue
            }
            if ([bool][int]$enforceManagedIdentityValue -or [bool][int]$enforceManagedIdentityServiceValue) {
                $preferManagedIdentity = $true
            }
            elseif ([bool][int]$ignoreManagedIdentityValue -or [bool][int]$ignoreManagedIdentityServiceValue) {
                $preferManagedIdentity = $false
            }
            if (-not $preferManagedIdentity) {
                Write-RjRbLog "Using automation connection '$AutomationConnectionName' instead of Azure managed identity (service ${serviceNameStub})"
                return $automationConnection
            }
        }

        # check if Azure managed identiy is really available
        try {
            Invoke-RestMethod -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER; 'Metadata' = 'True' } `
                -Uri "${env:IDENTITY_ENDPOINT}?resource=https://graph.microsoft.com/" -UseBasicParsing | Out-Null
            $managedIdentityAvailable = $true
        }
        catch {
            $managedIdentityAvailable = $false
        }

        # return what is available
        Write-RjRbDebug -Data @{
            managedIdentityAvailable      = $managedIdentityAvailable
            automationConnectionAvailable = $automationConnectionAvailable
        }
        if ($managedIdentityAvailable) {
            Write-RjRbLog "Found Azure managed identity and using it (service ${serviceNameStub})"
            return @{ Identity = $true }
        }
        elseif ($automationConnectionAvailable) {
            Write-RjRbLog "Did not find Azure managed identity, using automation connection '$AutomationConnectionName' instead (service ${serviceNameStub})"
            return $automationConnection
        }
        throw "Neither an Azure managed identity nor an automation connection was found (service ${serviceNameStub})"
    }

    else {
        return devGetAutomationConnectionFromLocalCertificate -Name $AutomationConnectionName
    }
}
