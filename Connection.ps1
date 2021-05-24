function getAutomationConnectionOrFromLocalCertificate([string] $AutomationConnectionName) {
    if ($RjRbRunningInAzure) {
        Write-RjRbLog "Getting automation connection '$AutomationConnectionName'"
        return Get-AutomationConnection -Name $AutomationConnectionName
    }
    else {
        return devGetAutomationConnectionFromLocalCertificate -Name $AutomationConnectionName
    }
}
