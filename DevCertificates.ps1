<#
Create local certificate for development purposes
Values in subject:
    CN: Automation Connection Name
    OU: ApplicationId (client id guid)
    DC: TenantId (guid)
    O:  SubscriptionId (guid)

New-SelfSignedCertificate -Subject 'CN=AzureRunAsConnection, OU=21cf3618-9cf9-49a4-bd12-a29e297254b5, DC=810017af-c1d4-4a3b-9e18-d693c12b0df1, O=e0e2ba22-1184-4254-90a4-cddcf7f39886' `
    -CertStoreLocation "cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(10) -KeySpec Signature | `
    Export-Certificate -FilePath "AzureRunAsConnection.cer"

Then upload the resulting .cer-File as an additional certificate to the Automation Account's App Registration in AzureAD
and you're good to go and test with exactly the same identity that the runbooks will get.

#>

function devGetAutomationConnectionFromLocalCertificate {
    param (
        [string] $Name = "AzureRunAsConnection"
    )

    Write-RjRbLog "Looking for local certificate matching '$Name'"
    $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -imatch "\bCN\s*=\s*$Name\b" } | Select-Object -First 1
    if (-not $cert) { throw [IO.FileNotFoundException]::new($Name) }

    Write-RjRbLog "Parsing certificate subject '$($cert.Subject)'"
    $subjectParsed = @{}; ($cert.Subject | Select-String "\b(OU|DC|O)\s*=\s*([^,]*)" -AllMatches).Matches | ForEach-Object { $subjectParsed[$_.Groups[1].Value] = $_.Groups[2].Value }
    $result = @{
        CertificateThumbprint = $cert.Thumbprint
    }

    function addToResult([string] $subjectKey, [string] $resultKey) {
        if ($subjectParsed.Keys -contains $subjectKey) {
            $result[$resultKey] = $subjectParsed[$subjectKey]
        }
    }
    addToResult "OU" "ApplicationId"
    addToResult "DC" "TenantId"
    addToResult "O"  "SubscriptionId"

    return $result;
}
