function Connect-RjRbExchangeOnline {
    [CmdletBinding()]
    param (
        [string] $AutomationConnectionName = "AzureRunAsConnection"
    )

    # see RealmJoin.RunbookHelper.psm1
    $Global:VerbosePreference = "SilentlyContinue"

    $connectArgs = @{ ShowBanner = $false }

    $connectArgsRaw = getConnectArgs 'EXO' $false $AutomationConnectionName
    if ($connectArgsRaw['Identity']) {
        $connectArgs += @{ ManagedIdentity = $true }
    }
    else {
        $connectArgs += @{
            Organization          = $connectArgsRaw.TenantId
            AppId                 = $connectArgsRaw.ApplicationId
            CertificateThumbprint = $connectArgsRaw.CertificateThumbprint
        }
    }


    if ($connectArgs['Organization'] -inotlike "*.onmicrosoft.com") {
        Write-RjRbLog "Trying to determine initial domain name (*.onmicrosoft.com) using Graph"
        Connect-RjRbGraph
        $connectArgs.Organization = Invoke-RjRbRestMethodGraph /organization | Select-Object -ExpandProperty verifiedDomains | `
            Where-Object { $_.isInitial } | Select-Object -First 1 -ExpandProperty name
    }

    Write-RjRbLog "Connecting with ExchangeOnline module" $connectArgs
    if ($connectArgs['ManagedIdentity']) {
        $exoVersion = (Import-Module -Name 'ExchangeOnlineManagement' -Global -PassThru).Version
        if ($exoVersion -lt '3.0.0') {
            throw "Connecting to Exchange Online with a Managed Identity requires at least version 3.0.0 of 'ExchangeOnlineManagement', but only version $exoVersion was found."
        }
    }
    Connect-ExchangeOnline @connectArgs
}
