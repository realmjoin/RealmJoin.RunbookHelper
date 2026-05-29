# Module manifest for module 'RealmJoin.RunbookHelper'

@{
    RootModule        = 'RealmJoin.RunbookHelper.psm1'
    ModuleVersion     = '0.8.6'
    GUID              = '50c59179-6cb8-4968-bf76-e7de04f02957'
    Author            = 'glueckkanja AG'
    CompanyName       = 'glueckkanja AG'
    Copyright         = '(c) glueckkanja AG. All rights reserved.'
    Description       = 'Helps to integrate Azure Automation scripts with RealmJoin.'
    PowerShellVersion = '5.1'
    # RequiredModules = @()

    # should specify all three of the following to speed up command auto-discovery
    FunctionsToExport = @(
        'Use-RjRbInterface', 'Write-RjRbLog', 'Write-RjRbDebug',
        'Invoke-RjRbRestMethod', 'Invoke-RjRbRestMethodGraph', 'Invoke-RjRbRestMethodDefenderATP',
        'Connect-RjRbAzAccount', 'Connect-RjRbAzureAD', 'Get-RjRbAzureADTenantDetail', 'Connect-RjRbExchangeOnline',
        'Connect-RjRbGraph', 'Connect-RjRbDefenderATP', 'Send-RjReportEmail',
        'Publish-RjRbFilesToStorageContainer'
    )
    CmdletsToExport   = @()
    AliasesToExport   = @('Use-RJInterface')

    FileList          = @(
        'RealmJoin.RunbookHelper.psm1',
        'Connection.ps1',
        'ConnectionAz.ps1',
        'ConnectionAzureAD.ps1',
        'ConnectionExchangeOnline.ps1',
        'ConnectionOAuth2.ps1',
        'DevCertificates.ps1',
        'Interface.ps1',
        'InternalHelpers.ps1',
        'Logging.ps1',
        'MailReport.ps1',
        'FileReport.ps1',
        'Rest.ps1',
        'Assets\Header.png',
        'Assets\Footer.png'
    )

    PrivateData       = @{
        PSData = @{
            ProjectUri = 'https://github.com/realmjoin/RealmJoin.RunbookHelper'
            # Prerelease = 'rc1'

            # Informational only - NOT enforced at Import-Module time.
            # Consuming runbooks must declare the modules they actually use via #Requires.
            # - Az.Accounts: required by Publish-RjRbFilesToStorageContainer and Connect-RjRbAzAccount.
            # - Microsoft.Graph.Authentication: only required by Send-RjReportEmail when -UseNativeGraphRequest is set.
            ExternalModuleDependencies = @('Az.Accounts', 'Microsoft.Graph.Authentication')
        }
    }
}

