# --- Key Vault publishing ----------------------------------------------------
# Push a value into an Azure Key Vault - a secret, a key, or a certificate - and grant
# object-scoped ("per item") RBAC read/use access to selected principals. This lets a
# runbook hand exactly one vault item to specific users, groups, or apps without exposing the rest of the
# vault: access is granted on the item object only
# (.../vaults/<vault>/{secrets|keys|certificates}/<name>), never at vault, resource group,
# or subscription scope. Object-level access requires the vault's Azure RBAC permission
# model (classic access policies are always vault-wide).
#
# Public:
#   Publish-RjRbKeyVaultSecret       - set a secret value
#   Publish-RjRbKeyVaultKey          - create or import a key
#   Publish-RjRbKeyVaultCertificate  - create (self-signed/issuer) or import a certificate
# Private: Verb-RjRbKv* helpers (context, vault lookup, scope/role mapping, reader
#          resolution, RBAC grant, certificate wait, portal URL building, result shaping)
#
# Azure cmdlet dependencies (declare in the consuming runbook via #Requires):
#   Az.Accounts, Az.KeyVault, Az.Resources

# Built-in RBAC role that grants data-plane read/use access to one Key Vault item kind.
function Get-RjRbKvReadRoleName {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType
    )

    switch ($ItemType) {
        'Secret' { 'Key Vault Secrets User' }
        'Certificate' { 'Key Vault Certificate User' }
        'Key' { 'Key Vault Crypto User' }
    }
}

# Maps an item type to the ARM child-resource segment used in scopes and data-plane URIs.
function Get-RjRbKvItemSegment {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType
    )

    switch ($ItemType) {
        'Secret' { 'secrets' }
        'Certificate' { 'certificates' }
        'Key' { 'keys' }
    }
}

# Builds the ARM scope of a single logical Key Vault object. RBAC is object-scoped,
# not version-scoped, so the version is intentionally not part of the scope.
function Get-RjRbKvObjectScope {
    param(
        [Parameter(Mandatory = $true)][string] $VaultResourceId,
        [Parameter(Mandatory = $true)][ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType,
        [Parameter(Mandatory = $true)][string] $ItemName
    )

    '{0}/{1}/{2}' -f $VaultResourceId.TrimEnd('/'), (Get-RjRbKvItemSegment -ItemType $ItemType), $ItemName
}

# Throws a single clear error if any required Az cmdlet is missing from the runbook session.
function Test-RjRbKvRequiredCmdlet {
    param(
        [Parameter(Mandatory = $true)][string[]] $Names
    )

    foreach ($name in $Names) {
        if (-not (Get-Command -Name $name -ErrorAction SilentlyContinue)) {
            throw "Publish-RjRbKeyVault* requires the Az modules. Add '#Requires -Modules Az.Accounts, Az.KeyVault, Az.Resources' to the calling runbook (missing cmdlet '$name')."
        }
    }
}

# Ensures an Az context (reusing the module's connection helper) and optionally pins the
# target subscription before any vault operation.
function Set-RjRbKvAzContext {
    param(
        [string] $SubscriptionId
    )

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if ((-not $context) -or (-not $context.Account)) {
        Connect-RjRbAzAccount
    }
    if ($SubscriptionId) {
        Write-RjRbLog "Setting Az context to subscription '$SubscriptionId'"
        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
    }
}

# Resolves the target vault once so any vault/permission problem fails early with one
# clear error instead of surfacing later as a confusing data-plane failure.
function Get-RjRbKvVault {
    param(
        [Parameter(Mandatory = $true)][string] $VaultName,
        [string] $ResourceGroupName
    )

    $getArgs = @{ VaultName = $VaultName; ErrorAction = 'Stop' }
    if ($ResourceGroupName) { $getArgs['ResourceGroupName'] = $ResourceGroupName }

    $vault = $null
    try {
        $vault = Get-AzKeyVault @getArgs
    }
    catch {
        throw "Key Vault '$VaultName' was not found, or the current identity cannot read it. $($_.Exception.Message)"
    }

    if (-not $vault) {
        $rgHint = if ($ResourceGroupName) { " / resource group '$ResourceGroupName'" } else { '' }
        throw "Key Vault '$VaultName'$rgHint was not found in the current subscription context."
    }

    $vault
}

# Object-scoped access is only possible with the Azure RBAC permission model.
function Test-RjRbKvRbacVault {
    param(
        [Parameter(Mandatory = $true)] $Vault
    )

    if ($Vault.EnableRbacAuthorization -ne $true) {
        throw "Key Vault '$($Vault.VaultName)' does not use the Azure RBAC permission model. Object-scoped access requires RBAC; classic access policies are always vault-wide."
    }
}

# Verifies the required Az cmdlets are present, ensures an Az context (optionally pinning
# the subscription), resolves the target vault, and confirms it uses the Azure RBAC model.
# Returns the validated vault object; it does not create or modify Azure resources.
function Get-RjRbKvValidatedTargetVault {
    param(
        [Parameter(Mandatory = $true)][string[]] $RequiredCmdlets,
        [Parameter(Mandatory = $true)][string] $KeyVaultName,
        [string] $KeyVaultResourceGroupName,
        [string] $SubscriptionId
    )

    Test-RjRbKvRequiredCmdlet -Names $RequiredCmdlets
    Set-RjRbKvAzContext -SubscriptionId $SubscriptionId
    $vault = Get-RjRbKvVault -VaultName $KeyVaultName -ResourceGroupName $KeyVaultResourceGroupName
    Test-RjRbKvRbacVault -Vault $vault
    $vault
}

# Merges caller tags over a default Source tag (caller values win) for traceability.
function New-RjRbKvEffectiveTag {
    param(
        [ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType,
        [hashtable] $Tag
    )

    $merged = @{ Source = "RealmJoin.RunbookHelper/Publish-RjRbKeyVault$ItemType" }
    if ($Tag) { foreach ($key in $Tag.Keys) { $merged[$key] = $Tag[$key] } }
    $merged
}

# Splits a delimited reader list and resolves each entry to an Entra principal object id.
# Accepts user UPN/mail, or an exact display name / object id for users, groups, and service
# principals. Object ids are used as-is without a directory lookup.
function Resolve-RjRbKvReader {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]] $ReaderUsers
    )

    # Allow comma/semicolon/newline separated values inside each array element, too.
    $identifiers = @(
        $ReaderUsers |
            ForEach-Object { $_ -split '[,;\r\n]+' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    $guidPattern = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    $resolved = @()

    foreach ($id in $identifiers) {
        # Plain object id: trust it directly, no directory call required.
        if ($id -match $guidPattern) {
            $resolved += [pscustomobject]@{ Input = $id; ObjectId = [guid]$id; DisplayName = $id; Type = 'ObjectId' }
            continue
        }

        # Try user UPN, then mail, then fall back to exact display-name matches across
        # users, groups, and service principals.
        $user = $null
        if ($id -like '*@*') {
            $user = Get-AzADUser -UserPrincipalName $id -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $user) {
                $user = Get-AzADUser -Mail $id -ErrorAction SilentlyContinue | Select-Object -First 1
            }
        }

        $principalMatches = @()

        if ($user) {
            $principalMatches += [pscustomobject]@{
                Input       = $id
                ObjectId    = [guid]$user.Id
                DisplayName = $user.DisplayName
                Type        = 'User'
            }
        }
        else {
            $userMatches = @(Get-AzADUser -DisplayName $id -ErrorAction SilentlyContinue | Where-Object { $_ })
            foreach ($userMatch in $userMatches) {
                $principalMatches += [pscustomobject]@{
                    Input       = $id
                    ObjectId    = [guid]$userMatch.Id
                    DisplayName = $userMatch.DisplayName
                    Type        = 'User'
                }
            }

            $groupMatches = @(Get-AzADGroup -DisplayName $id -ErrorAction SilentlyContinue | Where-Object { $_ })
            foreach ($groupMatch in $groupMatches) {
                $principalMatches += [pscustomobject]@{
                    Input       = $id
                    ObjectId    = [guid]$groupMatch.Id
                    DisplayName = $groupMatch.DisplayName
                    Type        = 'Group'
                }
            }

            $servicePrincipalMatches = @(Get-AzADServicePrincipal -DisplayName $id -ErrorAction SilentlyContinue | Where-Object { $_ })
            foreach ($servicePrincipalMatch in $servicePrincipalMatches) {
                $principalMatches += [pscustomobject]@{
                    Input       = $id
                    ObjectId    = [guid]$servicePrincipalMatch.Id
                    DisplayName = $servicePrincipalMatch.DisplayName
                    Type        = 'ServicePrincipal'
                }
            }
        }

        if (@($principalMatches).Count -eq 0) {
            throw "Could not resolve '$id' to a Microsoft Entra principal. Use a user UPN/mail address, or an exact display name / object id for a user, group, or service principal."
        }
        if (@($principalMatches).Count -gt 1) {
            $hint = (@($principalMatches) | ForEach-Object { "$($_.DisplayName) <$($_.Type)> [$($_.ObjectId)]" }) -join '; '
            throw "'$id' resolved to multiple principals. Please use the object id. Matches: $hint"
        }

        $resolved += @($principalMatches)[0]
    }

    @($resolved)
}

# Resolves the reader list to principals plus de-duplicated object ids and logs a one-line
# summary. Shared by all three publish functions so the resolve/log/dedupe logic lives once.
function Get-RjRbKvReader {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]] $ReaderUsers
    )

    $readers = @(Resolve-RjRbKvReader -ReaderUsers $ReaderUsers)
    if (@($readers).Count -gt 0) {
        Write-RjRbLog "Resolved $(@($readers).Count) reader principal(s): $((@($readers) | ForEach-Object { $_.DisplayName }) -join ', ')"
    }

    [pscustomobject]@{
        Readers   = @($readers)
        ObjectIds = [guid[]]@($readers | ForEach-Object { $_.ObjectId } | Select-Object -Unique)
    }
}

# Idempotently assigns the read/use role for an item type at the given object scope.
# Only an existing assignment on the exact scope counts as "already granted"; broader
# inherited assignments (vault/RG/subscription) are intentionally left untouched.
function Grant-RjRbKvObjectAccess {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)][guid[]] $PrincipalObjectIds,
        [Parameter(Mandatory = $true)][ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType,
        [Parameter(Mandatory = $true)][string] $Scope
    )

    $roleName = Get-RjRbKvReadRoleName -ItemType $ItemType
    $assignments = @()

    foreach ($objectId in (@($PrincipalObjectIds) | Select-Object -Unique)) {
        $principalId = $objectId.ToString()

        $existing = Get-AzRoleAssignment -ObjectId $principalId -RoleDefinitionName $roleName -Scope $Scope -ErrorAction SilentlyContinue |
            Where-Object { $_.Scope -eq $Scope } | Select-Object -First 1

        if ($existing) {
            $assignments += [pscustomobject]@{
                PrincipalObjectId = $principalId
                RoleDefinition    = $roleName
                Scope             = $Scope
                Created           = $false
                AssignmentId      = $existing.RoleAssignmentId
            }
            continue
        }

        if ($PSCmdlet.ShouldProcess("$principalId at $Scope", "Assign '$roleName'")) {
            try {
                $new = New-AzRoleAssignment -ObjectId $principalId -RoleDefinitionName $roleName -Scope $Scope -ErrorAction Stop
                $assignments += [pscustomobject]@{
                    PrincipalObjectId = $principalId
                    RoleDefinition    = $roleName
                    Scope             = $Scope
                    Created           = $true
                    AssignmentId      = $new.RoleAssignmentId
                }
            }
            catch {
                # Tolerate the idempotent race: a concurrent run - or role-assignment read lag
                # that hid the assignment from the pre-check above - can make New-AzRoleAssignment
                # fail even though the desired assignment now exists. Re-query and treat that as
                # already granted; if the assignment still is not present, rethrow the original error.
                $already = Get-AzRoleAssignment -ObjectId $principalId -RoleDefinitionName $roleName -Scope $Scope -ErrorAction SilentlyContinue |
                    Where-Object { $_.Scope -eq $Scope } | Select-Object -First 1

                if (-not $already) {
                    throw
                }

                $assignments += [pscustomobject]@{
                    PrincipalObjectId = $principalId
                    RoleDefinition    = $roleName
                    Scope             = $Scope
                    Created           = $false
                    AssignmentId      = $already.RoleAssignmentId
                }
            }
        }
    }

    @($assignments)
}

# Waits for an asynchronous certificate issuance (create) to finish, then returns the
# issued certificate. Self-signed certs usually complete within seconds.
function Wait-RjRbKvCertificate {
    param(
        [Parameter(Mandatory = $true)][string] $VaultName,
        [Parameter(Mandatory = $true)][string] $Name,
        [int] $TimeoutSeconds = 180
    )

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($true) {
        $operation = Get-AzKeyVaultCertificateOperation -VaultName $VaultName -Name $Name -ErrorAction Stop
        if ($operation.Status -eq 'completed') { break }
        if (($operation.Status -eq 'failed') -or $operation.ErrorMessage) {
            throw "Certificate '$Name' issuance failed: $($operation.ErrorMessage)"
        }
        if ([DateTime]::UtcNow -ge $deadline) {
            throw "Certificate '$Name' was not issued within $TimeoutSeconds seconds (status '$($operation.Status)')."
        }
        Start-Sleep -Seconds 2
    }

    $certificate = Get-AzKeyVaultCertificate -VaultName $VaultName -Name $Name -ErrorAction Stop
    if (-not $certificate) {
        throw "Certificate '$Name' was reported issued but could not be read back."
    }
    $certificate
}

# Returns the "#@domain" (or "#@tenantId") segment used by Azure Portal deep links.
# The verified default domain yields nicer links; the tenant id is the fallback.
function Get-RjRbKvPortalTenantSegment {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    $tenantId = if ($ctx -and $ctx.Tenant) { $ctx.Tenant.Id } else { $null }

    if (Get-Command -Name Get-AzTenant -ErrorAction SilentlyContinue) {
        try {
            $tenant = Get-AzTenant -ErrorAction Stop | Where-Object { (-not $tenantId) -or ($_.Id -eq $tenantId) } | Select-Object -First 1
            $domain = $null
            if ($tenant) {
                if ($tenant.DefaultDomain) { $domain = $tenant.DefaultDomain }
                elseif (@($tenant.Domains).Count -gt 0) { $domain = @($tenant.Domains)[0] }
            }
            if ($domain) { return "#@$domain" }
        }
        catch {
            Write-RjRbLog "Could not resolve tenant default domain ($($_.Exception.Message)); using tenant id for portal URL."
        }
    }

    if ($tenantId) { return "#@$tenantId" }
    throw 'Could not determine a tenant hint (default domain or tenant id) for the portal URL.'
}

# Builds the data-plane URI and the portal deep links for a published item version.
function New-RjRbKvObjectUrl {
    param(
        [Parameter(Mandatory = $true)][string] $VaultName,
        [Parameter(Mandatory = $true)][string] $VaultResourceId,
        [Parameter(Mandatory = $true)][ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType,
        [Parameter(Mandatory = $true)][string] $ItemName,
        [string] $Version,
        [Parameter(Mandatory = $true)][string] $TenantSegment,
        [string] $ObjectScope
    )

    $segment = Get-RjRbKvItemSegment -ItemType $ItemType
    # Portal "asset" blade type per item kind (used for the exact-version deep link).
    $assetType = switch ($ItemType) {
        'Secret' { 'Secret' }
        'Certificate' { 'Certificate' }
        'Key' { 'Key' }
    }

    # Data-plane identifier of the exact version (or the logical object when not yet known).
    $versionUri = if ($Version) {
        "https://$VaultName.vault.azure.net/$segment/$ItemName/$Version"
    }
    else {
        "https://$VaultName.vault.azure.net/$segment/$ItemName"
    }

    # Portal deep link straight to the exact item version (e.g. "Show Secret Value" works here).
    $portalItemVersionUrl = 'https://portal.azure.com/{0}/asset/Microsoft_Azure_KeyVault/{1}/{2}' -f $TenantSegment, $assetType, $versionUri

    # Portal link to the logical object's versions blade (the RBAC/overview view), not a version.
    $objectId = "https://$VaultName.vault.azure.net/$segment/$ItemName"
    $encObjectId = [uri]::EscapeDataString($objectId)
    $encVaultId = [uri]::EscapeDataString($VaultResourceId)
    $portalObjectUrl = 'https://portal.azure.com/#view/Microsoft_Azure_KeyVault/ListObjectVersionsRBACBlade/~/overview/objectType/{0}/objectId/{1}/vaultResourceUri/{2}/vaultId/{2}/lifecycleState~/null' -f $segment, $encObjectId, $encVaultId

    # Portal link to the item's ARM resource metadata blade.
    $portalArmResourceUrl = $null
    if ($ObjectScope) {
        $portalArmResourceUrl = 'https://portal.azure.com/{0}/resource/{1}/overview' -f $TenantSegment, $ObjectScope.Trim('/')
    }

    [pscustomobject]@{
        ItemVersionUri       = $versionUri
        PortalItemVersionUrl = $portalItemVersionUrl
        PortalObjectUrl      = $portalObjectUrl
        PortalArmResourceUrl = $portalArmResourceUrl
    }
}

# Assembles the full, uniform result object shared by all three publish functions.
function New-RjRbKvPublishResult {
    param(
        [Parameter(Mandatory = $true)][string] $KeyVaultName,
        [Parameter(Mandatory = $true)] $Vault,
        [Parameter(Mandatory = $true)][ValidateSet('Secret', 'Certificate', 'Key')][string] $ItemType,
        [Parameter(Mandatory = $true)][string] $ItemName,
        $Item,
        [Parameter(Mandatory = $true)][string] $ObjectScope,
        [object[]] $Readers,
        [object[]] $RoleAssignments
    )

    $tenantSegment = Get-RjRbKvPortalTenantSegment
    # $Item is null under -WhatIf; fall back to the logical (unversioned) URLs.
    $version = if ($Item) { $Item.Version } else { $null }
    $urls = New-RjRbKvObjectUrl -VaultName $KeyVaultName -VaultResourceId $Vault.ResourceId -ItemType $ItemType `
        -ItemName $ItemName -Version $version -TenantSegment $tenantSegment -ObjectScope $ObjectScope

    [pscustomobject]@{
        KeyVaultName         = $KeyVaultName
        KeyVaultResourceId   = $Vault.ResourceId
        ItemType             = $ItemType
        ItemName             = $ItemName
        Version              = $version
        ItemVersionUri       = $urls.ItemVersionUri
        ObjectScope          = $ObjectScope
        PortalItemVersionUrl = $urls.PortalItemVersionUrl
        PortalObjectUrl      = $urls.PortalObjectUrl
        PortalArmResourceUrl = $urls.PortalArmResourceUrl
        Readers              = @($Readers)
        RoleAssignments      = @($RoleAssignments)
    }
}

# Shapes the function output from the full result per the caller's -Return selection:
# a single value -> scalar; several values -> subset object; 'All' -> full object.
function Select-RjRbKvReturnValue {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $Result,
        [Parameter(Mandatory = $true)][string[]] $Return
    )

    if ((@($Return).Count -eq 1) -and ($Return -notcontains 'All')) {
        return $Result.($Return[0])
    }
    if ($Return -contains 'All') {
        return $Result
    }
    $selected = [ordered]@{}
    foreach ($name in $Return) { $selected[$name] = $Result.$name }
    [pscustomobject]$selected
}

function Publish-RjRbKeyVaultSecret {
    <#
        .SYNOPSIS
        Push a secret into an Azure Key Vault and grant object-scoped ("per secret") RBAC
        read access to one or more principals.

        .DESCRIPTION
        Stores (creates or updates) a secret in the target Key Vault, then assigns the
        built-in 'Key Vault Secrets User' role on the secret object scope only
        (.../vaults/<vault>/secrets/<name>) to the supplied reader principals. No access is
        granted at the vault, resource group, or subscription scope, so a value can be handed
        to specific principals without exposing the rest of the vault.

        The target vault must use the Azure RBAC permission model. The secret value is never
        written to output or the RealmJoin/runbook log. By default the function returns the
        portal deep link to the exact secret version - the most useful artifact to send to a
        user. Use -Return to select other attributes, or 'All' for the full result object.

        Required Azure modules in the consuming runbook (declare via #Requires):
        Az.Accounts, Az.KeyVault, Az.Resources.

        .PARAMETER KeyVaultName
        Target Key Vault name.

        .PARAMETER SecretName
        Name of the secret to create or update.

        .PARAMETER SecretValue
        Plain-text value to store. Never logged. Empty strings are allowed by the parameter,
        but the Key Vault service itself rejects empty secret values.

        .PARAMETER ReaderUsers
        Principals that should be able to read this exact secret. Accepts an array and/or
        comma/semicolon/newline separated user UPNs/mail addresses, or exact display names /
        object ids for users, groups, and service principals. Optional - omit to only push the
        secret without granting access.

        .PARAMETER KeyVaultResourceGroupName
        Resource group of the vault. Optional, but disambiguates same-named lookups.

        .PARAMETER SubscriptionId
        Optional subscription to switch the Az context to before the operation.

        .PARAMETER Tag
        Optional extra tags to set on the secret (merged over the default Source tag).

        .PARAMETER Return
        Which attribute(s) to output. The default 'PortalItemVersionUrl' returns a single
        string. Pass several values, or 'All', to return a structured object instead.

        .EXAMPLE
        Publish-RjRbKeyVaultSecret -KeyVaultName 'kv-transfer' -SecretName 'WifiPwd' `
            -SecretValue $value -ReaderUsers 'jane@contoso.com'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $KeyVaultName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $SecretName,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string] $SecretValue,
        [string[]] $ReaderUsers = @(),
        [string] $KeyVaultResourceGroupName,
        [string] $SubscriptionId,
        [hashtable] $Tag,
        [ValidateSet('PortalItemVersionUrl', 'PortalObjectUrl', 'PortalArmResourceUrl',
            'ItemVersionUri', 'ObjectScope', 'Version', 'RoleAssignments', 'Readers', 'All')]
        [string[]] $Return = @('PortalItemVersionUrl')
    )

    $vault = Get-RjRbKvValidatedTargetVault -RequiredCmdlets 'Get-AzContext', 'Get-AzKeyVault', 'Set-AzKeyVaultSecret', 'Get-AzRoleAssignment', 'New-AzRoleAssignment', 'Get-AzADUser', 'Get-AzADGroup', 'Get-AzADServicePrincipal' `
        -KeyVaultName $KeyVaultName -KeyVaultResourceGroupName $KeyVaultResourceGroupName -SubscriptionId $SubscriptionId

    # Resolve readers up front so we never push a secret we then cannot share.
    $readerInfo = Get-RjRbKvReader -ReaderUsers $ReaderUsers
    $readers = @($readerInfo.Readers)
    $readerIds = $readerInfo.ObjectIds

    $effectiveTag = New-RjRbKvEffectiveTag -ItemType Secret -Tag $Tag
    # Convert to a SecureString only at the boundary; keep the value out of logs/output.
    if ($SecretValue.Length -eq 0) {
        throw "SecretValue cannot be empty. Azure Key Vault rejects empty secret values."
    }
    $secureValue = ConvertTo-SecureString -String $SecretValue -AsPlainText -Force

    $item = $null
    if ($PSCmdlet.ShouldProcess("secret '$SecretName' in vault '$KeyVaultName'", 'Set secret and grant object-scoped read access')) {
        Write-RjRbLog "Setting secret '$SecretName' in Key Vault '$KeyVaultName' (value not logged)"
        $item = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $SecretName -SecretValue $secureValue -Tag $effectiveTag -ErrorAction Stop
    }

    $objectScope = Get-RjRbKvObjectScope -VaultResourceId $vault.ResourceId -ItemType Secret -ItemName $SecretName
    $roleAssignments = @()
    if ($item -and (@($readerIds).Count -gt 0)) {
        $roleAssignments = @(Grant-RjRbKvObjectAccess -PrincipalObjectIds $readerIds -ItemType Secret -Scope $objectScope)
    }

    $result = New-RjRbKvPublishResult -KeyVaultName $KeyVaultName -Vault $vault -ItemType Secret -ItemName $SecretName `
        -Item $item -ObjectScope $objectScope -Readers $readers -RoleAssignments $roleAssignments
    Select-RjRbKvReturnValue -Result $result -Return $Return
}

function Publish-RjRbKeyVaultKey {
    <#
        .SYNOPSIS
        Create or import a key in an Azure Key Vault and grant object-scoped ("per key") RBAC
        use access to one or more principals.

        .DESCRIPTION
        Without -KeyFilePath a new key is generated in the vault; with -KeyFilePath an existing
        key is imported from a .pfx/.byok/.pem file. Either way the built-in 'Key Vault Crypto
        User' role is then assigned on the key object scope only
        (.../vaults/<vault>/keys/<name>) to the supplied readers. That role lets a principal use
        the key for crypto operations (sign/verify, encrypt/decrypt, wrap/unwrap) and read the
        public key; it does not allow exporting the private key.

        The target vault must use the Azure RBAC permission model. By default the function
        returns the portal deep link to the exact key version. Use -Return for other attributes.

        Required Azure modules in the consuming runbook (declare via #Requires):
        Az.Accounts, Az.KeyVault, Az.Resources.

        .PARAMETER KeyVaultName
        Target Key Vault name.

        .PARAMETER KeyName
        Name of the key to create or import.

        .PARAMETER KeyFilePath
        Path to a key file (.pfx/.byok/.pem) to import. When set, the key is imported instead
        of generated.

        .PARAMETER KeyFilePassword
        Password (SecureString) protecting the import file, if any.

        .PARAMETER KeyType
        Key type for a newly generated key: 'RSA' (default) or 'EC'.

        .PARAMETER Size
        RSA key size in bits (2048, 3072, or 4096) for a newly generated RSA key.

        .PARAMETER Curve
        Elliptic curve name for a newly generated EC key.

        .PARAMETER Destination
        'Software' (default) or 'HSM'. Applies to both create and import.

        .PARAMETER KeyOperation
        Optional list of permitted key operations (e.g. sign, verify, wrapKey, unwrapKey).

        .PARAMETER Expires
        Optional expiry (DateTime) for the key.

        .PARAMETER NotBefore
        Optional "not before" (DateTime) for the key.

        .PARAMETER ReaderUsers
        Principals to grant object-scoped key use access. Same formats as the secret cmdlet.

        .PARAMETER KeyVaultResourceGroupName
        Resource group of the vault. Optional.

        .PARAMETER SubscriptionId
        Optional subscription to switch the Az context to before the operation.

        .PARAMETER Tag
        Optional extra tags (merged over the default Source tag).

        .PARAMETER Return
        Which attribute(s) to output. Default 'PortalItemVersionUrl' (a single string).

        .EXAMPLE
        Publish-RjRbKeyVaultKey -KeyVaultName 'kv-transfer' -KeyName 'signing' -KeyType RSA -Size 3072 -ReaderUsers 'app-sp-object-id'

        .EXAMPLE
        Publish-RjRbKeyVaultKey -KeyVaultName 'kv-transfer' -KeyName 'imported' -KeyFilePath 'C:\tmp\key.pfx' -KeyFilePassword $pwd
    #>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Create')]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $KeyVaultName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $KeyName,
        [Parameter(Mandatory = $true, ParameterSetName = 'Import')][string] $KeyFilePath,
        [Parameter(ParameterSetName = 'Import')][securestring] $KeyFilePassword,
        [Parameter(ParameterSetName = 'Create')][ValidateSet('RSA', 'EC')][string] $KeyType = 'RSA',
        [Parameter(ParameterSetName = 'Create')][ValidateSet(2048, 3072, 4096)][int] $Size,
        [Parameter(ParameterSetName = 'Create')][ValidateSet('P-256', 'P-384', 'P-521', 'P-256K')][string] $Curve,
        [ValidateSet('Software', 'HSM')][string] $Destination = 'Software',
        [string[]] $KeyOperation,
        [datetime] $Expires,
        [datetime] $NotBefore,
        [string[]] $ReaderUsers = @(),
        [string] $KeyVaultResourceGroupName,
        [string] $SubscriptionId,
        [hashtable] $Tag,
        [ValidateSet('PortalItemVersionUrl', 'PortalObjectUrl', 'PortalArmResourceUrl',
            'ItemVersionUri', 'ObjectScope', 'Version', 'RoleAssignments', 'Readers', 'All')]
        [string[]] $Return = @('PortalItemVersionUrl')
    )

    $vault = Get-RjRbKvValidatedTargetVault -RequiredCmdlets 'Get-AzContext', 'Get-AzKeyVault', 'Add-AzKeyVaultKey', 'Get-AzRoleAssignment', 'New-AzRoleAssignment', 'Get-AzADUser', 'Get-AzADGroup', 'Get-AzADServicePrincipal' `
        -KeyVaultName $KeyVaultName -KeyVaultResourceGroupName $KeyVaultResourceGroupName -SubscriptionId $SubscriptionId

    $readerInfo = Get-RjRbKvReader -ReaderUsers $ReaderUsers
    $readers = @($readerInfo.Readers)
    $readerIds = $readerInfo.ObjectIds

    $effectiveTag = New-RjRbKvEffectiveTag -ItemType Key -Tag $Tag
    $isImport = $PSCmdlet.ParameterSetName -eq 'Import'

    # Build the Add-AzKeyVaultKey arguments for the chosen mode (import vs. generate).
    $keyArgs = @{ VaultName = $KeyVaultName; Name = $KeyName; Destination = $Destination; Tag = $effectiveTag; ErrorAction = 'Stop' }
    if ($isImport) {
        if (-not (Test-Path -LiteralPath $KeyFilePath -PathType Leaf)) {
            throw "Key import file '$KeyFilePath' was not found."
        }
        $keyArgs['KeyFilePath'] = $KeyFilePath
        if ($KeyFilePassword) { $keyArgs['KeyFilePassword'] = $KeyFilePassword }
    }
    else {
        $keyArgs['KeyType'] = $KeyType
        if ($Size) { $keyArgs['Size'] = $Size }
        if ($Curve) { $keyArgs['CurveName'] = $Curve }
        if ($KeyOperation) { $keyArgs['KeyOps'] = $KeyOperation }
    }
    if ($Expires) { $keyArgs['Expires'] = $Expires }
    if ($NotBefore) { $keyArgs['NotBefore'] = $NotBefore }

    $verb = if ($isImport) { 'Importing' } else { 'Creating' }
    $action = if ($isImport) { 'Import key and grant object-scoped use access' } else { 'Create key and grant object-scoped use access' }

    $item = $null
    if ($PSCmdlet.ShouldProcess("key '$KeyName' in vault '$KeyVaultName'", $action)) {
        Write-RjRbLog "$verb key '$KeyName' in Key Vault '$KeyVaultName'"
        $item = Add-AzKeyVaultKey @keyArgs
    }

    $objectScope = Get-RjRbKvObjectScope -VaultResourceId $vault.ResourceId -ItemType Key -ItemName $KeyName
    $roleAssignments = @()
    if ($item -and (@($readerIds).Count -gt 0)) {
        $roleAssignments = @(Grant-RjRbKvObjectAccess -PrincipalObjectIds $readerIds -ItemType Key -Scope $objectScope)
    }

    $result = New-RjRbKvPublishResult -KeyVaultName $KeyVaultName -Vault $vault -ItemType Key -ItemName $KeyName `
        -Item $item -ObjectScope $objectScope -Readers $readers -RoleAssignments $roleAssignments
    Select-RjRbKvReturnValue -Result $result -Return $Return
}

function Publish-RjRbKeyVaultCertificate {
    <#
        .SYNOPSIS
        Create (self-signed/issuer) or import a certificate in an Azure Key Vault and grant
        object-scoped ("per certificate") RBAC read access to one or more principals.

        .DESCRIPTION
        Three modes, selected by the parameters used:
        - Create (convenience): pass -SubjectName (and optionally -DnsName, -IssuerName,
          -ValidityInMonths) to issue a new certificate from a generated policy. -IssuerName
          defaults to 'Self' (self-signed).
        - Create (advanced): pass a ready -CertificatePolicy built with
          New-AzKeyVaultCertificatePolicy.
        - Import: pass -CertificateFilePath (.pfx/.pem) and optional -CertificateFilePassword.

        The built-in 'Key Vault Certificate User' role is then assigned on the certificate
        object scope only (.../vaults/<vault>/certificates/<name>) to the supplied readers.
        To also let readers download the certificate WITH its private key (PFX), pass
        -GrantPrivateKeyAccess, which additionally grants 'Key Vault Secrets User' on the
        certificate's backing secret object (same name under .../secrets/<name>).

        The target vault must use the Azure RBAC permission model. Certificate creation is
        asynchronous; the function waits for issuance to complete. By default it returns the
        portal deep link to the exact certificate version. Use -Return for other attributes.

        Required Azure modules in the consuming runbook (declare via #Requires):
        Az.Accounts, Az.KeyVault, Az.Resources.

        .PARAMETER KeyVaultName
        Target Key Vault name.

        .PARAMETER CertificateName
        Name of the certificate to create or import.

        .PARAMETER CertificateFilePath
        Path to a .pfx/.pem file to import. When set, the certificate is imported.

        .PARAMETER CertificateFilePassword
        Password (SecureString) protecting the import file, if any.

        .PARAMETER CertificatePolicy
        A certificate policy object (from New-AzKeyVaultCertificatePolicy) for advanced create.

        .PARAMETER SubjectName
        X.500 subject for a convenience create, e.g. 'CN=transfer.contoso.com'.

        .PARAMETER DnsName
        Optional Subject Alternative Name DNS entries for a convenience create.

        .PARAMETER IssuerName
        Issuer for a convenience create. Defaults to 'Self' (self-signed).

        .PARAMETER ValidityInMonths
        Validity in months for a convenience create (default 12).

        .PARAMETER GrantPrivateKeyAccess
        Also grant readers 'Key Vault Secrets User' on the certificate's backing secret so they
        can download the full certificate including its private key (PFX).

        .PARAMETER IssuanceTimeoutSeconds
        How long to wait for certificate creation to complete (default 180).

        .PARAMETER ReaderUsers
        Principals to grant object-scoped certificate access. Same formats as the secret cmdlet.

        .PARAMETER KeyVaultResourceGroupName
        Resource group of the vault. Optional.

        .PARAMETER SubscriptionId
        Optional subscription to switch the Az context to before the operation.

        .PARAMETER Tag
        Optional extra tags (merged over the default Source tag).

        .PARAMETER Return
        Which attribute(s) to output. Default 'PortalItemVersionUrl' (a single string).

        .EXAMPLE
        Publish-RjRbKeyVaultCertificate -KeyVaultName 'kv-transfer' -CertificateName 'web' `
            -SubjectName 'CN=web.contoso.com' -DnsName 'web.contoso.com' -ReaderUsers 'jane@contoso.com'

        .EXAMPLE
        Publish-RjRbKeyVaultCertificate -KeyVaultName 'kv-transfer' -CertificateName 'imported' `
            -CertificateFilePath 'C:\tmp\cert.pfx' -CertificateFilePassword $pwd -GrantPrivateKeyAccess -ReaderUsers 'jane@contoso.com'
    #>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Create')]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $KeyVaultName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $CertificateName,
        [Parameter(Mandatory = $true, ParameterSetName = 'Import')][string] $CertificateFilePath,
        [Parameter(ParameterSetName = 'Import')][securestring] $CertificateFilePassword,
        [Parameter(Mandatory = $true, ParameterSetName = 'CreatePolicy')] $CertificatePolicy,
        [Parameter(Mandatory = $true, ParameterSetName = 'Create')][string] $SubjectName,
        [Parameter(ParameterSetName = 'Create')][string[]] $DnsName,
        [Parameter(ParameterSetName = 'Create')][string] $IssuerName = 'Self',
        [Parameter(ParameterSetName = 'Create')][int] $ValidityInMonths = 12,
        [switch] $GrantPrivateKeyAccess,
        [int] $IssuanceTimeoutSeconds = 180,
        [string[]] $ReaderUsers = @(),
        [string] $KeyVaultResourceGroupName,
        [string] $SubscriptionId,
        [hashtable] $Tag,
        [ValidateSet('PortalItemVersionUrl', 'PortalObjectUrl', 'PortalArmResourceUrl',
            'ItemVersionUri', 'ObjectScope', 'Version', 'RoleAssignments', 'Readers', 'All')]
        [string[]] $Return = @('PortalItemVersionUrl')
    )

    $requiredCmdlets = @(
        'Get-AzContext',
        'Get-AzKeyVault',
        'Get-AzRoleAssignment',
        'New-AzRoleAssignment',
        'Get-AzADUser',
        'Get-AzADGroup',
        'Get-AzADServicePrincipal'
    )
    if ($PSCmdlet.ParameterSetName -eq 'Import') {
        $requiredCmdlets += 'Import-AzKeyVaultCertificate'
    }
    else {
        $requiredCmdlets += 'Add-AzKeyVaultCertificate', 'Get-AzKeyVaultCertificateOperation', 'Get-AzKeyVaultCertificate'
        if ($PSCmdlet.ParameterSetName -eq 'Create') {
            $requiredCmdlets += 'New-AzKeyVaultCertificatePolicy'
        }
    }

    $vault = Get-RjRbKvValidatedTargetVault -RequiredCmdlets $requiredCmdlets `
        -KeyVaultName $KeyVaultName -KeyVaultResourceGroupName $KeyVaultResourceGroupName -SubscriptionId $SubscriptionId

    $readerInfo = Get-RjRbKvReader -ReaderUsers $ReaderUsers
    $readers = @($readerInfo.Readers)
    $readerIds = $readerInfo.ObjectIds

    $effectiveTag = New-RjRbKvEffectiveTag -ItemType Certificate -Tag $Tag
    $isImport = $PSCmdlet.ParameterSetName -eq 'Import'

    $item = $null
    if ($isImport) {
        if (-not (Test-Path -LiteralPath $CertificateFilePath -PathType Leaf)) {
            throw "Certificate import file '$CertificateFilePath' was not found."
        }
        if ($PSCmdlet.ShouldProcess("certificate '$CertificateName' in vault '$KeyVaultName'", 'Import certificate and grant object-scoped read access')) {
            Write-RjRbLog "Importing certificate '$CertificateName' into Key Vault '$KeyVaultName'"
            $importArgs = @{ VaultName = $KeyVaultName; Name = $CertificateName; FilePath = $CertificateFilePath; Tag = $effectiveTag; ErrorAction = 'Stop' }
            if ($CertificateFilePassword) { $importArgs['Password'] = $CertificateFilePassword }
            $item = Import-AzKeyVaultCertificate @importArgs
        }
    }
    else {
        # Use the supplied policy, or build a convenience one from SubjectName/DnsName/etc.
        $policy = if ($PSCmdlet.ParameterSetName -eq 'CreatePolicy') {
            $CertificatePolicy
        }
        else {
            $policyArgs = @{ SubjectName = $SubjectName; IssuerName = $IssuerName; ValidityInMonths = $ValidityInMonths; SecretContentType = 'application/x-pkcs12' }
            if ($DnsName) { $policyArgs['DnsName'] = $DnsName }
            New-AzKeyVaultCertificatePolicy @policyArgs
        }

        if ($PSCmdlet.ShouldProcess("certificate '$CertificateName' in vault '$KeyVaultName'", 'Create certificate and grant object-scoped read access')) {
            $issuerLabel = if ($PSCmdlet.ParameterSetName -eq 'CreatePolicy') { $null } else { $IssuerName }
            $logMessage =  "Creating certificate '$CertificateName' in Key Vault '$KeyVaultName' (issuer '$issuerLabel')"

            Write-RjRbLog $logMessage
            Add-AzKeyVaultCertificate -VaultName $KeyVaultName -Name $CertificateName -CertificatePolicy $policy -Tag $effectiveTag -ErrorAction Stop | Out-Null
            # Creation is asynchronous - wait for issuance and read the issued certificate back.
            $item = Wait-RjRbKvCertificate -VaultName $KeyVaultName -Name $CertificateName -TimeoutSeconds $IssuanceTimeoutSeconds
        }
    }

    $objectScope = Get-RjRbKvObjectScope -VaultResourceId $vault.ResourceId -ItemType Certificate -ItemName $CertificateName
    $roleAssignments = @()
    if ($item -and (@($readerIds).Count -gt 0)) {
        $roleAssignments += @(Grant-RjRbKvObjectAccess -PrincipalObjectIds $readerIds -ItemType Certificate -Scope $objectScope)

        # Optionally also grant read on the backing secret so readers can pull the PFX (private key).
        if ($GrantPrivateKeyAccess) {
            $secretScope = Get-RjRbKvObjectScope -VaultResourceId $vault.ResourceId -ItemType Secret -ItemName $CertificateName
            Write-RjRbLog "Also granting backing-secret access for private-key (PFX) download at '$secretScope'"
            $roleAssignments += @(Grant-RjRbKvObjectAccess -PrincipalObjectIds $readerIds -ItemType Secret -Scope $secretScope)
        }
    }

    $result = New-RjRbKvPublishResult -KeyVaultName $KeyVaultName -Vault $vault -ItemType Certificate -ItemName $CertificateName `
        -Item $item -ObjectScope $objectScope -Readers $readers -RoleAssignments $roleAssignments
    Select-RjRbKvReturnValue -Result $result -Return $Return
}
