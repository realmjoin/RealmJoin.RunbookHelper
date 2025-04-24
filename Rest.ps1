function Invoke-RjRbRestMethodGraph {
    [CmdletBinding()]
    param (
        [string] $Resource,
        [string[]] $UriQueryParam = @(),
        [string] $UriQueryRaw,
        [string] $OdFilter,
        [string] $OdSelect,
        [int] $OdTop,
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Default,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [string] $InFile,
        [string] $ContentType,
        [switch] $Beta,
        [Nullable[bool]] $ReturnValueProperty,
        [switch] $FollowPaging,
        [Management.Automation.ActionPreference] $NotFoundAction
    )

    $invokeArguments = rjRbGetParametersFiltered -exclude 'Beta'

    $invokeArguments['Uri'] = "https://graph.microsoft.com/$(if($Beta) {'beta'} else {'v1.0'})"

    if (-not $Headers -and (Test-Path Variable:Script:RjRbGraphAuthHeaders)) {
        $invokeArguments['Headers'] = $Script:RjRbGraphAuthHeaders
    }

    Invoke-RjRbRestMethod -JsonEncodeBody @invokeArguments
}

function Invoke-RjRbRestMethodDefenderATP {
    [CmdletBinding()]
    param (
        [string] $Resource,
        [string[]] $UriQueryParam = @(),
        [string] $UriQueryRaw,
        [string] $OdFilter,
        [string] $OdSelect,
        [int] $OdTop,
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Default,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [string] $InFile,
        [string] $ContentType,
        [Nullable[bool]] $ReturnValueProperty,
        [switch] $FollowPaging,
        [Management.Automation.ActionPreference] $NotFoundAction
    )

    $invokeArguments = rjRbGetParametersFiltered

    $invokeArguments['Uri'] = "https://api.securitycenter.microsoft.com/api"

    if (-not $Headers -and (Test-Path Variable:Script:RjRbDefenderATPAuthHeaders)) {
        $invokeArguments['Headers'] = $Script:RjRbDefenderATPAuthHeaders
    }

    Invoke-RjRbRestMethod -JsonEncodeBody @invokeArguments
}

function Invoke-RjRbRestMethod {
    [CmdletBinding()]
    param (
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Default,
        [uri] $Uri,
        [Alias('Resource')][string] $UriSuffix,
        [string[]] $UriQueryParam = @(),
        [string] $UriQueryRaw,
        [string] $OdFilter,
        [string] $OdSelect,
        [int] $OdTop,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [switch] $JsonEncodeBody,
        [string] $InFile,
        [string] $ContentType,
        [Management.Automation.ActionPreference] $NotFoundAction,
        [switch] $FollowPaging,
        [Nullable[bool]] $ReturnValueProperty
    )

    $invokeArguments = rjRbGetParametersFiltered -exclude 'UriSuffix', 'UriQueryParam', 'UriQueryRaw', 'OdFilter', 'OdSelect', 'OdTop', 'FollowPaging', 'ReturnValueProperty'

    $uriBuilder = [UriBuilder]::new($Uri)
    function appendToQuery([string] $newQueryOrParamName, [object] $paramValue <# [string] would never be $null #>, [switch] $split, [switch] $skipEmpty) {
        if ($split) {
            $splitPos = $newQueryOrParamName.IndexOf('=')
            $paramValue = $newQueryOrParamName.Substring($splitPos + 1)
            $newQueryOrParamName = $newQueryOrParamName.Substring(0, $splitPos)
        }
        if ($skipEmpty -and (-not $paramValue)) {
            return
        }
        if ($null -ne $paramValue) {
            $newQueryOrParamName += "=$([Web.HttpUtility]::UrlEncode($paramValue))"
        }
        if (-not $uriBuilder.Query) {
            $uriBuilder.Query = $newQueryOrParamName
        }
        else {
            $uriBuilder.Query = $uriBuilder.Query.TrimStart('?') + "&" + $newQueryOrParamName
        }
    }

    if ($UriSuffix) { $uriBuilder.Path += $UriSuffix }
    if ($UriQueryRaw) { appendToQuery $UriQueryRaw }
    $UriQueryParam | Foreach-Object { appendToQuery $_ -split }
    $PSBoundParameters.Keys -ilike 'Od*' | Foreach-Object { appendToQuery "`$$($_.Substring(2).ToLower())" $PSBoundParameters[$_] -skipEmpty }

    $invokeArguments['Uri'] = $uriBuilder.Uri

    $result = invokeRjRbRestMethodInternal @invokeArguments

    if ($null -ne $result) {

        if ($FollowPaging -and $result.PSObject.Properties['value']) {
            # successively release results to PS pipeline
            Write-Output $result.value
            $invokeNextLinkArguments = rjRbGetParametersFiltered -sourceValues $invokeArguments -include 'Method', 'Headers'
            while ($result.PSObject.Properties['@odata.nextLink'] -and $result.PSObject.Properties['value']) {
                $invokeNextLinkArguments['Uri'] = $result.'@odata.nextLink'
                $result = invokeRjRbRestMethodInternal @invokeNextLinkArguments
                Write-Output $result.value
            }
            return # result has already been return using Write-Output
        }

        if (($ReturnValueProperty -eq $true) -or (($ReturnValueProperty -ne $false) -and $result.PSObject.Properties['value'])) {
            $result = $result.value
        }

    }

    return $result
}

function invokeRjRbRestMethodInternal {
    [CmdletBinding()]
    param (
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Default,
        [uri] $Uri,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [switch] $JsonEncodeBody,
        [string] $InFile,
        [string] $ContentType,
        [int] $ThrottleMaxTries = 3,
        [Management.Automation.ActionPreference] $NotFoundAction
    )

    $invokeArguments = rjRbGetParametersFiltered -exclude 'JsonEncodeBody', 'ThrottleMaxTries', 'NotFoundAction'

    if ($Method -eq [Microsoft.PowerShell.Commands.WebRequestMethod]::Default) {
        if ($Body) {
            $invokeArguments['Method'] = [Microsoft.PowerShell.Commands.WebRequestMethod]::Post
        }
        elseif ($InFile) {
            $invokeArguments['Method'] = [Microsoft.PowerShell.Commands.WebRequestMethod]::Put
        }
        else {
            $invokeArguments['Method'] = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get
        }
    }

    if ($JsonEncodeBody -and $Body -and (-not ($Body -is [byte[]] -or $Body -is [IO.Stream]))) {
        # need to explicetly set charset in ContentType for Invoke-RestMethod to detect it and to correctly encode JSON string
        $invokeArguments['ContentType'] = "application/json; charset=UTF-8"
        $invokeArguments['Body'] = $Body | ConvertTo-Json -Depth 20
    }
    if ($InFile -and -not $ContentType) {
        $invokeArguments['ContentType'] = [Web.MimeMapping]::GetMimeMapping($InFile)
    }

    # remove empty string parameters since they will never be $null but empty only
    @('InFile', 'ContentType') | Where-Object { $invokeArguments.ContainsKey($_) -and $invokeArguments[$_] -eq [string]::Empty } | `
        ForEach-Object { $invokeArguments.Remove($_) }

    $invokeArguments['UseBasicParsing'] = $true

    Write-RjRbDebug "Invoke-RestMethod arguments" $invokeArguments

    $tryCount = 0
    do {

        $tryCount++
        $result = $null # Write-Error down below might not be terminating
        try {
            $result = Invoke-RestMethod @invokeArguments
        }

        catch {
            $isWebException = $_.Exception -is [Net.WebException]

            $isThrottled = $isWebException -and $_.Exception.Response.StatusCode -eq 429 # .NET 4.7 does not (yet) contain [Net.HttpStatusCode]::TooManyRequests
            if ($isThrottled -and $tryCount -lt $ThrottleMaxTries) {
                $retryAfter = [double]$_.Exception.Response.Headers["Retry-After"]
                if (-not $retryAfter) { $retryAfter = 15 }

                Write-RjRbLog "Request has been throttled (http status 429). Delaying for $retryAfter seconds and then trying again (this was try $tryCount of $ThrottleMaxTries)."
                Start-Sleep -Seconds $retryAfter

                continue # retry
            }

            $isNotFound = $isWebException -and $_.Exception.Response.StatusCode -eq [Net.HttpStatusCode]::NotFound
            $errorAction = $(if ($isNotFound -and $null -ne $NotFoundAction) { $NotFoundAction } else { $ErrorActionPreference })

            # no need to write error details to log on SilentlyContinue or Ignore
            if ($errorAction -notin @([Management.Automation.ActionPreference]::SilentlyContinue, [Management.Automation.ActionPreference]::Ignore)) {

                # avoid dumping full credentials outside of debug (use .Clone() to ensure to _not_ modify args for subsequent uses)
                $invokeArgsSanitized = $invokeArguments.Clone()
                if ($invokeArgsSanitized['Headers'] -and $invokeArgsSanitized['Headers']['Authorization']) {
                    $invokeArgsSanitized['Headers'] = $invokeArgsSanitized['Headers'].Clone()
                    $invokeArgsSanitized['Headers']['Authorization'] = $invokeArgsSanitized['Headers']['Authorization'] -replace '(?s)(?<=^\S+ \S{8}).*$', '...'
                }
                Write-RjRbLog "Invoke-RestMethod arguments" $invokeArgsSanitized -NoDebugOnly

                # get error response if available
                if ($isWebException) {
                    $errorResponse = $null; $responseReader = $null
                    try {
                        $responseStream = $_.Exception.Response.GetResponseStream()
                        if ($responseStream) {
                            $responseReader = [IO.StreamReader]::new($responseStream)
                            $errorResponse = $responseReader.ReadToEnd()
                            $errorResponse = $errorResponse | ConvertFrom-Json
                        }
                    }
                    catch { } # ignore all errors
                    finally {
                        if ($responseReader) {
                            $responseReader.Close()
                        }
                    }
                    Write-RjRbLog "Invoke-RestMethod error response" $errorResponse
                }
            }

            Write-Error -ErrorRecord $_ -ErrorAction $errorAction
        }

        break # always break since retries due to throttling will have already been handled above

    } while ($true) # need this for 'continue' to work

    Write-RjRbDebug "Invoke-RestMethod result" $result

    return $result
}
