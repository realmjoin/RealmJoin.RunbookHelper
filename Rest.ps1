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

    $invokeArguments = rjRbGetParametersFiltered -exclude 'Beta', 'ReturnValueProperty', 'FollowPaging'

    $invokeArguments['Uri'] = "https://graph.microsoft.com/$(if($Beta) {'beta'} else {'v1.0'})"
    if (-not $Headers -and (Test-Path Variable:Script:RjRbGraphAuthHeaders)) {
        $invokeArguments['Headers'] = $Script:RjRbGraphAuthHeaders
    }
    if (-not ($Body -is [byte[]] -or $Body -is [IO.Stream])) {
        $invokeArguments['JsonEncodeBody'] = $true
    }

    $result = Invoke-RjRbRestMethod @invokeArguments
    if ($null -ne $result) {

        if ($FollowPaging -and $result.PSObject.Properties['value']) {
            # successively release results to PS pipeline
            Write-Output $result.value
            $invokeNextLinkArguments = rjRbGetParametersFiltered -sourceValues $invokeArguments -include 'Method', 'Headers'
            while ($result.PSObject.Properties['@odata.nextLink'] -and $result.PSObject.Properties['value']) {
                $invokeNextLinkArguments['Uri'] = $result.'@odata.nextLink'
                $result = Invoke-RjRbRestMethod @invokeNextLinkArguments
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
        [Management.Automation.ActionPreference] $NotFoundAction
    )

    $invokeArguments = rjRbGetParametersFiltered -exclude 'UriSuffix', 'UriQueryParam', 'UriQueryRaw', 'OdFilter', 'OdSelect', 'OdTop', 'JsonEncodeBody', 'NotFoundAction'

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

    if ($UriSuffix) { $uriBuilder.Path += $UriSuffix }
    if ($UriQueryRaw) { appendToQuery $UriQueryRaw }
    $UriQueryParam | Foreach-Object { appendToQuery $_ -split }
    $PSBoundParameters.Keys -ilike 'Od*' | Foreach-Object { appendToQuery "`$$($_.Substring(2).ToLower())" $PSBoundParameters[$_] -skipEmpty }
    if ($Body -and $JsonEncodeBody) {
        # need to explicetly set charset in ContenType for Invoke-RestMethod to detect it and to correctly encode JSON string
        $invokeArguments['ContentType'] = "application/json; charset=UTF-8"
        $invokeArguments['Body'] = $Body | ConvertTo-Json
    }
    if ($InFile -and -not $ContentType) {
        $invokeArguments['ContentType'] = [Web.MimeMapping]::GetMimeMapping($InFile)
    }

    # remove empty string parameters since they will never be $null but empty only
    @('InFile', 'ContentType') | Where-Object { $invokeArguments.ContainsKey($_) -and $invokeArguments[$_] -eq [string]::Empty } | `
        ForEach-Object { $invokeArguments.Remove($_) }

    $invokeArguments['Uri'] = $uriBuilder.Uri
    $invokeArguments['UseBasicParsing'] = $true

    Write-RjRbDebug "Invoke-RestMethod arguments" $invokeArguments
    $result = $null # Write-Error down below might not be terminating
    try {
        $result = Invoke-RestMethod @invokeArguments
    }
    catch {
        $isWebException = $_.Exception -is [Net.WebException]
        $isNotFound = $isWebException -and $_.Exception.Response.StatusCode -eq [Net.HttpStatusCode]::NotFound
        $errorAction = $(if ($isNotFound -and $null -ne $NotFoundAction) { $NotFoundAction } else { $ErrorActionPreference })

        # no need to write error details to log on SilentlyContinue or Ignore
        if ($errorAction -notin @([Management.Automation.ActionPreference]::SilentlyContinue, [Management.Automation.ActionPreference]::Ignore)) {

            Write-RjRbLog "Invoke-RestMethod arguments" $invokeArguments -NoDebugOnly

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

    Write-RjRbDebug "Invoke-RestMethod result" $result

    return $result
}
