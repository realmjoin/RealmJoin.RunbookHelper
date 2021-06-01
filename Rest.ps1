function Invoke-RjRbRestMethodGraph {
    [CmdletBinding()]
    param (
        [string] $Resource,
        [string[]] $UriQueryParam = @(),
        [string] $UriQueryRaw,
        [string] $OdFilter,
        [string] $OdSelect,
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Default,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [switch] $Beta
    )

    $invokeParameters = rjRbGetParametersFiltered -exclude 'Beta'

    $invokeParameters['Uri'] = "https://graph.microsoft.com/$(if($Beta) {'beta'} else {'v1.0'})"
    if (-not $Headers -and (Test-Path Variable:Script:RjRbGraphAuthHeaders)) {
        $invokeParameters['Headers'] = $Script:RjRbGraphAuthHeaders
    }
    $invokeParameters['JsonEncodeBody'] = $true

    Invoke-RjRbRestMethod @invokeParameters | Select-Object -ExpandProperty value
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
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [switch] $JsonEncodeBody
    )

    $invokeParameters = rjRbGetParametersFiltered -exclude 'UriSuffix', 'UriQueryParam', 'UriQueryRaw', 'OdFilter', 'OdSelect', 'JsonEncodeBody'

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
        $invokeParameters['Method'] = $(if ($Body) { [Microsoft.PowerShell.Commands.WebRequestMethod]::Post } else { [Microsoft.PowerShell.Commands.WebRequestMethod]::Get })
    }
    if ($UriSuffix) { $uriBuilder.Path += $UriSuffix }
    if ($UriQueryRaw) { appendToQuery $UriQueryRaw }
    $UriQueryParam | Foreach-Object { appendToQuery $_ -split }
    $PSBoundParameters.Keys -ilike 'Od*' | Foreach-Object { appendToQuery "`$$($_.Substring(2).ToLower())" $PSBoundParameters[$_] -skipEmpty }
    if ($Body -and $JsonEncodeBody) {
        # need to explicetly set charset in ContenType for Invoke-RestMethod to detect it and to correctly encode JSON string
        $invokeParameters['ContentType'] = "application/json; charset=UTF-8"
        $invokeParameters['Body'] = $Body | ConvertTo-Json
    }
    $invokeParameters['Uri'] = $uriBuilder.Uri
    $invokeParameters['UseBasicParsing'] = $true

    Write-RjRbDebug "Invoke-RestMethod arguments" $invokeParameters
    try {
        $result = Invoke-RestMethod @invokeParameters
    }
    catch {
        # get error response if available
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
    Write-RjRbLog "Invoke-RestMethod arguments" $invokeParameters -NoDebugOnly
    Write-RjRbLog "Invoke-RestMethod error response" $errorResponse
        throw
    }

    Write-RjRbDebug "Invoke-RestMethod result" $result

    return $result
}
