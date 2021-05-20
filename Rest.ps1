function Invoke-RjRbRestMethodGraph {
    [CmdletBinding()]
    param (
        [string] $Resource,
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [switch] $Beta
    )

    $uri = "https://graph.microsoft.com/$(if($Beta) {"beta"} else {"v1.0"})"
    $PSBoundParameters.Remove('Beta') | Out-Null

    Invoke-RjRbRestMethod -Uri $uri -JsonEncodeBody @PSBoundParameters | Select-Object -ExpandProperty value
}

function Invoke-RjRbRestMethod {
    [CmdletBinding()]
    param (
        [uri] $Uri,
        [Microsoft.PowerShell.Commands.WebRequestMethod] $Method,
        [Collections.IDictionary] $Headers,
        [object] $Body,
        [Alias("Resource")][string] $UriSuffix,
        [switch] $JsonEncodeBody
    )

    $invokeParameters = $PSBoundParameters
    if ($Method -eq [Microsoft.PowerShell.Commands.WebRequestMethod]::Default) {
        $invokeParameters['Method'] = $(if ($Body) { [Microsoft.PowerShell.Commands.WebRequestMethod]::Post } else { [Microsoft.PowerShell.Commands.WebRequestMethod]::Get })
    }
    if ($UriSuffix) {
        $invokeParameters['Uri'] = [uri]($Uri.ToString() + $UriSuffix)
        $invokeParameters.Remove('UriSuffix') | Out-Null
    }
    if (-not $Headers -and $Global:RjRbGraphAuthHeaders) {
        $invokeParameters['Headers'] = $Global:RjRbGraphAuthHeaders
    }
    if ($Body -and $JsonEncodeBody) {
        # need to explicetly set charset in ContenType for Invoke-RestMethod to detect it and to correctly encode JSON string
        $invokeParameters['ContentType'] = "application/json; charset=UTF-8"
        $invokeParameters['Body'] = $Body | ConvertTo-Json
    }
    $invokeParameters.Remove('JsonEncodeBody') | Out-Null
    $invokeParameters['UseBasicParsing'] = $true

    Write-RjRbDebug "Invoke-RestMethod" $invokeParameters
    try {
        $result = Invoke-RestMethod @invokeParameters
    }
    catch {
        Write-RjRbLog -NoDebugOnly "Invoke-RestMethod" $invokeParameters
        throw
    }

    Write-RjRbDebug "Result" $result

    return $result
}
