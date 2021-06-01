function rjRbGetParametersFiltered {
    param (
        [string[]] $include,
        [string[]] $exclude,
        [int] $callStackLevel = 1
    )

    $callStackItem = $(Get-PSCallStack)[$callStackLevel]

    $parameterNames = $callStackItem.InvocationInfo.MyCommand.Parameters.Keys
    if ($null -ne $include -and $include.Length) {
        $parameterNames = $parameterNames | Where-Object { $include -icontains $_ }
    }
    if ($null -ne $exclude -and $exclude.Length) {
        $parameterNames = $parameterNames | Where-Object { $exclude -inotcontains $_ }
    }

    $paramsAndValues = @{}
    Get-Variable -Scope $callStackLevel | Where-Object { $parameterNames -icontains $_.Name } | ForEach-Object { $paramsAndValues[$_.Name] = $_.Value }

    return $paramsAndValues
}
