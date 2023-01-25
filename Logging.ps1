function Write-RjRbLog {
    [CmdletBinding()]
    param (
        [string] $Message,
        [object] $Data,
        [switch] $NoDebugOnly
    )

    if ($NoDebugOnly -and ($DebugPreference -ne [Management.Automation.ActionPreference]::SilentlyContinue)) {
        return
    }

    Write-RjRbLogImpl @PSBoundParameters
}

function Write-RjRbDebug {
    [CmdletBinding()]
    param (
        [string] $Message,
        [object] $Data
    )

    if ($DebugPreference -eq [Management.Automation.ActionPreference]::SilentlyContinue) {
        return
    }

    Write-RjRbLogImpl @PSBoundParameters
}

function Write-RjRbLogImpl {
    [CmdletBinding()]
    param (
        [string] $Message,
        [object] $Data,
        [switch] $NoDebugOnly # dummy only
    )

    if ($null -ne $Data) {
        if ($Message) {
            $Message += ": "
        }
        if ($Data -is [Hashtable] -and $Data.Count -eq 1) {
            $Message += "$($Data.Keys[0]): $($Data.Values[0] | ConvertTo-Json -Depth 20)"
        }
        else {
            $Message += ($Data | ConvertTo-Json -Depth 20)
        }
    }

    $callingFunction = [string]$(Get-PSCallStack)[2].FunctionName
    if ($callingFunction -ine "<ScriptBlock>") {
        $Message = "$callingFunction`: $Message"
    }

    # use -Verbose to even write message to verbose stream in case preference is off
    Write-Verbose -Verbose -Message $Message
}
