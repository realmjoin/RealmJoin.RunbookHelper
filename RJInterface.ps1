function Use-RJInterface {
    param (
        [ValidateSet("Graph", "Number", "DateTime", "Textarea")]
        [string]$Type,
        [ValidateSet("User", "Group", "Device")]
        [string]$Entity,
        [string]$Attribute,
        [string]$Filter,
        [switch]$Date,
        [switch]$Time,
        [object]$MinValue,
        [object]$MaxValue
    )

    if ($DebugPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue) {
        $debug = @"
Type      = $Type
Entity    = $Entity
Attribute = $Attribute
Filter    = $Filter
Date      = $Date
Time      = $Time
MinValue  = $MinValue
MaxValue  = $MaxValue
"@

        Write-Debug "Use-RJInterface: `r`n$debug"
    }

    return $true
}
