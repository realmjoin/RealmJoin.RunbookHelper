function Use-RjRbInterface {
    param (
        [ValidateSet("Graph", "Number", "DateTime", "Textarea", "Setting")]
        [string]$Type,
        [ValidateSet("User", "Group", "Device")]
        [string]$Entity,
        [string]$Attribute,
        [string]$Filter,
        [switch]$Date,
        [switch]$Time,
        [object]$MinValue,
        [object]$MaxValue,
        [string]$DisplayName,
        [string]$Description
    )

    if ($DebugPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue) {
        $debug = @"
Type        = $Type
Entity      = $Entity
Attribute   = $Attribute
Filter      = $Filter
Date        = $Date
Time        = $Time
MinValue    = $MinValue
MaxValue    = $MaxValue
DisplayName = $DisplayName
Description = $Description
"@

        Write-Debug "Use-RjRbInterface: `r`n$debug"
    }

    return $true
}

New-Alias -Name Use-RJInterface -Value Use-RjRbInterface
