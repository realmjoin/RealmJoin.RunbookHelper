# RealmJoin.RunbookHelper
Helps to integrate Azure Automation scripts with RealmJoin.

# Powershell Gallery
https://www.powershellgallery.com/packages/RealmJoin.RunbookHelper/

# Usage in Azure Automation
Consider a runbook "Group: Add guest to group". By using `Use-RjRbInterface` RealmJoin will show an enhanced UI when running this script.

```powershell
#Requires -Module @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.3.0" }

param(
    [Parameter(Mandatory = $true)]
    [string] $GroupID,

    [Parameter(Mandatory = $true)]
    [ValidateScript( { Use-RjRbInterface Graph -Entity user -Filter "userType eq 'Guest'" } )]
    [string]$GuestID,

    [Parameter(Mandatory = $true)]
    [ValidateScript( { Use-RjRbInterface -Date -MaxValue "2022-01-01" } )]
    [DateTime]$ValidUntil
)

Write-Output "Adding $GuestID to group $GroupID."
Write-Output "Setting timer to remove guest on $ValidUntil."
```

RealmJoin attaches this script to any group. When called in the context of a group it will automatically populate `$GroupID`.

By using `Use-RjRbInterface Graph` a rich selector for `$GuestID` will be shown, populated with data from the Graph, filtered by `userType` in this example.

Finally, since `$ValidUntil` is of type `[DateTime]` a date and a time picker will be shown by default, however `Use-RjRbInterface` narrows this down to a single date picker with additional `MaxValue` constraint.

# Available options
```
NAME
    Use-RjRbInterface
    
SYNTAX
    Use-RjRbInterface [[-Type] {Graph | Number | DateTime | Textarea}] [[-Entity] {User | Group | Device}] [[-Attribute] <string>] [[-Filter] <string>] [[-MinValue] <Object>] [[-MaxValue] <Object>] [-Date] [-Time] 
    
    
PARAMETERS
    -Attribute <string>
        
        Required?                    false
        Position?                    2
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Date
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Entity <string>
        
        Required?                    false
        Position?                    1
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Filter <string>
        
        Required?                    false
        Position?                    3
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -MaxValue <Object>
        
        Required?                    false
        Position?                    5
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -MinValue <Object>
        
        Required?                    false
        Position?                    4
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Time
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Type <string>
        
        Required?                    false
        Position?                    0
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    
INPUTS
    None
    
    
OUTPUTS
    System.Object
    
ALIASES
    None
    

REMARKS
    None
```