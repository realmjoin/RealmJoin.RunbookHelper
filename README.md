# RealmJoin.RunbookHelper
Helps to integrate Azure Automation scripts with RealmJoin.

# Powershell Gallery
https://www.powershellgallery.com/packages/RealmJoin.RunbookHelper/

# Usage in Azure Automation
Consider a runbook "Group: Add guest to group". By using `Use-RJInterface` RealmJoin will show an enhanced UI when running this script.

```powershell
using module @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.3.0" }

param(
    [Parameter(Mandatory = $true)]
    [string] $GroupID,

    [Parameter(Mandatory = $true)]
    [ValidateScript( { Use-RJInterface Graph -Entity user -Filter "userType eq 'Guest'" } )]
    [string]$GuestID,

    [Parameter(Mandatory = $true)]
    [ValidateScript( { Use-RJInterface -Date -MaxValue "2022-01-01" } )]
    [DateTime]$ValidUntil
)

Write-Output "Adding $GuestID to group $GroupID."
Write-Output "Setting timer to remove guest on $ValidUntil."
```

RealmJoin attaches this script to any group. When called in the context of a group it will automatically populate `$GroupID`.

By using `Use-RJInterface Graph` a rich selector for `$GuestID` will be shown, populated with data from the Graph, filtered by `userType` in this example.

Finally, since `$ValidUntil` is of type `[DateTime]` a date and a time picker will be shown by default, however `Use-RJInterface` narrows this down to a single date picker with additional `MaxValue` constraint.
