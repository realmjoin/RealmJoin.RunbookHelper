#requires -Version 5.0

class RealmJoinUIAttribute : Attribute
{
    [Entity]$Graph = [Entity]::None;
    [Picker]$Picker = [Picker]::None;
    [string]$Attribute = $null;
    [string]$Filter = $null;
}

enum Entity
{
    None = 0
    User = 10
    Group = 20
    Device = 30
}

enum Picker
{
    None = 0
    Date = 10
    Time = 20
    DateTime = 30
}
