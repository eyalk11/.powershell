if ($PSVersionTable.PSVersion -like "7*")
{
    Remove-Alias cd
}
Import-Module $PSScriptRoot\common.psm1 -DisableNameChecking
if (Test-path $PSScriptRoot\my.ps1) 
{
    . $PSScriptRoot\my.ps1
}
if (Test-path $PSScriptRoot\secret.ps1) 
{
    . $PSScriptRoot\secret.ps1
}
