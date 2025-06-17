Usage:

. path\to\Utils.ps1

$data = Get-Sheet-From-PWD (Get-Location)      #get from pwd

$data = Get-Sheet-From-PWD $PSScriptRoot       #get from script directory

$data = Get-Sheet-From-PWD "C:\Path\to\Excel\" #get from path
