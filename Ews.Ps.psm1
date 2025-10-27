## Import the EWS Managed API Module
$EwsDLL = "$($PSScriptRoot)\dll\Microsoft.Exchange.WebServices.dll"
Import-Module -Name $EwsDLL -ErrorAction Stop -Force

# Import function code
Get-ChildItem "$($PSScriptRoot)\source\*.ps1" -Recurse -File | ForEach-Object {
    . "$($_.FullName)"
}

if ($PSVersionTable.PSVersion -gt 7.2) {
    $PSStyle.Progress.View = 'Classic'
}