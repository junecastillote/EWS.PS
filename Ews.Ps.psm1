## Check registry if EWS Managed API is installed
# $EwsDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
# if (!($EwsDLL) -or !(Test-Path $EwsDLL)) {
#     Write-Error "The EWS Managed API is not found. Go to https://www.microsoft.com/en-us/download/details.aspx?id=42951 to download and install."
#     return $null
# }

## Import the EWS Managed API Module
$EwsDLL = "$($PSScriptRoot)\dll\Microsoft.Exchange.WebServices.dll"
Import-Module -Name $EwsDLL -ErrorAction Stop -Force


Get-ChildItem "$($PSScriptRoot)\src\*.ps1" | ForEach-Object {
    . "$($_.FullName)"
}