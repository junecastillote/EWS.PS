[CmdletBinding()]
param (
    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ModulePath
)

Get-ChildItem $PSScriptRoot -Recurse | Unblock-File

$Moduleinfo = Test-ModuleManifest -Path ((Get-ChildItem $PSScriptRoot\*.psd1).FullName)
$ModulePath = $ModulePath + "\$($Moduleinfo.Name.ToString())\$($Moduleinfo.Version.ToString())"

if (!(Test-Path $ModulePath)) {
    New-Item -Path $ModulePath -ItemType Directory -Force | Out-Null
}

Copy-Item -Path $PSScriptRoot\*.psd1,$PSScriptRoot\*.psm1 -Destination $ModulePath -Force -Confirm:$false -Verbose
#Copy-Item -Path $PSScriptRoot\*.psm1 -Destination $ModulePath -Force -Confirm:$false
Copy-Item -Path $PSScriptRoot\src\* -Destination (New-Item -ItemType Directory $ModulePath\src -Force).FullName -Force -Confirm:$false -Verbose

Remove-Module ($Moduleinfo.Name) -ErrorAction SilentlyContinue
Import-Module ($Moduleinfo.Name)