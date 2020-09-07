[CmdletBinding()]
param (
)
$Moduleinfo = Test-ModuleManifest -Path ((Get-ChildItem $PSScriptRoot\*.psd1).FullName)
$ModulePath = (Get-Module -Name ($Moduleinfo.Name.ToString())).ModuleBase.ToString()
$ModulePath = $ModulePath -replace ($ModulePath.Split('\\')[-1],$null)

if (Test-Path $ModulePath) {
    Remove-Module ($Moduleinfo.Name) -ErrorAction SilentlyContinue -Verbose
    Remove-Item $ModulePath -Recurse -Force -Verbose
}
