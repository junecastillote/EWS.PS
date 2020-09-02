Get-ChildItem "$($PSScriptRoot)\src\*.ps1" | ForEach-Object {
    . "$($_.FullName)"
}