$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$app = Get-ChildItem -Path $root -Filter "ReimbursementMatcher.dll" -Recurse |
    Where-Object { $_.FullName -notmatch "\\src\\" -and $_.FullName -notmatch "\\obj\\" -and $_.FullName -notmatch "\\bin\\" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $app) {
    throw "ReimbursementMatcher.dll was not found. Please publish the app first."
}

if ($args.Count -gt 0) {
    dotnet $app.FullName --invoice-recognition-audit $args[0]
} else {
    dotnet $app.FullName --invoice-recognition-audit
}

Write-Host ""
Write-Host "Invoice recognition audit finished. Check the output folder."
