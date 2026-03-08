[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DistRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) "dist"),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$PackageName = "dcmcollect-windows-x64",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Version = (Get-Date -Format "yyyy.MM.dd"),

    [Parameter(Mandatory = $false)]
    [switch]$Rebuild,

    [Parameter(Mandatory = $false)]
    [switch]$NoInstallPs2Exe
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$buildExeScript = Join-Path $PSScriptRoot "Build-Executable.ps1"
$stageDir = Join-Path $DistRoot "dcmcollect"

if ($Rebuild) {
    if (-not (Test-Path -LiteralPath $buildExeScript -PathType Leaf)) {
        throw "Build script not found: $buildExeScript"
    }

    Write-Host "Rebuilding executable package before release zip..."
    $buildParams = @{
        OutputRoot = $DistRoot
    }
    if ($NoInstallPs2Exe) {
        $buildParams["NoInstallPs2Exe"] = $true
    }
    & $buildExeScript @buildParams
}

if (-not (Test-Path -LiteralPath $stageDir -PathType Container)) {
    throw "Staging directory not found: $stageDir. Run .\\build\\Build-Executable.ps1 first or use -Rebuild."
}

$zipFileName = "{0}-{1}.zip" -f $PackageName, $Version
$zipPath = Join-Path $DistRoot $zipFileName

if (Test-Path -LiteralPath $zipPath -PathType Leaf) {
    Remove-Item -LiteralPath $zipPath -Force
}

Write-Host "Creating release zip: $zipPath"
Compress-Archive -Path (Join-Path $stageDir "*") -DestinationPath $zipPath -CompressionLevel Optimal -Force

Write-Host "Done."
Write-Host "Release archive: $zipPath"
