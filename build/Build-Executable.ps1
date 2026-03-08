[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) "dist"),

    [Parameter(Mandatory = $false)]
    [switch]$NoInstallPs2Exe
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$entryScript = Join-Path $repoRoot "Collect-DicomMedia.ps1"
$guiScript = Join-Path $repoRoot "Launch-DcmCollectGui.ps1"
$iconFile = Join-Path $repoRoot "assets\dcmcollect.ico"
$binDir = Join-Path $repoRoot "bin"
$weasisZip = Join-Path $repoRoot "weasis-portable.zip"
$licenseFile = Join-Path $repoRoot "LICENSE"
$thirdPartyNotices = Join-Path $repoRoot "THIRD_PARTY_NOTICES.md"
$readmeFile = Join-Path $repoRoot "README.md"

if (-not (Test-Path -LiteralPath $entryScript -PathType Leaf)) {
    throw "Entry script not found: $entryScript"
}

if (-not (Test-Path -LiteralPath $binDir -PathType Container)) {
    throw "Required folder not found: $binDir"
}

if (-not (Test-Path -LiteralPath $weasisZip -PathType Leaf)) {
    throw "Required file not found: $weasisZip"
}

$invokePs2Exe = Get-Command -Name Invoke-ps2exe -ErrorAction SilentlyContinue
if (-not $invokePs2Exe) {
    if ($NoInstallPs2Exe) {
        throw "Invoke-ps2exe command not found and -NoInstallPs2Exe was specified. Install module 'ps2exe' and retry."
    }

    Write-Host "Preparing PowerShell Gallery prerequisites..."
    $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nugetProvider) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }

    $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psGallery -and $psGallery.InstallationPolicy -ne "Trusted") {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    Write-Host "Installing ps2exe module for current user..."
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
    $invokePs2Exe = Get-Command -Name Invoke-ps2exe -ErrorAction SilentlyContinue
}

if (-not $invokePs2Exe) {
    throw "Invoke-ps2exe command is unavailable after installation attempt."
}

$distRoot = Join-Path $OutputRoot "dcmcollect"
if (Test-Path -LiteralPath $distRoot) {
    Remove-Item -LiteralPath $distRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $distRoot -Force | Out-Null

$exePath = Join-Path $distRoot "dcmcollect.exe"
$guiExePath = Join-Path $distRoot "dcmcollect-gui.exe"

Write-Host "Compiling executable..."
$ps2ExeParams = @{
    inputFile    = $entryScript
    outputFile   = $exePath
    x64          = $true
    STA          = $true
    requireAdmin = $false
    noConsole    = $false
}
if (Test-Path -LiteralPath $iconFile -PathType Leaf) {
    $ps2ExeParams["iconFile"] = $iconFile
}
Invoke-ps2exe @ps2ExeParams

if (Test-Path -LiteralPath $guiScript -PathType Leaf) {
    Write-Host "Compiling GUI executable..."
    $guiPs2ExeParams = @{
        inputFile    = $guiScript
        outputFile   = $guiExePath
        x64          = $true
        STA          = $true
        requireAdmin = $false
        noConsole    = $true
    }
    if (Test-Path -LiteralPath $iconFile -PathType Leaf) {
        $guiPs2ExeParams["iconFile"] = $iconFile
    }
    Invoke-ps2exe @guiPs2ExeParams
}

Write-Host "Copying runtime dependencies..."
Copy-Item -LiteralPath $binDir -Destination (Join-Path $distRoot "bin") -Recurse -Force
Copy-Item -LiteralPath $weasisZip -Destination (Join-Path $distRoot "weasis-portable.zip") -Force
if (Test-Path -LiteralPath $iconFile -PathType Leaf) {
    $assetsOut = Join-Path $distRoot "assets"
    New-Item -ItemType Directory -Path $assetsOut -Force | Out-Null
    Copy-Item -LiteralPath $iconFile -Destination (Join-Path $assetsOut "dcmcollect.ico") -Force
}

if (Test-Path -LiteralPath $licenseFile -PathType Leaf) {
    Copy-Item -LiteralPath $licenseFile -Destination (Join-Path $distRoot "LICENSE") -Force
}
if (Test-Path -LiteralPath $thirdPartyNotices -PathType Leaf) {
    Copy-Item -LiteralPath $thirdPartyNotices -Destination (Join-Path $distRoot "THIRD_PARTY_NOTICES.md") -Force
}
if (Test-Path -LiteralPath $readmeFile -PathType Leaf) {
    Copy-Item -LiteralPath $readmeFile -Destination (Join-Path $distRoot "README.md") -Force
}

Write-Host "Done."
Write-Host "Output folder: $distRoot"
Write-Host "Executable: $exePath"
if (Test-Path -LiteralPath $guiExePath -PathType Leaf) {
    Write-Host "GUI executable: $guiExePath"
}
Write-Host "Example run:"
Write-Host '  .\dcmcollect.exe -Src "C:\input" -Dest "C:\output_media" -PackageWeasis -VerifyDicomdir'
if (Test-Path -LiteralPath $guiExePath -PathType Leaf) {
    Write-Host '  .\dcmcollect-gui.exe'
}
