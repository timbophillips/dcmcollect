<#
.SYNOPSIS
  Collect DICOM (Part-10) files from a folder tree using DCMTK dcmftest,
  copy them into a destination "media" folder, and generate a DICOMDIR using dcmmkdir.

.NOTES
  Expects DCMTK binaries in a ".\bin\" subfolder beneath this script:
    .\bin\dcmftest.exe
    .\bin\dcmmkdir.exe

.USAGE
  .\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media"
  .\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -Subdir "IMAGES"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Src,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$Dest,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Subdir = "IMAGES"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([Parameter(Mandatory=$true)][string]$Path)
    try { return (Resolve-Path -LiteralPath $Path).Path }
    catch { return $Path }  # Dest may not exist yet
}

# --- Locate DCMTK binaries relative to this script ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$binDir    = Join-Path $scriptDir "bin"

$dcmftest  = Join-Path $binDir "dcmftest.exe"
$dcmmkdir  = Join-Path $binDir "dcmmkdir.exe"

if (-not (Test-Path -LiteralPath $dcmftest -PathType Leaf)) {
    throw "Missing DCMTK binary: $dcmftest"
}
if (-not (Test-Path -LiteralPath $dcmmkdir -PathType Leaf)) {
    throw "Missing DCMTK binary: $dcmmkdir"
}

# --- Validate paths ---
$SrcFull  = Resolve-FullPath $Src
$DestFull = Resolve-FullPath $Dest

if (-not (Test-Path -LiteralPath $SrcFull -PathType Container)) {
    throw "Source folder does not exist or is not a directory: $SrcFull"
}

# Ensure destination and media subdir exist
$mediaDir = Join-Path $DestFull $Subdir
New-Item -ItemType Directory -Force -Path $mediaDir | Out-Null

# Catalogue file
$catalog = Join-Path $DestFull "catalogue.csv"
"seq,new_filename,source_path" | Set-Content -Encoding ASCII $catalog

$i = 1

Get-ChildItem -LiteralPath $SrcFull -Recurse -File | ForEach-Object {
    $f = $_.FullName

    # dcmftest returns exit code 0 if file is DICOM Part-10
    & $dcmftest $f *> $null
    if ($LASTEXITCODE -eq 0) {
        $new = "{0:D8}" -f $i
        $outRel = Join-Path $Subdir $new
        $outAbs = Join-Path $DestFull $outRel

        Copy-Item -LiteralPath $f -Destination $outAbs -Force

        # CSV-escape quotes
        $escaped = $f.Replace('"','""')
        "$i,$outRel,""$escaped""" | Add-Content -Encoding ASCII $catalog

        $i++
    }
}

# Build DICOMDIR (scan media subdir under DEST, write DEST\DICOMDIR)
$dicomdirPath = Join-Path $DestFull "DICOMDIR"

# Run dcmmkdir with DEST as working directory so relative paths resolve correctly.
Push-Location -LiteralPath $DestFull
try {
    & $dcmmkdir +r +id $Subdir +D $dicomdirPath -v
}
finally {
    Pop-Location
}

Write-Host "Done."
Write-Host "Copied DICOM files: $($i - 1)"
Write-Host "Media folder: $mediaDir"
Write-Host "DICOMDIR: $dicomdirPath"
Write-Host "Catalogue: $catalog"