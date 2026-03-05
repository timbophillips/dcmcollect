<#
.SYNOPSIS
    Collect DICOM files from a folder tree using DCMTK tools,
  copy them into a destination "media" folder, and generate a DICOMDIR using dcmmkdir.

.NOTES
  Expects DCMTK binaries in a ".\bin\" subfolder beneath this script:
    .\bin\dcmftest.exe
    .\bin\dcmdump.exe
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
$dcmdump   = Join-Path $binDir "dcmdump.exe"
$dcmmkdir  = Join-Path $binDir "dcmmkdir.exe"

if (-not (Test-Path -LiteralPath $dcmftest -PathType Leaf)) {
    throw "Missing DCMTK binary: $dcmftest"
}
if (-not (Test-Path -LiteralPath $dcmdump -PathType Leaf)) {
    throw "Missing DCMTK binary: $dcmdump"
}
if (-not (Test-Path -LiteralPath $dcmmkdir -PathType Leaf)) {
    throw "Missing DCMTK binary: $dcmmkdir"
}

function Test-IsDicom {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$DcmfTestExe,
        [Parameter(Mandatory = $true)][string]$DcmDumpExe
    )

    # Fast path: strict DICOM Part-10 check.
    & $DcmfTestExe $FilePath *> $null
    if ($LASTEXITCODE -eq 0) {
        return $true
    }

    # Fallback: parse as file format OR dataset and look for a core DICOM tag.
    & $DcmDumpExe -q +P "SOPClassUID" $FilePath *> $null
    return ($LASTEXITCODE -eq 0)
}

function Get-DicomdirReferencedFiles {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$DcmDumpExe
    )

    $results = New-Object System.Collections.Generic.List[string]

    $dicomdirs = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -Force |
        Where-Object { $_.Name -ieq "DICOMDIR" }

    foreach ($dd in $dicomdirs) {
        $baseDir = Split-Path -Parent $dd.FullName

        $lines = & $DcmDumpExe -q +P "ReferencedFileID" $dd.FullName 2>$null
        if ($LASTEXITCODE -ne 0) {
            continue
        }

        foreach ($line in $lines) {
            if ($line -match '\(0004,1500\).*\[(.*?)\]') {
                $ref = $matches[1].Trim()
                if ([string]::IsNullOrWhiteSpace($ref)) {
                    continue
                }

                $refPath = ($ref -replace '/', '\')
                $candidate = Join-Path $baseDir $refPath

                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    [void]$results.Add((Resolve-Path -LiteralPath $candidate).Path)
                }
            }
        }
    }

    return $results
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


$candidateFiles = New-Object System.Collections.Generic.List[string]
Get-ChildItem -LiteralPath $SrcFull -Recurse -File -Force | ForEach-Object {
    [void]$candidateFiles.Add($_.FullName)
}

$dicomdirRefs = Get-DicomdirReferencedFiles -SourceRoot $SrcFull -DcmDumpExe $dcmdump
foreach ($refFile in $dicomdirRefs) {
    [void]$candidateFiles.Add($refFile)
}

$seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$i = 1

foreach ($f in $candidateFiles) {
    if (-not $seen.Add($f)) {
        continue
    }

    if (Test-IsDicom -FilePath $f -DcmfTestExe $dcmftest -DcmDumpExe $dcmdump) {
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
Write-Host "DICOMDIR references discovered: $($dicomdirRefs.Count)"
Write-Host "Media folder: $mediaDir"
Write-Host "DICOMDIR: $dicomdirPath"
Write-Host "Catalogue: $catalog"