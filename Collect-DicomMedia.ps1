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
    .\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -Deidentify
    .\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -PackageWeasis
    .\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -PackageWeasis -WeasisSource "C:\tools\weasis-portable.zip"
    .\Collect-DicomMedia.ps1 -Dest "C:\output_media" -WeasisOnly
        .\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -VerifyDicomdir
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$Src,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$Dest,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Subdir = "IMAGES",

    [Parameter(Mandatory = $false)]
    [switch]$PackageWeasis,

    [Parameter(Mandatory = $false)]
    [string]$WeasisSource
,

    [Parameter(Mandatory = $false)]
    [switch]$WeasisOnly,

    [Parameter(Mandatory = $false)]
    [switch]$VerifyDicomdir
,
    [Parameter(Mandatory = $false)]
    [switch]$Deidentify
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# In PowerShell 7+, native tools that write informational lines to stderr can be
# promoted to terminating errors under ErrorActionPreference=Stop.
if (Test-Path variable:PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

function Resolve-FullPath {
    param([Parameter(Mandatory=$true)][string]$Path)
    try { return (Resolve-Path -LiteralPath $Path).Path }
    catch { return $Path }  # Dest may not exist yet
}

function Get-ScriptDirectory {
    if ($PSScriptRoot -and (Test-Path -LiteralPath $PSScriptRoot -PathType Container)) {
        return $PSScriptRoot
    }

    $myCommandPath = $null
    if ($MyInvocation -and $MyInvocation.MyCommand) {
        if ($MyInvocation.MyCommand -is [string]) {
            $myCommandPath = $MyInvocation.MyCommand
        }
        elseif ($MyInvocation.MyCommand.PSObject.Properties.Match("Path").Count -gt 0) {
            $myCommandPath = $MyInvocation.MyCommand.Path
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($myCommandPath)) {
        return (Split-Path -Parent $myCommandPath)
    }

    # Fallback for compiled executable context.
    return [System.AppContext]::BaseDirectory.TrimEnd([char]'\')
}

# --- Locate binaries relative to this script ---
$scriptDir = Get-ScriptDirectory
$binDir    = Join-Path $scriptDir "bin"

$dcmftest  = Join-Path $binDir "dcmftest.exe"
$dcmdump   = Join-Path $binDir "dcmdump.exe"
$dcmmkdir  = Join-Path $binDir "dcmmkdir.exe"
$dcmodify  = Join-Path $binDir "dcmodify.exe"

function Invoke-DeidentifyDicom {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$DcmodifyExe
    )

    # Basic de-identification profile: overwrite a core set of direct identifiers.
    $args = @(
        '-q',
        '-nb',
        '-ma', '(0010,0010)=ANON^PATIENT',
        '-ma', '(0010,0020)=ANONID',
        '-ma', '(0010,0030)=',
        '-ma', '(0010,0040)=O',
        '-ma', '(0010,1000)=',
        '-ma', '(0010,1040)=',
        '-ma', '(0008,0050)=',
        '-ma', '(0008,0090)=',
        '-ma', '(0008,0080)=',
        '-ma', '(0008,0081)=',
        '-ma', '(0008,1010)=',
        '-ma', '(0008,1030)=ANON_STUDY',
        '-ma', '(0008,103E)=ANON_SERIES',
        $FilePath
    )

    $outFile = [System.IO.Path]::GetTempFileName()
    $errFile = [System.IO.Path]::GetTempFileName()
    try {
        $startParams = @{
            FilePath               = $DcmodifyExe
            ArgumentList           = $args
            NoNewWindow            = $true
            Wait                   = $true
            PassThru               = $true
            RedirectStandardOutput = $outFile
            RedirectStandardError  = $errFile
        }
        $proc = Start-Process @startParams

        if ($proc.ExitCode -ne 0) {
            $errText = (Get-Content -LiteralPath $errFile -ErrorAction SilentlyContinue | Out-String).Trim()
            if ([string]::IsNullOrWhiteSpace($errText)) {
                throw "dcmodify failed with exit code $($proc.ExitCode) for file: $FilePath"
            }
            throw "dcmodify failed with exit code $($proc.ExitCode) for file: $FilePath`n$errText"
        }
    }
    finally {
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    }
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

function Get-DicomdirReferencedFileIds {
    param(
        [Parameter(Mandatory = $true)][string]$DicomdirPath,
        [Parameter(Mandatory = $true)][string]$DcmDumpExe
    )

    $results = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $DicomdirPath -PathType Leaf)) {
        return $results
    }

    $lines = & $DcmDumpExe -q +P "ReferencedFileID" $DicomdirPath 2>$null
    if ($LASTEXITCODE -ne 0) {
        return $results
    }

    foreach ($line in $lines) {
        if ($line -match '\(0004,1500\).*\[(.*?)\]') {
            $ref = $matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($ref)) {
                [void]$results.Add(($ref -replace '/', '\'))
            }
        }
    }

    return $results
}

function Invoke-DicomdirVerification {
    param(
        [Parameter(Mandatory = $true)][string]$DicomdirPath,
        [Parameter(Mandatory = $true)][string]$MediaSubdir,
        [Parameter(Mandatory = $true)][string]$DcmDumpExe
    )

    Write-Host "Verification: checking ReferencedFileID entries in DICOMDIR"

    $refs = Get-DicomdirReferencedFileIds -DicomdirPath $DicomdirPath -DcmDumpExe $DcmDumpExe
    if ($refs.Count -eq 0) {
        Write-Warning "Verification: no ReferencedFileID entries were found."
        return
    }

    $normalizedSubdir = ($MediaSubdir -replace '/', '\').Trim('\')
    $expectedPrefix = "$normalizedSubdir\"
    $mismatches = New-Object System.Collections.Generic.List[string]

    foreach ($ref in $refs) {
        if (-not ($ref.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase))) {
            [void]$mismatches.Add($ref)
        }
    }

    $sampleCount = [Math]::Min(10, $refs.Count)
    Write-Host "Verification: sampled ReferencedFileID entries ($sampleCount of $($refs.Count)):"
    for ($idx = 0; $idx -lt $sampleCount; $idx++) {
        Write-Host ("  [{0}] {1}" -f ($idx + 1), $refs[$idx])
    }

    if ($mismatches.Count -gt 0) {
        Write-Warning "Verification: found $($mismatches.Count) entries that do not start with '$expectedPrefix'."
        $mismatchSample = [Math]::Min(5, $mismatches.Count)
        Write-Host "Verification: mismatch examples:"
        for ($idx = 0; $idx -lt $mismatchSample; $idx++) {
            Write-Host ("  [X{0}] {1}" -f ($idx + 1), $mismatches[$idx])
        }
    }
    else {
        Write-Host "Verification: OK - all ReferencedFileID entries start with '$expectedPrefix'."
    }
}

function Install-WeasisPackage {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationRoot,
        [Parameter(Mandatory = $true)][string]$MediaSubdir,
        [Parameter(Mandatory = $true)][string]$DicomdirPath
    )

    $srcResolved = Resolve-FullPath $SourcePath
    if (-not (Test-Path -LiteralPath $srcResolved)) {
        throw "Weasis source does not exist: $srcResolved"
    }

    $weasisDir = Join-Path $DestinationRoot "weasis"
    if (Test-Path -LiteralPath $weasisDir) {
        Remove-Item -LiteralPath $weasisDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $weasisDir | Out-Null

    $srcItem = Get-Item -LiteralPath $srcResolved
    if ($srcItem.PSIsContainer) {
        Copy-Item -LiteralPath (Join-Path $srcResolved "*") -Destination $weasisDir -Recurse -Force
    }
    else {
        $ext = [System.IO.Path]::GetExtension($srcResolved)
        if ($ext -ieq ".zip") {
            Expand-Archive -LiteralPath $srcResolved -DestinationPath $weasisDir -Force
        }
        else {
            throw "Unsupported Weasis source format: $srcResolved (expected a folder or .zip)"
        }
    }

    $exeCandidates = Get-ChildItem -LiteralPath $weasisDir -Recurse -File -Filter "*.exe" |
        Where-Object {
            $_.Name -match '^(weasis|viewer-win).*\.exe$' -and
            $_.Name -notmatch 'updater'
        } |
        Sort-Object { $_.FullName.Length }

    if (-not $exeCandidates) {
        # Fallback for non-standard package layouts: pick first EXE in package tree.
        $exeCandidates = Get-ChildItem -LiteralPath $weasisDir -Recurse -File -Filter "*.exe" |
            Sort-Object { $_.FullName.Length }
    }

    if (-not $exeCandidates) {
        throw "Unable to find Weasis executable under: $weasisDir"
    }

    $weasisExe = $exeCandidates[0].FullName

    # Keep launcher portable by resolving executable at runtime from destination root.
    $weasisExeRel = $weasisExe.Substring($DestinationRoot.Length).TrimStart([char]'\', [char]'/')

    $launcherPs1 = Join-Path $DestinationRoot "Launch-Weasis.ps1"
    $launcherCmd = Join-Path $DestinationRoot "Launch-Weasis.cmd"

    $ps1Content = @"
`$ErrorActionPreference = "Stop"
`$destRoot = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$weasisExe = Join-Path `$destRoot "$weasisExeRel"
`$dicomdir = Join-Path `$destRoot "DICOMDIR"
`$mediaDir = Join-Path `$destRoot "$MediaSubdir"

if (-not (Test-Path -LiteralPath `$weasisExe -PathType Leaf)) {
    throw "Weasis executable not found: `$weasisExe"
}

# Weasis portable launchers only forward arguments that start with '$' or 'weasis://'.
# Use an encoded weasis:// URI and include IMAGES as fallback input.
if (Test-Path -LiteralPath `$dicomdir -PathType Leaf) {
    `$startupCmd = ('`$dicom:get -l "{0}" -l "{1}"' -f (`$dicomdir -replace '\\', '/'), (`$mediaDir -replace '\\', '/'))
}
else {
    `$startupCmd = ('`$dicom:get -l "{0}"' -f (`$mediaDir -replace '\\', '/'))
}

`$startupArg = ('weasis://?{0}' -f [uri]::EscapeDataString(`$startupCmd))
& `$weasisExe `$startupArg
"@

    Set-Content -LiteralPath $launcherPs1 -Encoding ASCII -Value $ps1Content

    $cmdContent = "@echo off`r`nsetlocal`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"%~dp0Launch-Weasis.ps1`"`r`n"
    Set-Content -LiteralPath $launcherCmd -Encoding ASCII -Value $cmdContent

    return [pscustomobject]@{
        WeasisDir    = $weasisDir
        WeasisExe    = $weasisExe
        LauncherPs1  = $launcherPs1
        LauncherCmd  = $launcherCmd
    }
}

# --- Validate paths ---
$DestFull = Resolve-FullPath $Dest
$SrcFull  = $null

if ($WeasisOnly) {
    $PackageWeasis = $true
    if ($Deidentify) {
        Write-Warning "-Deidentify is ignored in -WeasisOnly mode because no DICOM copy step is run."
    }
}

if ($PackageWeasis -and [string]::IsNullOrWhiteSpace($WeasisSource)) {
    $defaultWeasisZip = Join-Path $scriptDir "weasis-portable.zip"
    if (Test-Path -LiteralPath $defaultWeasisZip -PathType Leaf) {
        $WeasisSource = $defaultWeasisZip
    }
    else {
        throw "-PackageWeasis was specified but no -WeasisSource was provided and default package was not found: $defaultWeasisZip"
    }
}

if (-not $WeasisOnly) {
    if ([string]::IsNullOrWhiteSpace($Src)) {
        throw "-Src is required unless -WeasisOnly is specified."
    }

    if (-not (Test-Path -LiteralPath $dcmftest -PathType Leaf)) {
        throw "Missing DCMTK binary: $dcmftest"
    }
    if (-not (Test-Path -LiteralPath $dcmdump -PathType Leaf)) {
        throw "Missing DCMTK binary: $dcmdump"
    }
    if (-not (Test-Path -LiteralPath $dcmmkdir -PathType Leaf)) {
        throw "Missing DCMTK binary: $dcmmkdir"
    }
    if ($Deidentify -and (-not (Test-Path -LiteralPath $dcmodify -PathType Leaf))) {
        throw "Missing DCMTK binary required for de-identification: $dcmodify"
    }

    $SrcFull = Resolve-FullPath $Src
    if (-not (Test-Path -LiteralPath $SrcFull -PathType Container)) {
        throw "Source folder does not exist or is not a directory: $SrcFull"
    }
}

if ($WeasisOnly -and (-not (Test-Path -LiteralPath $DestFull -PathType Container))) {
    throw "Destination folder must already exist when using -WeasisOnly: $DestFull"
}

# Ensure destination and media subdir exist
$mediaDir = Join-Path $DestFull $Subdir
New-Item -ItemType Directory -Force -Path $mediaDir | Out-Null

$catalog = Join-Path $DestFull "catalogue.csv"
$dicomdirPath = Join-Path $DestFull "DICOMDIR"
$dicomdirRefs = @()
$copiedCount = 0
$deidentifiedCount = 0

if (-not $WeasisOnly) {
    $discoveryWatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Catalogue file
    "seq,new_filename,source_path" | Set-Content -Encoding ASCII $catalog

    $candidateFiles = New-Object System.Collections.Generic.List[string]
    $discoveredCandidates = 0

    Get-ChildItem -LiteralPath $SrcFull -Recurse -File -Force | ForEach-Object {
        [void]$candidateFiles.Add($_.FullName)
        $discoveredCandidates++
        Write-Output ("DISCOVER|COUNT={0}|FILE={1}" -f $discoveredCandidates, $_.FullName)
        Write-Output ("Found candidate [{0}] {1}" -f $discoveredCandidates, $_.FullName)
    }

    $dicomdirRefs = Get-DicomdirReferencedFiles -SourceRoot $SrcFull -DcmDumpExe $dcmdump
    foreach ($refFile in $dicomdirRefs) {
        [void]$candidateFiles.Add($refFile)
        $discoveredCandidates++
        Write-Output ("DISCOVER|COUNT={0}|FILE={1}" -f $discoveredCandidates, $refFile)
        Write-Output ("Found candidate [{0}] {1}" -f $discoveredCandidates, $refFile)
    }

    $totalCandidates = $candidateFiles.Count
    $discoveryWatch.Stop()
    Write-Output ("STAT|TOTAL_CANDIDATES={0}" -f $totalCandidates)
    Write-Output ("PHASE|DISCOVERY_SECONDS={0:N2}" -f $discoveryWatch.Elapsed.TotalSeconds)

    $scanWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $i = 1
    $processedCandidates = 0

    foreach ($f in $candidateFiles) {
        $processedCandidates++

        if (-not $seen.Add($f)) {
            Write-Output ("PROGRESS|PROCESSED={0}|COPIED={1}|CURRENT={2}" -f $processedCandidates, ($i - 1), $f)
            continue
        }

        if (Test-IsDicom -FilePath $f -DcmfTestExe $dcmftest -DcmDumpExe $dcmdump) {
            $new = "{0:D8}" -f $i
            $outRel = Join-Path $Subdir $new
            $outAbs = Join-Path $DestFull $outRel

            Copy-Item -LiteralPath $f -Destination $outAbs -Force
            Write-Output ("Copied [{0}] {1} <= {2}" -f $i, $outRel, $f)

            if ($Deidentify) {
                Invoke-DeidentifyDicom -FilePath $outAbs -DcmodifyExe $dcmodify
                $deidentifiedCount++
                Write-Output ("De-identified [{0}] {1}" -f $deidentifiedCount, $outRel)
            }

            # CSV-escape quotes
            $escaped = $f.Replace('"','""')
            "$i,$outRel,""$escaped""" | Add-Content -Encoding ASCII $catalog

            $i++
        }

        Write-Output ("PROGRESS|PROCESSED={0}|COPIED={1}|CURRENT={2}" -f $processedCandidates, ($i - 1), $f)
    }

    $scanWatch.Stop()
    Write-Output ("PHASE|SCAN_SECONDS={0:N2}" -f $scanWatch.Elapsed.TotalSeconds)

    $copiedCount = $i - 1

    # Build DICOMDIR (scan media subdir under DEST, write DEST\DICOMDIR).
    # Use Start-Process with redirected streams to avoid PowerShell native stderr promotion.
    $dicomdirWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()
    try {
        $dcmmkdirStartParams = @{
            FilePath               = $dcmmkdir
            ArgumentList           = @('+r', '+id', '.', '+D', $dicomdirPath, '-v', $Subdir)
            WorkingDirectory       = $DestFull
            NoNewWindow            = $true
            Wait                   = $true
            PassThru               = $true
            RedirectStandardOutput = $stdoutFile
            RedirectStandardError  = $stderrFile
        }

        $proc = Start-Process @dcmmkdirStartParams

        foreach ($line in (Get-Content -LiteralPath $stdoutFile -ErrorAction SilentlyContinue)) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Output $line
            }
        }

        foreach ($line in (Get-Content -LiteralPath $stderrFile -ErrorAction SilentlyContinue)) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Output $line
            }
        }

        if ($proc.ExitCode -ne 0) {
            throw "dcmmkdir failed with exit code $($proc.ExitCode)"
        }
    }
    finally {
        Remove-Item -LiteralPath $stdoutFile -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrFile -Force -ErrorAction SilentlyContinue
        $dicomdirWatch.Stop()
        Write-Output ("PHASE|DICOMDIR_SECONDS={0:N2}" -f $dicomdirWatch.Elapsed.TotalSeconds)
    }
}

$weasisPackage = $null
if ($PackageWeasis) {
    $weasisPackage = Install-WeasisPackage -SourcePath $WeasisSource -DestinationRoot $DestFull -MediaSubdir $Subdir -DicomdirPath $dicomdirPath
}

if ($VerifyDicomdir) {
    if (-not (Test-Path -LiteralPath $dcmdump -PathType Leaf)) {
        throw "Missing DCMTK binary required for verification: $dcmdump"
    }

    Invoke-DicomdirVerification -DicomdirPath $dicomdirPath -MediaSubdir $Subdir -DcmDumpExe $dcmdump
}

Write-Host "Done."
if ($WeasisOnly) {
    Write-Host "Mode: Weasis package only"
}
else {
    Write-Host "Copied DICOM files: $copiedCount"
    $dicomdirRefCount = @($dicomdirRefs).Count
    Write-Host "DICOMDIR references discovered: $dicomdirRefCount"
    if ($Deidentify) {
        Write-Host "De-identified files: $deidentifiedCount"
    }
}
Write-Host "Media folder: $mediaDir"
Write-Host "DICOMDIR: $dicomdirPath"
if (-not $WeasisOnly) {
    Write-Host "Catalogue: $catalog"
}
if ($PackageWeasis) {
    Write-Host "Weasis folder: $($weasisPackage.WeasisDir)"
    Write-Host "Weasis executable: $($weasisPackage.WeasisExe)"
    Write-Host "Launch script: $($weasisPackage.LauncherPs1)"
    Write-Host "Launch shortcut: $($weasisPackage.LauncherCmd)"
}