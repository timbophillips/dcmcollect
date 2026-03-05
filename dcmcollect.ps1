$SRC  = "C:\path\to\source_tree"
$DEST = "C:\path\to\output_media"
$SUB  = "IMAGES"

New-Item -ItemType Directory -Force -Path (Join-Path $DEST $SUB) | Out-Null

$catalog = Join-Path $DEST "catalogue.csv"
"seq,new_filename,source_path" | Set-Content -Encoding ASCII $catalog

$i = 1
Get-ChildItem -Path $SRC -Recurse -File | ForEach-Object {
    $f = $_.FullName
    & .\bin\dcmftest.exe $f *> $null
    if ($LASTEXITCODE -eq 0) {
        $new = "{0:D8}" -f $i
        $outRel = Join-Path $SUB $new
        $outAbs = Join-Path $DEST $outRel
        Copy-Item -LiteralPath $f -Destination $outAbs
        "$i,$outRel,""{0}""" -f $f.Replace('"','""') | Add-Content -Encoding ASCII $catalog
        $i++
    }
}

# Build DICOMDIR (scan output_media recursively, write output_media\DICOMDIR)
& .\bin\dcmmkdir.exe +r +id $DEST +p "*" +D (Join-Path $DEST "DICOMDIR") $SUB -v