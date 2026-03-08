# dcmcollect

Collect DICOM files from a source folder tree into a flat media folder and generate a DICOMDIR using DCMTK.

This repository provides a PowerShell script, `Collect-DicomMedia.ps1`, that:

- Recursively scans a source tree for candidate files.
- Detects DICOM using a two-step check:
  - Fast check with `dcmftest` (strict Part-10).
  - Fallback parse with `dcmdump` for valid dataset files that are not Part-10 wrapped.
- Reads every `DICOMDIR` found in the source tree and adds referenced files to collection candidates.
- De-duplicates discovered files (case-insensitive paths).
- Copies accepted files into a sequentially named media folder.
- Writes `catalogue.csv` mapping output files back to source paths.
- Builds a destination `DICOMDIR` with `dcmmkdir`.
- Optionally bundles Weasis into the destination and creates one-click launchers.
- Supports a Weasis-only mode for already prepared destination folders.

## Requirements

- Windows PowerShell or PowerShell 7+
- DCMTK binaries in `bin/` next to the script:
  - `bin/dcmftest.exe`
  - `bin/dcmdump.exe`
  - `bin/dcmmkdir.exe`

## Repository Layout

```text
.
|-- Collect-DicomMedia.ps1
|-- bin/
`-- README.md
```

## Usage

```powershell
.\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media"
```

Optional media subfolder name:

```powershell
.\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -Subdir "IMAGES"
```

Bundle Weasis from a local zip or folder:

```powershell
.\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -PackageWeasis
```

Or provide an explicit package path:

```powershell
.\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -PackageWeasis -WeasisSource "C:\tools\weasis-portable.zip"
```

Package Weasis only into an existing destination folder (no DICOM copy/rebuild):

```powershell
.\Collect-DicomMedia.ps1 -Dest "C:\output_media" -WeasisOnly
```

Run a post-build DICOMDIR reference verification:

```powershell
.\Collect-DicomMedia.ps1 -Src "C:\input" -Dest "C:\output_media" -VerifyDicomdir
```

### Parameters

- `-Src` (required): Source root directory to scan.
- `-Dest` (required): Output root directory.
- `-Subdir` (optional, default: `IMAGES`): Subfolder under `-Dest` where copied DICOM files are placed.
- `-PackageWeasis` (optional): If set, package Weasis into the destination folder.
- `-WeasisSource` (optional): Path to a Weasis folder or `.zip` package.
  - If omitted and `-PackageWeasis` is set, the script uses `weasis-portable.zip` from the script folder.
- `-WeasisOnly` (optional): Package Weasis and launchers only. Skips source scan, copy, catalogue, and DICOMDIR generation.
  - Requires `-Dest` to already exist.
  - Implies `-PackageWeasis`.
- `-VerifyDicomdir` (optional): After processing, parse destination `DICOMDIR` and:
  - print a sample of `ReferencedFileID` entries,
  - warn when any entry does not start with the configured `-Subdir` prefix.

## Output

Given `-Dest C:\output_media` and default `-Subdir IMAGES`, the script creates:

- `C:\output_media\IMAGES\00000001`, `00000002`, ... (copied DICOM files)
- `C:\output_media\catalogue.csv`
- `C:\output_media\DICOMDIR`

If `-PackageWeasis` is used:

- `C:\output_media\weasis\...` (copied/extracted Weasis files)
- `C:\output_media\Launch-Weasis.ps1`
- `C:\output_media\Launch-Weasis.cmd`

If `-WeasisOnly` is used, existing DICOM content is left as-is and only the Weasis package/launchers are updated.

`catalogue.csv` columns:

- `seq`: Sequence number assigned during copy.
- `new_filename`: Relative output path (`IMAGES\00000001`, etc.).
- `source_path`: Original absolute source path.

## How DICOM Detection Works

The script accepts a file if either check succeeds:

1. `dcmftest <file>` returns exit code `0` (strict Part-10).
2. `dcmdump -q +P "SOPClassUID" <file>` returns exit code `0`.

This helps include valid DICOM dataset files that may not have a Part-10 file meta header.

## DICOMDIR-Assisted Discovery

In addition to filesystem recursion, the script:

- Finds all files named `DICOMDIR` under `-Src`.
- Extracts each `(0004,1500) ReferencedFileID` entry.
- Resolves each reference relative to that `DICOMDIR` location.
- Adds existing referenced files to candidate collection.

This improves completeness when source media includes index files.

## Weasis Packaging

When packaging is enabled, the script:

- Copies or extracts Weasis under `Dest\weasis`.
- Locates a `weasis*.exe` executable inside that folder.
- Generates launchers at destination root:
  - `Launch-Weasis.cmd` for end users
  - `Launch-Weasis.ps1` for direct PowerShell use

The launcher opens `Dest\DICOMDIR` when available, and falls back to `Dest\Subdir` if needed.

## Notes

- Paths are handled with `-LiteralPath` where appropriate.
- The script runs `dcmmkdir` from inside `-Dest` so relative IDs resolve correctly.
- Existing output files with matching sequence names are overwritten (`-Force`).

## Troubleshooting

- `Missing DCMTK binary: ...`
  - Ensure required executables are present in `bin/`.

- Fewer files than expected:
  - Confirm source tree is correct.
  - Check whether files are valid DICOM or referenced by source `DICOMDIR`.
  - Review `catalogue.csv` to see exactly what was copied.

- `dcmmkdir` warnings/errors:
  - Verify copied files under `Dest\Subdir` are readable and DICOM-compliant.

- `Unable to find Weasis executable under: ...`
  - Confirm your package contains a Windows launcher executable (for example `weasis*.exe` or `viewer-win32.exe`) in its extracted tree.

## License

Open source
