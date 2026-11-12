# Enterprise File Manager (EFM)

EFM is a portable Windows desktop application for cataloguing a folder of project documents. It is designed to run without installation or administrator access.

## Features

- Browse and search documents in a selected folder
- Store document metadata locally in a `filedata.efm` file
- Record title, type, number, issue date, sheet information, version, and status
- Open files through their normal Windows application
- Link associated documents and version-history documents, with in-app navigation and Back support

## Running the application

Download a packaged `EFM.exe` from the project releases, then run it directly on 64-bit Windows 10 or Windows 11.

1. Select a document folder.
2. Select a document from the table.
3. Edit its metadata and choose **Update Details**.
4. Use **Open File** for PDFs and other formats.

The selected folder owns its own metadata: EFM creates or updates `filedata.efm` alongside the documents. EFM does not rename, move, or delete managed files.

## Development and packaging

EFM is packaged as a single executable with PyInstaller using `EFM.spec`.

```powershell
python -m pip install --target buildtools pyinstaller
$env:PYTHONPATH = (Resolve-Path buildtools).Path
python -m PyInstaller --noconfirm --clean --distpath dist --workpath pyinstaller_build EFM.spec
```

The build targets 64-bit Windows.

## Data compatibility

Existing `filedata.efm` files from earlier EFM versions remain readable. New fields are written when a document is saved.

## Licence

No licence has been chosen yet. Add one before distributing this repository publicly.
