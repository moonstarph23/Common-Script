# VBA import and export guide

This guide defines the `exported_vba` folder format expected by `import_exported_vba.py` and the correct way to export VBA components so Excel can import them without name or `OleObjectBlob` errors.

## Required `exported_vba` contents

Keep only importable VBA source files and required UserForm binary companions in the folder.

| Component in Excel | Exported file | Requirements |
|---|---|---|
| Standard module | `.bas` | Include `Attribute VB_Name`, use the same file basename as the module name, and retain CRLF line endings. |
| Ordinary class module | `.cls` | Export the complete native class file, including its class descriptor and attributes. |
| `ThisWorkbook` or worksheet code module | `.cls` | Include the document-class attributes and executable source. Export only document modules that contain code; omit empty worksheet modules. |
| UserForm | `.frm` and same-name `.frx` | The `.frm` must contain the complete form descriptor, `OleObjectBlob` reference, attributes, and code. The `.frx` must be Excel's matching binary designer file. Keep the pair together. |

Example:

```text
exported_vba/
├── CallOverall.bas
├── Funcs.bas
├── ThisWorkbook.cls
├── addItems.frm
├── addItems.frx
├── viewPL.frm
└── viewPL.frx
```

Do not place these files in `exported_vba`:

- A complete `vbaProject.bin` or individually extracted `.bin` streams.
- `manifest.json`, `PROJECT.txt`, or `references.txt` unless metadata was specifically requested for another workflow.
- Excel import `.log` files; the importer moves them beside the Python script.
- Empty worksheet `.cls` files.
- Workbooks or temporary Office files.

Every component name must be unique, and each `.bas`, `.cls`, or `.frm` basename must match its `Attribute VB_Name` value. The importer rejects duplicates, mismatched names, orphan `.frx` files, missing form companions, invalid FRX record headers or payload lengths, invalid form bounds, missing OLE compound signatures, and tracked text containing `<REDACTED>`.

## Recommended export method: Excel VBA editor

Use the VBA editor on Windows whenever possible. Its native component export produces the correct form descriptor and FRX compound-file structure.

1. Work from a private copy of the source `.xlsm` or `.xlsb` workbook.
2. Open Excel and press **Alt+F11**.
3. In Project Explorer, select the component to export.
4. Use **File > Export File** or right-click the component and choose **Export File**.
5. Save the component directly into the intended `exported_vba` folder.

Export each component as follows:

- **Standard modules:** export as `.bas`.
- **Ordinary class modules:** export as `.cls`.
- **ThisWorkbook and coded worksheet objects:** export as `.cls`. These files are source for code replacement; they are not imported as ordinary classes.
- **UserForms:** export the `.frm`. Excel automatically creates the same-name `.frx`; copy and distribute both files together.

For a UserForm named `addItems`, the valid pair is:

```text
addItems.frm
addItems.frx
```

The `.frm` must include records similar to:

```vb
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addItems
   OleObjectBlob   =   "addItems.frx":0000
   ' Native form properties...
End
Attribute VB_Name = "addItems"
```

Never rename only one member of a form pair, import the `.frx` directly, replace the `.frx` with raw `CompObj`/`f`/`o` streams, or construct a partial code-only `.frm`. Those approaches can produce `Property OleObjectBlob ... could not be set` during import.

A native VBE-exported `.frx` starts with a 24-byte `OleObjectBlob` record wrapper before the OLE compound file. That wrapper must identify an OLE storage record, declare the exact byte length of the compound payload, and contain valid form bounds. Copying a wrapper from another form or placing a bare compound file in the folder is not valid, even when an OLE reader can open the inner streams.

If Excel is unavailable, a UserForm export must still reproduce Excel's native format: build the `.frm` from the embedded `VBFrame` descriptor and VBA source, and build a valid FRX compound file containing the remaining designer storage. The CFB directory tree, CLSID, allocation tables, stream names, and all designer-stream bytes must remain valid. Prefer native Excel export because a binary that merely opens in a tolerant OLE reader can still be rejected by Excel.

## Credentials and private values

Tracked text exports must replace embedded credentials with `<REDACTED>`. Never commit the original values. Before importing, make a private local copy of `exported_vba` and restore the required values there. The importer stops before opening or changing the workbook if any `<REDACTED>` placeholder remains.

## Importer requirements

- Windows with desktop Microsoft Excel installed.
- Python 3.10 or newer.
- `pywin32`, installed with:

  ```powershell
  py -m pip install pywin32
  ```

- In Excel, enable **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**.
- Close the target workbook in all Excel sessions before running the importer.
- Unlock the VBA project before importing.

The Budget forms also require Microsoft Forms 2.0 and Microsoft Windows Common Controls 6.0 (`MSCOMCTL.OCX`). The projects reference OLE Automation, the Microsoft Office object library, and Microsoft XML 6.0.

## Configure paths

For repeated use, edit the variables near the top of `Budget/import_exported_vba.py`:

```python
TARGET_WORKBOOK_PATH = r"C:\Path\Budget.xlsm"
EXPORTED_VBA_FOLDER_PATH = r"C:\Path\exported_vba"
BACKUP_FOLDER_PATH = r"C:\Path\Backups"
```

Command-line paths override these variables. A blank workbook or source variable opens a Windows selection dialog. A blank backup variable stores the timestamped backup beside the target workbook.

## Preview and import

Run a read-only preview first:

```powershell
py Budget\import_exported_vba.py `
  --workbook "C:\Path\Budget.xlsm" `
  --source "C:\Path\exported_vba" `
  --dry-run
```

Run the replacement:

```powershell
py Budget\import_exported_vba.py `
  --workbook "C:\Path\Budget.xlsm" `
  --source "C:\Path\exported_vba" `
  --backup-folder "C:\Path\Backups"
```

Before changing the VBA project, the script displays the component counts and requires `REPLACE` as confirmation. Use `--yes` only for unattended execution.

## Import sequence

1. Validate the workbook, source files, form pairs, redaction state, project access, and project protection.
2. Open Excel with events, alerts, link updates, and automatic macro execution disabled.
3. Save the workbook and create a timestamped backup.
4. Delete all removable standard modules, ordinary classes, and UserForms. Preserve Excel's built-in `ThisWorkbook` and worksheet objects.
5. Save the stripped workbook, close it, and quit Excel so VBA releases cached component names and form designers.
6. Start a fresh Excel process and reopen the stripped workbook.
7. Import every `.bas`, ordinary `.cls`, and `.frm`. Excel reads each adjacent `.frx` automatically.
8. Replace code in matching document objects such as `ThisWorkbook` from their `.cls` source instead of importing those files as ordinary classes.
9. Verify the final component set and save the completed workbook.

If deletion fails before the stripped workbook is saved, the script closes without saving the partial removal. If the fresh-session import fails, the target remains in its saved stripped state and the timestamped backup retains the original project. The script never overwrites the target automatically from the backup.

Excel-generated diagnostics such as `addItems.log` are moved out of `exported_vba` and stored beside `import_exported_vba.py` with timestamped names.

## Manual import fallback

For manual migration:

1. Save a backup of the target workbook.
2. Remove its old standard modules, ordinary classes, and UserForms.
3. Save, close, and reopen the stripped workbook.
4. Import each `.bas` individually.
5. Keep every `.frm` beside its `.frx` and import the `.frm`, not the `.frx`.
6. Import ordinary `.cls` files normally.
7. For `ThisWorkbook.cls` or worksheet `.cls` files, copy only the executable code into the matching existing Excel document object. Importing a document `.cls` as an ordinary class does not bind workbook or worksheet events.
