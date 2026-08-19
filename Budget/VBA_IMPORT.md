# Importing exported VBA into Excel

Use `import_exported_vba.py` on Windows to replace the VBA modules and UserForms in a macro-enabled Excel workbook from an `exported_vba` folder. The script preserves Excel document objects such as `ThisWorkbook` and worksheet modules, replacing their code in place when a matching `.cls` export exists.

## Requirements

- Windows with desktop Microsoft Excel installed.
- Python 3.10 or newer.
- `pywin32`, installed with:

  ```powershell
  py -m pip install pywin32
  ```

- In Excel, enable **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**.
- Close the target workbook in Excel before running the importer.
- Restore any `<REDACTED>` values in a private local copy of the exported source. The importer stops before opening or changing the workbook if placeholders remain.

## Preview and import

For simple repeated use, open `Budget/import_exported_vba.py` and set the two variables near the top of the file:

```python
TARGET_WORKBOOK_PATH = r"C:\Path\Budget.xlsm"
EXPORTED_VBA_FOLDER_PATH = r"C:\Path\exported_vba"
BACKUP_FOLDER_PATH = r"C:\Path\Backups"
```

Then run `py Budget\import_exported_vba.py`. Command-line paths take priority over these variables. A blank workbook or source variable opens the corresponding Windows selection dialog. A blank backup variable stores the timestamped backup beside the target workbook.

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
  --source "C:\Path\exported_vba"
```

Use `--backup-folder "C:\Path\Backups"` to override the configured backup folder. The destination folder must already exist.

If either path is omitted, the script opens a Windows selection dialog. Before changing the VBA project, it displays the component counts and requires `REPLACE` as confirmation. Use `--yes` only for unattended execution.

## What the importer does

1. Validates the workbook type, exported component names, `.frm`/`.frx` pairs, redaction state, VBA project access, and project protection.
2. Opens Excel with events, alerts, external-link updates, and automatic macro execution disabled.
3. Saves the current workbook and creates a timestamped backup beside it, such as `Budget_backup_20260819_143000.xlsm`.
4. Deletes every removable standard module, ordinary class module, and UserForm from the target project. Excel's built-in `ThisWorkbook` and worksheet document objects are never deleted.
5. Saves the stripped workbook, closes it, and quits that Excel process so VBA releases cached component names and form designers.
6. Starts a new Excel process, reopens the stripped workbook, and verifies that the removed components did not reappear.
7. Imports each `.bas`, ordinary `.cls`, and `.frm`. Keep each `.frm` beside its same-name `.frx`; the VBA editor loads the `.frx` automatically.
8. For a document export such as `ThisWorkbook.cls`, clears the matching existing document object's code and inserts the exported source without its file-only attributes.
9. Verifies the final removable component set and saves the completed workbook.

If deletion fails before the stripped workbook is saved, the script closes it without saving the partial removal. If the later fresh-session import fails, the target remains in its intentionally saved stripped state and the timestamped backup retains the original project. The script does not automatically overwrite the target with that backup.

When Excel creates a form-import diagnostic such as `addItems.log`, the importer moves it out of `exported_vba` and places it beside `import_exported_vba.py`. Logs receive timestamps, for example `addItems_20260819_105543.log`, so a later run does not overwrite an earlier diagnostic. Existing logs in the source folder are moved before replacement begins, and a newly generated log is moved after Excel closes.

## Manual fallback

To import without the script, remove old standard modules and UserForms in the VBA editor, import every `.bas`, and import each `.frm` rather than its `.frx`. Do not import `ThisWorkbook.cls` or worksheet `.cls` files as ordinary classes. Instead, copy their executable code into the matching existing workbook or worksheet object.

The Budget UserForms require Microsoft Forms 2.0 and Microsoft Windows Common Controls 6.0 (`MSCOMCTL.OCX`). The projects also reference OLE Automation, the Microsoft Office object library, and Microsoft XML 6.0.
