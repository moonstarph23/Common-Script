#!/usr/bin/env python3
"""Replace an Excel workbook's VBA project from an exported_vba folder.

This script requires Windows, Microsoft Excel, pywin32, and Excel's
"Trust access to the VBA project object model" setting.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Sequence


# Optional default paths for running this file without command-line arguments.
# Example: TARGET_WORKBOOK_PATH = r"C:\Budget\Budget.xlsm"
# Example: EXPORTED_VBA_FOLDER_PATH = r"C:\Budget\exported_vba"
# Example: BACKUP_FOLDER_PATH = r"C:\Budget\Backups"
# Leave either input value blank to select it from a Windows dialog. Leave the
# backup value blank to store timestamped backups beside the target workbook.
TARGET_WORKBOOK_PATH = r""
EXPORTED_VBA_FOLDER_PATH = r""
BACKUP_FOLDER_PATH = r""
SCRIPT_FOLDER = Path(__file__).resolve().parent

SUPPORTED_WORKBOOK_EXTENSIONS = {".xlsm", ".xlsb", ".xlam", ".xltm"}

VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MS_FORM = 3
VBEXT_CT_DOCUMENT = 100
VBEXT_PP_LOCKED = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

VB_NAME_RE = re.compile(
    r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', re.IGNORECASE | re.MULTILINE
)
OLE_OBJECT_BLOB_RE = re.compile(
    r'^\s*OleObjectBlob\s*=\s*"([^"]+\.frx)":0000\s*$',
    re.IGNORECASE | re.MULTILINE,
)
VB_BASE_RE = re.compile(r"^\s*Attribute\s+VB_Base\s*=", re.IGNORECASE | re.MULTILINE)
VB_PREDECLARED_TRUE_RE = re.compile(
    r"^\s*Attribute\s+VB_PredeclaredId\s*=\s*True\s*$",
    re.IGNORECASE | re.MULTILINE,
)


class VbaImportError(RuntimeError):
    """An expected validation or Excel automation failure."""


@dataclass(frozen=True)
class SourceComponent:
    path: Path
    name: str


@dataclass(frozen=True)
class SourceInventory:
    standard_modules: tuple[SourceComponent, ...]
    class_modules: tuple[SourceComponent, ...]
    forms: tuple[SourceComponent, ...]

    @property
    def all_components(self) -> tuple[SourceComponent, ...]:
        return self.standard_modules + self.class_modules + self.forms


@dataclass(frozen=True)
class ReplacementPlan:
    removable_names: tuple[str, ...]
    standard_modules: tuple[SourceComponent, ...]
    ordinary_classes: tuple[SourceComponent, ...]
    forms: tuple[SourceComponent, ...]
    document_classes: tuple[SourceComponent, ...]


def read_vba_text(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise VbaImportError(f"Cannot decode VBA text file: {path}")


def extract_vb_name(path: Path, text: str | None = None) -> str:
    source = read_vba_text(path) if text is None else text
    match = VB_NAME_RE.search(source)
    if not match:
        raise VbaImportError(f"Missing Attribute VB_Name header: {path}")
    return match.group(1)


def document_code_body(path: Path) -> str:
    """Return executable source suitable for an existing document CodeModule."""
    lines = read_vba_text(path).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    index = 0

    while index < len(lines) and not lines[index].strip():
        index += 1
    if index < len(lines) and lines[index].lstrip().upper().startswith("VERSION "):
        index += 1

    while index < len(lines) and not lines[index].strip():
        index += 1
    if index < len(lines) and lines[index].strip().upper() == "BEGIN":
        index += 1
        while index < len(lines) and lines[index].strip().upper() != "END":
            index += 1
        if index >= len(lines):
            raise VbaImportError(f"Unterminated class metadata block: {path}")
        index += 1

    while index < len(lines):
        stripped = lines[index].strip()
        if not stripped or stripped.upper().startswith("ATTRIBUTE "):
            index += 1
            continue
        break

    body = "\r\n".join(lines[index:]).strip("\r\n")
    if not body.strip():
        raise VbaImportError(f"Document class contains no executable source: {path}")
    return body + "\r\n"


def looks_like_document_class(path: Path) -> bool:
    source = read_vba_text(path)
    return bool(VB_BASE_RE.search(source) and VB_PREDECLARED_TRUE_RE.search(source))


def _files_by_suffix(folder: Path, suffix: str) -> list[Path]:
    return sorted(
        (path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == suffix),
        key=lambda path: path.name.casefold(),
    )


def build_source_inventory(folder: Path) -> SourceInventory:
    if not folder.is_dir():
        raise VbaImportError(f"Source folder does not exist: {folder}")

    paths = {
        ".bas": _files_by_suffix(folder, ".bas"),
        ".cls": _files_by_suffix(folder, ".cls"),
        ".frm": _files_by_suffix(folder, ".frm"),
        ".frx": _files_by_suffix(folder, ".frx"),
    }
    if not paths[".bas"] and not paths[".cls"] and not paths[".frm"]:
        raise VbaImportError(f"No .bas, .cls, or .frm components found in: {folder}")

    redacted = []
    components: dict[str, list[SourceComponent]] = {".bas": [], ".cls": [], ".frm": []}
    names_seen: dict[str, Path] = {}
    for suffix in (".bas", ".cls", ".frm"):
        for path in paths[suffix]:
            text = read_vba_text(path)
            if "<REDACTED>" in text:
                redacted.append(path.name)
            name = extract_vb_name(path, text)
            if name.casefold() != path.stem.casefold():
                raise VbaImportError(
                    f"File name and Attribute VB_Name do not match: {path.name} ({name})"
                )
            duplicate = names_seen.get(name.casefold())
            if duplicate:
                raise VbaImportError(
                    f"Duplicate VBA component name {name!r}: {duplicate.name}, {path.name}"
                )
            names_seen[name.casefold()] = path
            components[suffix].append(SourceComponent(path=path, name=name))

    if redacted:
        joined = ", ".join(sorted(redacted, key=str.casefold))
        raise VbaImportError(
            "Source contains <REDACTED> placeholders. Restore the private values in a local "
            f"copy before importing: {joined}"
        )

    frx_by_stem = {path.stem.casefold(): path for path in paths[".frx"]}
    form_stems = {component.path.stem.casefold() for component in components[".frm"]}
    for form in components[".frm"]:
        frx = frx_by_stem.get(form.path.stem.casefold())
        if frx is None:
            raise VbaImportError(f"Missing FRX companion for form: {form.path.name}")
        frx_data = frx.read_bytes()
        if (
            len(frx_data) < 32
            or frx_data[:2] != b"LB"
            or frx_data[24:32] != bytes.fromhex("d0cf11e0a1b11ae1")
        ):
            raise VbaImportError(f"FRX companion is not a valid UserForm binary: {frx.name}")
        blob_match = OLE_OBJECT_BLOB_RE.search(read_vba_text(form.path))
        if not blob_match or blob_match.group(1).casefold() != frx.name.casefold():
            raise VbaImportError(
                f"Form does not reference its matching FRX companion: {form.path.name}"
            )

    orphan_frx = [path.name for path in paths[".frx"] if path.stem.casefold() not in form_stems]
    if orphan_frx:
        raise VbaImportError("FRX file has no matching FRM: " + ", ".join(orphan_frx))

    return SourceInventory(
        standard_modules=tuple(components[".bas"]),
        class_modules=tuple(components[".cls"]),
        forms=tuple(components[".frm"]),
    )


def next_backup_path(
    workbook_path: Path,
    backup_folder: Path | None = None,
    now: datetime | None = None,
) -> Path:
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    destination = backup_folder or workbook_path.parent
    candidate = destination / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    counter = 2
    while candidate.exists():
        candidate = destination / (
            f"{workbook_path.stem}_backup_{timestamp}_{counter}{workbook_path.suffix}"
        )
        counter += 1
    return candidate


def next_import_log_path(
    source_log: Path,
    destination_folder: Path = SCRIPT_FOLDER,
    now: datetime | None = None,
) -> Path:
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    candidate = destination_folder / f"{source_log.stem}_{timestamp}.log"
    counter = 2
    while candidate.exists():
        candidate = destination_folder / f"{source_log.stem}_{timestamp}_{counter}.log"
        counter += 1
    return candidate


def relocate_excel_import_logs(
    source_folder: Path,
    *,
    destination_folder: Path = SCRIPT_FOLDER,
    strict: bool,
) -> tuple[Path, ...]:
    moved = []
    try:
        logs = sorted(
            (
                path
                for path in source_folder.iterdir()
                if path.is_file() and path.suffix.casefold() == ".log"
            ),
            key=lambda path: path.name.casefold(),
        )
    except Exception as exc:
        message = f"Unable to inspect Excel import logs in {source_folder}: {exc}"
        if strict:
            raise VbaImportError(message) from exc
        print(f"WARNING: {message}", file=sys.stderr)
        return ()
    for source_log in logs:
        destination = next_import_log_path(source_log, destination_folder)
        try:
            shutil.move(str(source_log), str(destination))
        except Exception as exc:
            message = f"Unable to move Excel import log {source_log} to {destination}: {exc}"
            if strict:
                raise VbaImportError(message) from exc
            print(f"WARNING: {message}", file=sys.stderr)
            continue
        moved.append(destination)
        print(f"Excel import log: {destination}")
    return tuple(moved)


def choose_workbook() -> Path:
    try:
        from tkinter import Tk, filedialog
    except ImportError as exc:
        raise VbaImportError("--workbook is required because tkinter is unavailable") from exc
    try:
        root = Tk()
        root.withdraw()
        selected = filedialog.askopenfilename(
            title="Select the target macro-enabled Excel workbook",
            filetypes=[
                ("Macro-enabled Excel workbooks", "*.xlsm *.xlsb *.xlam *.xltm"),
                ("All files", "*.*"),
            ],
        )
        root.destroy()
    except Exception as exc:
        raise VbaImportError("Unable to open the workbook selection dialog") from exc
    if not selected:
        raise VbaImportError("No target workbook selected")
    return Path(selected)


def choose_source_folder() -> Path:
    try:
        from tkinter import Tk, filedialog
    except ImportError as exc:
        raise VbaImportError("--source is required because tkinter is unavailable") from exc
    try:
        root = Tk()
        root.withdraw()
        selected = filedialog.askdirectory(title="Select the exported_vba source folder")
        root.destroy()
    except Exception as exc:
        raise VbaImportError("Unable to open the source-folder selection dialog") from exc
    if not selected:
        raise VbaImportError("No exported_vba folder selected")
    return Path(selected)


def validate_workbook_path(path: Path) -> Path:
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise VbaImportError(f"Target workbook does not exist: {resolved}")
    if resolved.suffix.lower() not in SUPPORTED_WORKBOOK_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_WORKBOOK_EXTENSIONS))
        raise VbaImportError(f"Unsupported workbook type {resolved.suffix!r}; expected {supported}")
    return resolved


def configured_path(value: str) -> Path | None:
    value = value.strip()
    return Path(value) if value else None


def vb_components(vb_project: Any) -> list[Any]:
    collection = vb_project.VBComponents
    return [collection.Item(index) for index in range(1, collection.Count + 1)]


def make_replacement_plan(vb_project: Any, inventory: SourceInventory) -> ReplacementPlan:
    existing = vb_components(vb_project)
    documents = {
        str(component.Name).casefold(): component
        for component in existing
        if int(component.Type) == VBEXT_CT_DOCUMENT
    }
    removable = tuple(
        str(component.Name)
        for component in existing
        if int(component.Type) in (VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM)
    )

    document_classes = []
    ordinary_classes = []
    for component in inventory.class_modules:
        if component.name.casefold() in documents:
            # Fail before the backup/deletion stage if the exported document
            # source cannot be inserted into an existing CodeModule.
            document_code_body(component.path)
            document_classes.append(component)
        elif looks_like_document_class(component.path):
            raise VbaImportError(
                f"Document class {component.name!r} has no matching workbook/worksheet object "
                "in the target workbook"
            )
        else:
            ordinary_classes.append(component)

    imported = inventory.standard_modules + tuple(ordinary_classes) + inventory.forms
    conflicts = [component.name for component in imported if component.name.casefold() in documents]
    if conflicts:
        raise VbaImportError(
            "Importable component names conflict with existing document objects: "
            + ", ".join(conflicts)
        )

    return ReplacementPlan(
        removable_names=tuple(sorted(removable, key=str.casefold)),
        standard_modules=inventory.standard_modules,
        ordinary_classes=tuple(ordinary_classes),
        forms=inventory.forms,
        document_classes=tuple(document_classes),
    )


def print_plan(
    workbook_path: Path,
    source_folder: Path,
    backup_folder: Path | None,
    plan: ReplacementPlan,
) -> None:
    print(f"Target workbook: {workbook_path}")
    print(f"Source folder:   {source_folder}")
    print(f"Backup folder:   {backup_folder or workbook_path.parent}")
    print(f"Remove:          {len(plan.removable_names)} standard/class/form components")
    print(f"Import:          {len(plan.standard_modules)} BAS module(s)")
    print(f"Import:          {len(plan.ordinary_classes)} ordinary CLS module(s)")
    print(f"Import:          {len(plan.forms)} FRM/FRX form pair(s)")
    print(f"Replace in place:{len(plan.document_classes):2d} document CLS module(s)")


def confirm_replacement(assume_yes: bool) -> None:
    if assume_yes:
        return
    if not sys.stdin.isatty():
        raise VbaImportError("Interactive confirmation is unavailable; rerun with --yes")
    answer = input("Type REPLACE to save a backup and rebuild the VBA project: ").strip()
    if answer != "REPLACE":
        raise VbaImportError("Replacement cancelled; no changes were made")


def remove_replaceable_components(vb_project: Any) -> int:
    removable = [
        component
        for component in vb_components(vb_project)
        if int(component.Type) in (VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM)
    ]
    for component in removable:
        vb_project.VBComponents.Remove(component)
    return len(removable)


def import_component(vb_project: Any, component: SourceComponent) -> None:
    imported = vb_project.VBComponents.Import(str(component.path))
    if str(imported.Name).casefold() != component.name.casefold():
        raise VbaImportError(
            f"Excel imported {component.path.name} as {imported.Name!r}, expected {component.name!r}"
        )


def replace_document_code(vb_project: Any, component: SourceComponent) -> None:
    target = None
    for existing in vb_components(vb_project):
        if int(existing.Type) == VBEXT_CT_DOCUMENT and str(existing.Name).casefold() == component.name.casefold():
            target = existing
            break
    if target is None:
        raise VbaImportError(
            f"No existing document component matches {component.name!r}; it cannot be imported normally"
        )

    body = document_code_body(component.path)
    code_module = target.CodeModule
    line_count = int(code_module.CountOfLines)
    if line_count:
        code_module.DeleteLines(1, line_count)
    code_module.AddFromString(body)

    written = str(code_module.Lines(1, code_module.CountOfLines)).replace("\r\n", "\n").strip()
    expected = body.replace("\r\n", "\n").strip()
    if written != expected:
        raise VbaImportError(f"Excel did not preserve the expected code for {component.name}")


def verify_final_components(vb_project: Any, plan: ReplacementPlan) -> None:
    expected = {
        component.name.casefold()
        for component in plan.standard_modules + plan.ordinary_classes + plan.forms
    }
    actual = {
        str(component.Name).casefold()
        for component in vb_components(vb_project)
        if int(component.Type) in (VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM)
    }
    if actual != expected:
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        raise VbaImportError(
            f"Final VBA component set does not match the source; missing={missing}, extra={extra}"
        )


def create_excel_application(win32com_client: Any) -> Any:
    excel = win32com_client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    return excel


def open_target_workbook(excel: Any, workbook_path: Path, *, read_only: bool) -> Any:
    workbook = excel.Workbooks.Open(
        Filename=str(workbook_path),
        UpdateLinks=0,
        ReadOnly=read_only,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
    )
    if not read_only and bool(workbook.ReadOnly):
        raise VbaImportError(
            "Excel opened the workbook read-only. Close it in other Excel sessions and retry."
        )
    return workbook


def accessible_vb_project(workbook: Any) -> Any:
    try:
        vb_project = workbook.VBProject
        _ = vb_project.VBComponents.Count
    except Exception as exc:
        raise VbaImportError(
            "Excel denied VBA project access. Enable 'Trust access to the VBA project "
            "object model' in Trust Center and ensure the workbook has a VBA project."
        ) from exc
    if int(vb_project.Protection) == VBEXT_PP_LOCKED:
        raise VbaImportError("The workbook VBA project is locked; unlock it before importing")
    return vb_project


def run_excel_import(
    workbook_path: Path,
    source_folder: Path,
    inventory: SourceInventory,
    *,
    backup_folder: Path | None,
    dry_run: bool,
    assume_yes: bool,
) -> None:
    if sys.platform != "win32":
        raise VbaImportError("Excel VBA import requires Windows and a desktop installation of Excel")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise VbaImportError(
            "pywin32 is required. Install it with: py -m pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    vb_project = None
    backup_path: Path | None = None
    mutation_started = False
    removal_saved = False
    saved_successfully = False
    try:
        excel = create_excel_application(win32com.client)
        workbook = open_target_workbook(excel, workbook_path, read_only=dry_run)
        vb_project = accessible_vb_project(workbook)

        plan = make_replacement_plan(vb_project, inventory)
        print_plan(workbook_path, source_folder, backup_folder, plan)
        if dry_run:
            print("Dry run complete; the workbook was opened read-only and was not modified.")
            return

        confirm_replacement(assume_yes)
        workbook.Save()
        backup_path = next_backup_path(workbook_path, backup_folder)
        workbook.SaveCopyAs(str(backup_path))
        print(f"Backup created:  {backup_path}")

        # Preserve diagnostics from any earlier run and keep exported_vba
        # limited to importable components before Excel can create a new log.
        relocate_excel_import_logs(source_folder, strict=True)

        mutation_started = True
        removed_count = remove_replaceable_components(vb_project)
        workbook.Save()
        removal_saved = True

        # Closing the stripped workbook and its Excel process clears VBA's
        # component-name/designer cache before the replacement import begins.
        vb_project = None
        workbook.Close(SaveChanges=False)
        workbook = None
        excel.Quit()
        excel = None
        print("Old components removed and saved. Reopening in a new Excel session...")

        excel = create_excel_application(win32com.client)
        workbook = open_target_workbook(excel, workbook_path, read_only=False)
        vb_project = accessible_vb_project(workbook)
        remaining = [
            str(component.Name)
            for component in vb_components(vb_project)
            if int(component.Type)
            in (VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM)
        ]
        if remaining:
            raise VbaImportError(
                "Removed components reappeared after reopening the workbook: "
                + ", ".join(sorted(remaining, key=str.casefold))
            )

        for component in plan.standard_modules:
            import_component(vb_project, component)
        for component in plan.ordinary_classes:
            import_component(vb_project, component)
        for component in plan.forms:
            import_component(vb_project, component)
        for component in plan.document_classes:
            replace_document_code(vb_project, component)
        verify_final_components(vb_project, plan)

        workbook.Save()
        saved_successfully = True
        print(f"Removed:         {removed_count} old component(s)")
        print(f"Imported:        {len(plan.standard_modules) + len(plan.ordinary_classes)} module(s)")
        print(f"Imported:        {len(plan.forms)} UserForm(s) with FRX companions")
        print(f"Updated:         {len(plan.document_classes)} document class module(s)")
        print(f"Workbook saved:  {workbook_path}")
    except VbaImportError:
        raise
    except Exception as exc:
        if removal_saved:
            state = " during the fresh-session import"
        elif mutation_started:
            state = " after component removal began"
        else:
            state = ""
        raise VbaImportError(f"Excel VBA import failed{state}: {exc}") from exc
    finally:
        vb_project = None
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if backup_path is not None:
            relocate_excel_import_logs(source_folder, strict=False)
        pythoncom.CoUninitialize()
        if mutation_started and not saved_successfully and backup_path is not None:
            if removal_saved:
                print(
                    "Import failed after the stripped workbook was saved. The target remains "
                    "without its old removable components; restore the original project from "
                    f"this backup if needed: {backup_path}",
                    file=sys.stderr,
                )
            else:
                print(
                    "Import failed before the removal phase was saved. The workbook was closed "
                    f"without saving partial changes. Backup: {backup_path}",
                    file=sys.stderr,
                )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Back up an Excel workbook and replace its VBA modules/forms from exported_vba."
    )
    parser.add_argument("--workbook", type=Path, help="Target .xlsm/.xlsb/.xlam/.xltm workbook")
    parser.add_argument("--source", type=Path, help="Source exported_vba folder")
    parser.add_argument(
        "--backup-folder",
        type=Path,
        help="Backup destination; defaults to the target workbook folder",
    )
    parser.add_argument("--dry-run", action="store_true", help="Validate and preview without changes")
    parser.add_argument("--yes", action="store_true", help="Skip the REPLACE confirmation prompt")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        workbook_path = validate_workbook_path(
            args.workbook or configured_path(TARGET_WORKBOOK_PATH) or choose_workbook()
        )
        source_folder = (
            args.source
            or configured_path(EXPORTED_VBA_FOLDER_PATH)
            or choose_source_folder()
        ).expanduser().resolve()
        configured_backup = args.backup_folder or configured_path(BACKUP_FOLDER_PATH)
        backup_folder = configured_backup.expanduser().resolve() if configured_backup else None
        if backup_folder is not None and not backup_folder.is_dir():
            raise VbaImportError(f"Backup folder does not exist: {backup_folder}")
        inventory = build_source_inventory(source_folder)
        run_excel_import(
            workbook_path,
            source_folder,
            inventory,
            backup_folder=backup_folder,
            dry_run=args.dry_run,
            assume_yes=args.yes,
        )
        return 0
    except VbaImportError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
