# wait_opera_print_pdf.py

Waits for an `OperaPrint*.pdf` file (e.g. `OperaPrint[1].pdf`) to appear **fully downloaded** in the current user's Windows INetCache, optionally moving it to a target location once it's ready.

## Usage (CLI)

```bash
python wait_opera_print_pdf.py <timeout> [move_to_path]
```

- `timeout` — seconds to wait before giving up (integer).
- `move_to_path` — optional full file path (including filename) to move the verified PDF to. If omitted, the PDF is left in place.

### Examples

```bash
# Just wait up to 60s; print where the PDF was found
python wait_opera_print_pdf.py 60

# Wait 90s, then move the PDF to a specific file path (overwrites if exists)
python wait_opera_print_pdf.py 90 "R:\Reports\opera.pdf"
```

## Function signature

```python
wait_for_opera_print_pdf(timeout: int, move_to_path: str = None) -> tuple
```

## Parameters

| Parameter       | Type | Required | Description                                                                                         |
|-----------------|------|----------|-----------------------------------------------------------------------------------------------------|
| `timeout`       | int  | yes      | Seconds to wait before timing out.                                                                  |
| `move_to_path`  | str  | no       | Full file path (including filename) to move the PDF to. Overwrites if the destination already exists. |

## Return value

| Result                | Meaning                                                                                       |
|-----------------------|-----------------------------------------------------------------------------------------------|
| `(True, path)`        | PDF found and fully downloaded. `path` is the destination path if moved, otherwise the source. |
| `(True, source_path)` | PDF found and verified, but the move failed. The PDF remains at `source_path`.                |
| `(False, None)`       | Timed out; no fully downloaded `OperaPrint*.pdf` was found.                                    |

The function does **not** raise on failure — it returns a tuple so UiPath/callers can branch on the boolean.

## "Fully downloaded" criteria

A file is considered fully downloaded only when **all three** are true:

1. File exists and `size > 0`.
2. Size is unchanged between two `os.path.getsize` calls **2 seconds apart**.
3. The file is not locked by another process — confirmed by attempting `os.open(path, os.O_RDWR)`; if it raises `PermissionError`/`OSError`, the file is treated as still being written.

## Where it scans

- `%LOCALAPPDATA%\Microsoft\Windows\INetCache` — **recursively** (subfolders like `IE`, `Low`, etc.).
- `LOCALAPPDATA` falls back to `%USERPROFILE%\AppData\Local` if the env var is unset.
- Match is case-insensitive (`OperaPrint*.pdf` and `OPERAPRINT*.pdf`).

## Polling behavior

- Outer loop checks every **2 seconds** until `timeout` is reached.
- Per attempt, all current matches are stat'd; the stability/lock check runs only on files whose size was already recorded on a previous pass (so a freshly-appearing file is watched for one extra cycle before being considered done).
- Logs attempt number, file basename, size, and stability status to stdout.

## UiPath Run activity

In a UiPath **Run** / **Start Process** activity, set:

| Field             | Value                                                                                                                         |
|-------------------|-------------------------------------------------------------------------------------------------------------------------------|
| FileName          | `python.exe` (or full path, e.g. `C:\Python39\python.exe`)                                                                   |
| Arguments         | see below                                                                                                                     |
| WorkingDirectory  | `R:\Finance\Revenue Audit\HOTEL\City Ledger Report\Processes\`                                                                |

Adjust the script path, timeout, and move target to your environment. Quote paths containing spaces.

### `Arguments` field — common pitfall (error `BC30198`)

The `Arguments` property is a **VB.NET expression** that must evaluate to a **single `String`**. Entering three bare tokens separated by spaces:

```vbnet
"R:\...\wait_opera_print_pdf.py" 60 "R:\Reports\opera.pdf"   ' WRONG
```

is invalid VB (a string literal cannot be followed by a bare integer and another string) and produces:

```
error BC30198: ')' expected
```

### Correct form — wrap in one outer string and **double** inner quotes

Double every inner `"` so the paths-with-spaces stay quoted at runtime. The whole thing is one VB string literal:

```vbnet
"""R:\Finance\Revenue Audit\HOTEL\Opera Download_Transaction is DATE\Script\wait_opera_print_pdf.py"" 60 ""R:\Reports\opera.pdf"""
```

At runtime this evaluates to the single string:

```
"R:\Finance\Revenue Audit\HOTEL\Opera Download_Transaction is DATE\Script\wait_opera_print_pdf.py" 60 "R:\Reports\opera.pdf"
```

which the launcher splits into the three argv tokens the script expects:

| argv | value                                            |
|------|--------------------------------------------------|
| `1`  | `R:\...\wait_opera_print_pdf.py` (script path)   |
| `2`  | `60` (timeout in seconds)                        |
| `3`  | `R:\Reports\opera.pdf` (move target)             |

### Alternative — `IEnumerable(Of String)`

If the activity accepts an argument collection instead of a single string, use:

```vbnet
New String() {"R:\...\wait_opera_print_pdf.py", "60", "R:\Reports\opera.pdf"}
```

This avoids the doubled-quote encoding entirely.

## Dependencies

Standard library only — no third-party packages required:

- `os`, `sys`, `time`, `glob`, `shutil`