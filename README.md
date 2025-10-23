## Process List Extractor (SysGauge → TXT) — PySide6 GUI

A small desktop GUI that parses SysGauge Process List reports from Excel/CSV and exports clean, tab‑delimited `.txt` files.

### What it does
- **Supported inputs**: `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.csx`
- **Auto‑detects** the Process List table (row containing "Process Name") across sheets/files
- **Cleans headers and rows**, removes placeholders and empty tails
- **Reorders and formats** columns to this exact schema:
  - `Process Name`, `Instances`, `CPU`, `Memory`, `Threads`, `Handles`, `Data`, `Status`
  - `CPU`/`Memory`/`Data` → 2 decimals (e.g., `658.21`)
  - `Threads`/`Handles` → thousands separators (e.g., `34,504`)
  - `Instances` → integer
- **Multi‑file TXT export**: pick many input files → write one `.txt` per file (filenames sanitized; duplicate names are safely de‑duplicated)

### Requirements
- **Python 3.10+**

### Install
```bash
pip install PySide6 pandas openpyxl
```

### Run (GUI)
```bash
python Untitled-1.py
```
- Click **Open Excel/CSV (multiple)**, preview the normalized table
- Click **Save .txt files (choose names)** to export one TXT per selected input

### Run tests (optional)
```bash
python Untitled-1.py --run-tests
```
Prints "All tests passed." on success. Tests run in‑memory and don't require sample files on disk.

### Notes
- This tool is **GUI‑only**; no separate CLI conversion mode is provided.
- If you see a message about PySide6 not being installed, install deps with `pip install PySide6` (see Install above).
- For CSVs using uncommon delimiters, the app attempts to auto‑detect via pandas’ python engine; `.csx` is treated like CSV.

### Export format
- Output files are UTF‑8, tab‑delimited, with Unix newlines (`\n`).
- Header row is always:

```text
Process Name	Instances	CPU	Memory	Threads	Handles	Data	Status
```

### File naming
- Proposed names are derived from the input file names and sanitized so they are safe on Windows/macOS/Linux.
- If a chosen destination name would overwrite a file selected earlier in the same session, a numeric suffix like `_2`, `_3`, … is added.

### Troubleshooting
- On Linux, ensure a working X/Wayland desktop is available for PySide6/Qt to create windows.
- If you encounter missing Qt platform plugins, reinstall `PySide6` in a clean virtual environment.
