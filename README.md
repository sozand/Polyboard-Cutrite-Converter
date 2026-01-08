# Polyboard Infra PP

Tools to prepare Polyboard project exports for production: clean and convert MPR files, merge cutlist data, and produce final Cutrite-compatible output.

## Features
- GUI cutlist generator (tkinter) with dark theme.
- Convention editor (JSON sidecar) with import/export to Excel.
- Deterministic `Unique_ID` per row (uuid5 over Project, Cabinet, Reference, Cutting_List_Number).
- Drill details and groove lengths:
  - Vertical/Horizontal drill signature columns.
  - Groove lengths for macros 109/124 with direction and workpiece size.
- Edge handling:
  - Edge code mapping via convention.
  - Edge band count column.
- MPR preprocessing on export:
  - Remove component block `<139 \Komponente\ ...>`.
  - Strip all macro 124.
  - Convert macro 109 with `T_` ending `xxxxx2` to macro 151 (milling from below) with adjustable tool diameter.
  - Confirmation dialog with per-file actions (component removed, 124 removed, 109→151, LA/BR).
  - Backups (`.bak`) before modifications.

## Requirements
- Python 3.10+ (tkinter, pandas, openpyxl).
- Git (optional, for version control).

## Setup
1) Install dependencies: pip install pandas openpyxl
2) Ensure `Polyboard_convention.json` is alongside the GUI script (auto-created if missing).

## Usage
1) Run the GUI: python polyboard_production_gui.py
2) Load & Preview:
   - Select project cutlist CSV (semicolon-delimited, no headers).
   - Convention JSON loads automatically; edit via “Edit Convention.”
   - Preview shows mapped edge codes, drill/groove details, edge band count, Unique_ID.
3) Export Final Cutlist:
   - Shows interactive MPR change summary; confirm to apply.
   - MPR files are backed up (`.bak`) before modifications.
   - Outputs final cutlist CSV (semicolon-separated).

## Build Windows .exe (PyInstaller one-folder)
1) Requirements: Python 3.12 x64, `pip install -r requirements.txt`.
2) Run `build_exe.bat` from the repo root. It creates `dist\PolyboardProduction\PolyboardProduction.exe`.
3) Bundled assets: convention JSON/XLSX, `mpr_format_reference.json`, `Polyboard_Convention_Column_Summary.md`, `woodwop-mpr4x-format-pdf-free.pdf`, and the `Edge_Diagram_Ref` image folder.
4) Verify after build: launch the EXE, open “Edit Convention” (JSON loads/edits), load a sample cutlist CSV, and confirm edge diagram thumbnails render. If Pillow is absent, the GUI will list image names instead of thumbnails.

## Configurable paths (portable defaults)
- A `polyboard_config.json` file lives next to the script/EXE to remember user-selected defaults:
  - `convention_json`: full path to the convention JSON.
  - `edge_dir`: full path to the `Edge_Diagram_Ref` folder.
- In the GUI, use the Browse buttons (Convention JSON, Edge Diagram Folder) and click “Save defaults” to persist. The app falls back to paths beside the script if the config is missing.

## Configuration
- Tool diameter (for 109→151) via spinbox in GUI (default 10mm).
- Convention stored in `Polyboard_convention.json`; import/export via dialog.
- Line endings: repo can use LF to avoid CRLF churn.

## Notes
- Macro 109 conversion only triggers when `T_` ends with `xxxxx2`.
- Groove length formats: `dx_On_PL<LA_...>` or `dy_On_PW<BR_...>` with milling/top suffix for 109; no suffix for 124.
- Drill signature detail columns use `@` separators.

## License
Specify your license here.
