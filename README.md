# Change Impact Analyzer

A local desktop prototype for instrumentation and control engineering revision review.

Compare previous and new Excel or text-based PDF revisions, identify added, removed, and changed tags, and generate rule-based prompts for potentially affected engineering documents.

## Features

- Excel comparison by tag and sheet/column field.
- Text-PDF comparison based on tag occurrence and nearby text.
- Rule-based impact prompts for signal, range, setpoint, fail action, location, type, power, cable, and alarm changes.
- Search, status filters, and per-result review details.
- CSV report export.
- Lavender desktop interface; processing remains local.

## Run

Python 3.10+ with Tk support is required. Designed for Windows.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe tools\create_samples.py
.\.venv\Scripts\python.exe app.py
```

Select `samples/previous.xlsx` and `samples/revised.xlsx`, then click **Analyze Changes**. The synthetic example contains a changed range, a removed tag, and an added tag. `previous.pdf` and `revised.pdf` demonstrate text-PDF context comparison separately.

## Build a Windows executable

Run `build_windows_exe.bat`. It installs build dependencies and creates `dist/ChangeImpactAnalyzer.exe` using PyInstaller. The generated executable and bundled libraries are excluded from source control. The desktop application itself does not require PyInstaller to run from source.

## Input assumptions

- Excel sheets need a recognizable tag column header within their first 25 rows, such as `Tag`, `Tag Number`, or `Instrument Tag`.
- Fields are keyed by sheet name and column header. Renaming either can produce apparent field changes.
- `.xlsx`, `.xlsm`, and `.xls` are accepted. Legacy `.xls` support requires `xlrd`.
- Text PDFs must contain selectable text. Scanned PDFs are not supported.
- PDF results concern tag occurrence and nearby text, not verified structured fields, drawing geometry, or logic changes.
- Cross-format comparisons do not share a normalized semantic schema; compare like-for-like files for meaningful results.

## Meaning of impact results

This version uses keyword rules, not AI or an imported relationship graph. For example, a range-field change suggests review of Data Sheets, DCS Scaling, Alarm/Interlock, and FAT Procedures. These are potential review targets, not proven dependencies or confirmed design errors. Validate the results against controlled engineering documents.

## Verification

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Tests cover actual Excel revision loading and rule-based comparisons, text-PDF changes, scanned/blank-PDF warnings, and unmapped-field fallback. Desktop interactions, executable packaging, and legacy `.xls` input are not tested here.

## Repository contents

Application source, documentation, build instructions, tests, and a synthetic example generator. No original project workbooks, PDFs, runtime bundles, credentials, or trained models are included. No redistribution license is granted with this repository.
