# EESISP Remittance Calculator

Weekly union remittance calculator built for **American Power Electrical Supply Co.**, a Philadelphia-area electrical distributor serving the NY metro market.

Replaced a manual Excel workflow that required editing formulas every time an employee was added, removed, or changed roles. Available as a standalone Windows desktop application (Python/Tkinter) and a browser-based version (vanilla HTML/JS).

---

## Features

- **Employee management** — add, rename, or remove employees; switch classifications; toggle active/inactive without losing the record
- **Four union classifications** — Sales, Drivers, Warehouse, Admin — each with its own configurable premium rate
- **JIB calculation** — computed automatically as a percentage of total gross wages, all classifications
- **401K C&L** — manual entry; combined with JIB for the remittance total
- **Rate change confirmation** — any edit to a premium percentage or the JIB rate requires explicit confirmation before applying
- **Excel paste import** — copy rows directly out of a payroll spreadsheet; fuzzy name matching handles reversed names, casing differences, and middle initials
- **Auto-sort** — employee list stays alphabetical; new hires slot into place when you finish typing the name
- **CSV export** — full breakdown by employee and classification
- **Printable summary** — opens a clean HTML report in the browser ready to print
- **Backup / Restore** — JSON export and import for moving data between machines
- **Auto-save** — data persists locally without any manual save step

---

## Files

| File | Description |
|---|---|
| `eesisp_calculator.py` | Desktop application — Python 3.8+, standard library only (tkinter) |
| `eesisp-calculator.html` | Browser version — single HTML file, no dependencies, localStorage persistence |
| `build_exe.bat` | Builds a standalone `EESISP_Calculator.exe` using PyInstaller |

---

## Running the Desktop App

Requires Python 3.8 or later with tkinter (included in the standard python.org installer).

```bat
python eesisp_calculator.py
```

Data is saved automatically to `%APPDATA%\EESISP Calculator\eesisp_data.json` — separate from the executable, so rebuilding never wipes your records.

---

## Building the .exe

Put `eesisp_calculator.py` and `build_exe.bat` in the same folder and double-click the batch file. It installs PyInstaller if needed and produces a single portable executable at `dist\EESISP_Calculator.exe`.

```bat
build_exe.bat
```

---

## Running the Browser Version

Open `eesisp-calculator.html` directly in Chrome, Edge, or Firefox — no server required. Data saves to `localStorage` in the browser automatically.

---

## How the Calculation Works

| Item | Formula |
|---|---|
| TGW (taxable gross wages) | Gross wages − Deductions |
| Classification premium | TGW × rate % |
| JIB | Total gross wages (all classes) × JIB % |
| **Remittance total** | **JIB + 401K C&L** |

---

## Tech

- Python 3 / tkinter — no third-party packages required to run
- PyInstaller — for building the Windows executable
- Vanilla HTML, CSS, JavaScript — no frameworks or build tools
- localStorage for browser persistence; JSON file for desktop persistence
