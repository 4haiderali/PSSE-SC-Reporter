# PSS/E Short Circuit Reporter (Summary Only)

**Batch IEC 60909 short circuit analysis across multiple PSS/E cases, exported to a formatted Excel summary.**

This tool automates running an IEC 60909 short circuit study on a user-defined set of buses across every `.sav` case in a folder (and its subfolders), and produces a clean, presentation-ready Excel summary of 3-phase and 1-phase fault currents per bus — no manual report reading or copy-pasting.

Buses that don't exist in a given case are automatically skipped, so the same bus list can be reused across multiple case variants without editing.

---

## Repository Contents

| File | Purpose |
|---|---|
| `sc_reporter_summary.py` | Main Python automation script for PSS/E. |
| `README.md` | Tool documentation and usage guide. |
| `LICENSE` | MIT license. |

---

## Features

- Recursive `.sav` file discovery across a project folder structure
- IEC 60909 short circuit analysis on user-defined buses
- Per-case bus filtering — skips buses that don't exist in a given case
- Robust bus name / base kV parsing using PSS/E's fixed-width name field, rather than fragile whitespace splitting
- Formatted Excel summary output — bus number, name, base kV, 3-phase and 1-phase fault currents (kA)
- Tool/version/author metadata footer on every output sheet for traceability
- Safe-save logic that auto-closes a locked output file before overwriting

---

## Requirements

| Requirement | Version |
|---|---|
| PSS/E | v34 |
| Python | 2.7, PSS/E environment |
| openpyxl | Any recent version compatible with Python 2.7 |

Install `openpyxl` in the PSS/E Python 2.7 environment:

```
<PSSE_PATH>\python.exe -m pip install openpyxl
```

Example PSS/E Python paths:

```
C:\Program Files (x86)\PTI\PSSE34\PSSPY27
C:\Program Files\PTI\PSSE34\PSSPY27
```

---

## Usage

1. Place `sc_reporter_summary.py` in a folder containing (or above) your PSS/E `.sav` case file(s).

2. Open the script and edit the **USER SETTINGS** section.

   At minimum, define the bus(es) to fault:

   ```
   SC_BUSES = [101, 205, 30012]
   ```

3. Confirm the PSS/E installation path:

   ```
   PSSE_PATH = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
   ```

4. Run from the PSS/E Python 2.7 environment:

   ```
   exec(open("sc_reporter_summary.py").read())
   ```
   Or run it using the PSS/E script runner.

5. Results are written to a `Short Circuit Reports` folder created alongside the script — one Excel file per `.sav` case, plus raw text logs.

---

## Important Settings

| Setting | Default | Description |
|---|---|---|
| `SC_BUSES` | `[]` | Required. One or more bus numbers to fault. |
| `SEARCH_LEVELS` | `[0, 1, 2, 3]` | Folder search depth for `.sav` discovery. |
| `OUTPUT_FOLDER_NAME` | `"Short Circuit Reports"` | Output folder name. |
| `BUS_NAME_FIELD_WIDTH` | `12` | Fixed-width bus name field used by PSS/E v34's SC report. |
| `MAX_PLAUSIBLE_KV` | `800.0` | Sanity ceiling for a parsed kV value; raise for higher-voltage models. |

---

## Output

Each `.sav` case produces one Excel file containing:

- **Short Circuit Summary** — bus number, bus name, base kV, 3-phase fault current (kA), 1-phase fault current (kA)
- A metadata footer below the table — tool name, version, author, GitHub, case name, and run timestamp

---

## Limitations

- Tested for PSS/E v34 and Python 2.7.
- Not yet tested on other PSS/E versions.
- Assumes IEC 60909 is the desired fault standard; ANSI/other standards are not covered by this script.
- The script does not modify the PSS/E case. It only reads data and writes an Excel report.

---

## Recommended Public-Repo Practice

Do not upload confidential project files, `.sav` cases, client-specific outputs, or study results. The repository should contain only:

- Generic script
- README
- LICENSE
- Optional screenshots or synthetic examples

---

## License

MIT — see `LICENSE`.

---

## Author

**Haider Ali**
Senior Power System Engineer
Power Planners International, Lahore, Pakistan

GitHub: [github.com/4haiderali](https://github.com/4haiderali)
LinkedIn: [linkedin.com/in/4haiderali](https://linkedin.com/in/4haiderali)

---

## Version

Current release: `1.0.0`
