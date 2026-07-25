# ============================================================
# PSS/E Short Circuit Reporter (Summary Only)
# Batch IEC 60909 fault analysis to a formatted Excel summary
#
# Version : 1.0.0
# Author  : Haider Ali (github.com/4haiderali)
# License : MIT
#
# Usage:
#   1. Place this script in a folder containing (or above) your
#      PSS/E .sav case file(s). It will recursively search for
#      .sav files in the current folder and subfolders.
#   2. Edit the USER SETTINGS section below:
#        SC_BUSES  - bus number(s) to fault
#        PSSE_PATH - path to your PSS/E 34 Python 2.7 installation
#   3. Run from the PSS/E Python 2.7 environment:
#        exec(open("sc_reporter_summary.py").read())
#      or launch via the PSS/E script runner.
#   4. Results are written per-case to the output folder defined
#      by OUTPUT_FOLDER_NAME, as one Excel file per .sav case.
#
# Requirements:
#   PSS/E v34, Python 2.7, openpyxl
#   Install openpyxl in the PSS/E Python 2.7 environment:
#     <PSSE_PATH>\python.exe -m pip install openpyxl
#
# Changelog:
#   1.0.0  Initial release. Recursive .sav discovery, IEC 60909 short
#          circuit analysis on user-defined buses, formatted Excel summary
#          output (bus, name, kV, 3-phase/1-phase kA), a metadata footer,
#          and safe-save logic for a locked output file.
# ============================================================

from __future__ import print_function, division  # MUST be the first import
import sys
import os
import re
import time
from datetime import datetime

# --- OpenPyXL import (compatible across versions) ---
try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
    try:
        from openpyxl.utils import get_column_letter
    except ImportError:
        from openpyxl.cell import get_column_letter
except ImportError:
    print("ERROR: openpyxl module not found. Install it with:")
    print("  <PSSE_PATH>\\python.exe -m pip install openpyxl")
    sys.exit()

__version__ = "1.0.0"
TOOL_NAME = "PSS/E Short Circuit Reporter (Summary Only)"
TOOL_AUTHOR = "Haider Ali"
TOOL_GITHUB = "github.com/4haiderali"

# ============================================================
# PSS/E v34 / Python 2.7 setup
# ============================================================

# Adjust this to match your PSS/E 34 installation.
# Common alternatives:
#   r"C:\Program Files\PTI\PSSE34\PSSPY27"       (64-bit installer)
#   r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27" (32-bit / default)
PSSE_PATH = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"

# Try common PSS/E 34 locations if the configured path is not found.
if not os.path.isdir(PSSE_PATH):
    _common_psse_paths = [
        r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27",
        r"C:\Program Files\PTI\PSSE34\PSSPY27",
    ]
    _found = None
    for _p in _common_psse_paths:
        if os.path.isdir(_p):
            _found = _p
            break
    if _found is not None:
        PSSE_PATH = _found
    else:
        raise RuntimeError(
            "PSS/E Python path not found: {}\n"
            "Edit PSSE_PATH in the USER SETTINGS/header section to match "
            "your PSS/E 34 installation.".format(PSSE_PATH)
        )

if PSSE_PATH not in sys.path:
    sys.path.append(PSSE_PATH)
os.environ["PATH"] = PSSE_PATH + ";" + os.environ.get("PATH", "")

try:
    import psse34
except ImportError:
    pass

import psspy
import redirect

redirect.psse2py()
psspy.psseinit(50000)

# ============================================================
# USER SETTINGS
# ============================================================

# REQUIRED: set the bus number(s) to fault before running.
# The script will stop with a clear error if this is left empty.
# Example: SC_BUSES = [101, 205, 30012]
SC_BUSES = []

# Folder search depth (0 = current folder, 1 = one level down, etc.)
SEARCH_LEVELS = [0, 1, 2, 3]

# Output folder name (created alongside this script).
OUTPUT_FOLDER_NAME = "Short Circuit Reports"

# Fixed-width bus name field in the PSS/E SC report (v34 caps names at 12).
BUS_NAME_FIELD_WIDTH = 12

# Sanity ceiling for parsed kV; raise if your model uses higher voltages.
MAX_PLAUSIBLE_KV = 800.0

BORDER_COLOR = "2F75B5"               # Strong Blue
HEADER_INSIDE_BORDER_COLOR = "EDF5FC" # Very Light Blue
HEADER_FILL_COLOR = "4A85B8"          # Standard Blue
LIGHT_FILL_COLOR = "DDEBF7"           # Light Blue
TEXT_COLOR = "FFFFFF"                 # White

if not SC_BUSES:
    raise RuntimeError(
        "SC_BUSES is empty. Please define SC_BUSES = [bus numbers to fault] "
        "in the USER SETTINGS section before running."
    )

# ============================================================
# System initialization
# ============================================================

mydir = os.getcwd()
os.chdir(mydir)
print("Working directory: {}".format(mydir))

output_folder = os.path.join(mydir, OUTPUT_FOLDER_NAME)
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# ============================================================
# Helper functions
# ============================================================

def find_sav_files(root_dir, allowed_levels):
    savs = []
    for root, dirs, files in os.walk(root_dir):
        rel_path = os.path.relpath(root, root_dir)
        depth = 0 if rel_path == "." else rel_path.count(os.sep) + 1
        if depth in allowed_levels:
            for f in files:
                if f.lower().endswith(".sav"):
                    savs.append(os.path.join(root, f))
    return savs


def filter_existing_buses(bus_list):
    """Return only buses that exist in the currently loaded case."""
    existing = []
    for b in bus_list:
        try:
            ierr, _ = psspy.busdat(int(b), 'BASE')
            if ierr == 0:
                existing.append(int(b))
        except Exception:
            pass
    return existing


def run_iec_sc_analysis(sav_path, report_path, log_path):
    """Run IEC 60909 short circuit analysis on SC_BUSES for one case."""
    ierr = psspy.case(sav_path)
    if ierr != 0:
        print("Error loading {}".format(sav_path))
        return False

    alert_path = os.path.splitext(report_path)[0] + "_Alert.txt"
    prompt_path = os.path.splitext(report_path)[0] + "_Prompt.txt"

    psspy.report_output(2, report_path, [0, 0])
    psspy.progress_output(2, log_path, [0, 0])
    psspy.alert_output(2, alert_path, [0, 0])
    psspy.prompt_output(2, prompt_path, [0, 0])

    try:
        psspy.short_circuit_units(1)
        psspy.short_circuit_z_units(1)
        psspy.short_circuit_coordinates(1)
        psspy.short_circuit_z_coordinates(1)

        valid_buses = filter_existing_buses(SC_BUSES)
        if not valid_buses:
            print("  - No valid SC buses found in this case. Skipping.")
            return False

        psspy.bsys(0, 0, [0.0, 0.0], 0, [], len(valid_buses), valid_buses, 0, [], 0, [])

        psspy.iecs_4(0, 0,
                     [1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 1, 0],
                     [0.1, 1.0],
                     "", "", "")
    except Exception as e:
        print("Error during simulation: {}".format(str(e)))
        return False
    finally:
        psspy.report_output(1, "", [0, 0])
        psspy.progress_output(1, "", [0, 0])
        psspy.alert_output(1, "", [0, 0])
        psspy.prompt_output(1, "", [0, 0])

    return True


def format_excel_sheet(ws, case_title):
    """Apply header/title styling, borders, widths, and row heights."""
    title_font = Font(bold=True, color=TEXT_COLOR, size=16)
    header_font = Font(bold=True, color=TEXT_COLOR, size=14)
    data_font = Font(bold=False, color="000000", size=12)

    header_fill = PatternFill(start_color=HEADER_FILL_COLOR, end_color=HEADER_FILL_COLOR, fill_type="solid")
    light_fill = PatternFill(start_color=LIGHT_FILL_COLOR, end_color=LIGHT_FILL_COLOR, fill_type="solid")

    center_style = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin_border = Border(left=Side(style='thin', color=BORDER_COLOR),
                          right=Side(style='thin', color=BORDER_COLOR),
                          top=Side(style='thin', color=BORDER_COLOR),
                          bottom=Side(style='thin', color=BORDER_COLOR))

    light_inside = Side(style="thin", color=HEADER_INSIDE_BORDER_COLOR)

    line1, line2 = psspy.titldt()
    ws['A1'] = "Short Circuit Fault Currents - " + line1.title()
    ws.merge_cells('A1:F1')

    ws['A2'] = "S. No."
    ws['B2'] = "Bus Number"
    ws['C2'] = "Bus Name"
    ws['D2'] = "Base kV"
    ws['E2'] = "3-Phase"
    ws['F2'] = "1-Phase"
    ws['E3'] = "(kA)"
    ws['F3'] = "(kA)"

    ws.merge_cells('A2:A3')
    ws.merge_cells('B2:B3')
    ws.merge_cells('C2:C3')
    ws.merge_cells('D2:D3')

    identifiers = ["", "(A)", "(B)", "(C)", "(D)", "(E)"]
    for i, char in enumerate(identifiers, 1):
        ws.cell(row=4, column=i).value = char

    max_row = ws.max_row
    max_col = 6

    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.border = thin_border
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = title_font
                cell.alignment = center_style
            elif cell.row in (2, 3):
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_style
            elif cell.row == 4:
                cell.fill = light_fill
                cell.font = data_font
                cell.alignment = center_style
            else:
                cell.font = data_font
                if cell.col_idx == 1:
                    cell.fill = light_fill
                    cell.alignment = center_style
                elif cell.col_idx == 3:
                    cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
                else:
                    cell.alignment = center_style

    header_max_row = 3
    for r in ws.iter_rows(min_row=1, max_row=header_max_row, max_col=max_col):
        for cell in r:
            b = cell.border
            r_idx, c_idx = cell.row, cell.column
            left = light_inside if c_idx > 1 else b.left
            right = light_inside if c_idx < max_col else b.right
            top = light_inside if r_idx > 1 else b.top
            bottom = light_inside if r_idx < header_max_row else b.bottom
            if r_idx == 4:
                bottom = Side(style='thin', color=BORDER_COLOR)
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    widths = {'A': 10, 'B': 17, 'C': 25, 'D': 15, 'E': 17, 'F': 17}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 25

    ws.freeze_panes = "A5"


def add_metadata_footer(ws, case_title, data_end_row):
    """Write tool name, version, author, GitHub, case, and timestamp below the table."""
    label_font = Font(bold=True, color="808080", size=9, italic=True)
    value_font = Font(bold=False, color="808080", size=9, italic=True)

    start_row = data_end_row + 2  # one blank row gap after the table

    footer_rows = [
        ("Generated by", "{} v{}".format(TOOL_NAME, __version__)),
        ("Author", TOOL_AUTHOR),
        ("GitHub", TOOL_GITHUB),
        ("Case", case_title),
        ("Generated on", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]

    for i, (label, value) in enumerate(footer_rows):
        r = start_row + i
        c_label = ws.cell(row=r, column=1, value=label)
        c_label.font = label_font
        c_value = ws.cell(row=r, column=2, value=value)
        c_value.font = value_font
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)


def parse_sc_report_to_excel(report_path, excel_path, case_title):
    print("Formatting Excel: {}".format(excel_path))

    if not os.path.exists(report_path):
        return

    with open(report_path, 'r') as f:
        lines = f.readlines()

    pattern = re.compile(
        r"^\s*(\d+)\s+"      # Bus Number
        r"\[(.*?)\]\s+"      # Name and kV
        r"(\w+)\s+"          # Unit
        r"([-\d\.]+)\s+"     # 3Ph I
        r"([-\d\.]+)\s+"     # 3Ph Ang
        r"([-\d\.]+)\s+"     # LG I
        r"([-\d\.]+)\s+"     # LG Ang
        r"([-\d\.]+)"        # Voltage factor
    )

    data_rows = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        match = pattern.search(line)
        if not match:
            continue

        bus_num = int(match.group(1))
        name_kv_raw = match.group(2)  # unstripped - slice depends on position

        name_part = name_kv_raw[:BUS_NAME_FIELD_WIDTH]
        kv_part = name_kv_raw[BUS_NAME_FIELD_WIDTH:]

        bus_name = name_part.strip()
        kv_str = kv_part.strip()

        try:
            base_kv_val = float(kv_str)
            if base_kv_val <= 0 or base_kv_val > MAX_PLAUSIBLE_KV:
                raise ValueError("implausible kV: {}".format(base_kv_val))
        except ValueError:
            base_kv_val = ""
            bus_name = name_kv_raw.strip()
            print("  - WARNING: could not parse kV for bus {} from '{}'; "
                  "check BUS_NAME_FIELD_WIDTH.".format(bus_num, name_kv_raw))

        i_3ph_ka = float(match.group(4)) / 1000.0
        i_lg_ka = float(match.group(6)) / 1000.0

        data_rows.append({
            'num': bus_num, 'name': bus_name, 'kv': base_kv_val,
            '3ph': i_3ph_ka, '1ph': i_lg_ka
        })

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Short Circuit Summary"

    current_row = 5
    for idx, row_data in enumerate(data_rows, 1):
        ws.cell(row=current_row, column=1, value=idx)
        ws.cell(row=current_row, column=2, value=row_data['num'])
        ws.cell(row=current_row, column=3, value=row_data['name'])
        ws.cell(row=current_row, column=4, value=row_data['kv'])

        c_3ph = ws.cell(row=current_row, column=5, value=row_data['3ph'])
        if row_data['3ph'] > 0 and round(row_data['3ph'], 2) != 0:
            c_3ph.number_format = '0.00'

        c_1ph = ws.cell(row=current_row, column=6, value=row_data['1ph'])
        if row_data['1ph'] > 0 and round(row_data['1ph'], 2) != 0:
            c_1ph.number_format = '0.00'

        current_row += 1

    last_data_row = current_row - 1
    format_excel_sheet(ws, case_title)
    add_metadata_footer(ws, case_title, last_data_row)

    # --- Safe save: auto-close the target file if it's open in Excel ---
    try:
        import win32com.client
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            xl = None

        abs_target = os.path.abspath(excel_path).lower()

        if xl is not None:
            try:
                open_books = [wb_open for wb_open in xl.Workbooks]
            except Exception:
                open_books = []
            for wb_open in open_books:
                try:
                    if wb_open.FullName and wb_open.FullName.lower() == abs_target:
                        wb_open.Close(SaveChanges=False)
                        break
                except Exception:
                    pass

        for _ in range(10):
            if not os.path.exists(excel_path):
                break
            try:
                os.remove(excel_path)
                break
            except Exception:
                time.sleep(0.2)
    except Exception:
        pass

    wb.active = wb.index(ws)
    wb.save(excel_path)
    print("Saved Excel: {}".format(excel_path))


# ============================================================
# Main execution
# ============================================================

def main():
    print("{} v{} - {}".format(TOOL_NAME, __version__, TOOL_GITHUB))
    print("Starting Short Circuit Batch Processing (Python 2.7)...")

    sav_files = find_sav_files(mydir, SEARCH_LEVELS)
    if not sav_files:
        print("No .sav files found.")
        return

    print("Found {} .sav file(s).".format(len(sav_files)))

    text_output_folder = os.path.join(output_folder, "Text Logs")
    if not os.path.exists(text_output_folder):
        os.makedirs(text_output_folder)

    for i, sav_path in enumerate(sav_files):
        print("\nProcessing: {}".format(os.path.basename(sav_path)))

        case_name = os.path.splitext(os.path.basename(sav_path))[0]
        report_file = os.path.join(text_output_folder, "{}-{}_SC.txt".format(i, case_name))
        log_file = os.path.join(text_output_folder, "{}-{}_Log.txt".format(i, case_name))
        excel_file = os.path.join(output_folder, "{}-{}_SC.xlsx".format(i, case_name))

        for f in [report_file, log_file]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass

        success = run_iec_sc_analysis(sav_path, report_file, log_file)
        if success:
            print("  - Report saved.")
            parse_sc_report_to_excel(report_file, excel_file, case_name)

    print("\nProcessing Complete.")


if __name__ == "__main__":
    main()
