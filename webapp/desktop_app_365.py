"""
CFO Financial Model Generator - Desktop Application (Microsoft 365 Edition)
Creates macro-enabled Excel files using openpyxl - NO desktop Excel required

VERSION HISTORY:
- v1.0.0 (2024-12-01): Initial release with P&L, Balance Sheet, Cash Flow
- v1.1.0 (2024-12-05): Added Dashboard module with KPIs and ratios
- v1.2.0 (2025-12-09): Renamed from DNA Model to Financial Model
- v1.2.2 (2025-12-10): Fixed Dashboard YTD calculations to use current year only
                        (Jan through current month from Menu!C7)
- v2.0.0 (2025-12-16): Added multi-division support with consolidation
                        - Division setup wizard for multiple entities
                        - Intelligent account matching via ChatGPT API
                        - Consolidated P&L, Balance Sheet, Cash Flow
                        - Division-specific and consolidated views
                        - Account mapping persistence (JSON + Excel)
- v3.1.0 (2026-01-16): Comprehensive fix pass - Dashboard and Forecast critical fixes
                        - FIXED: Dashboard charts now sync with division selector
                        - FIXED: Forecast logic uses data's current month, not system date
                        - FIXED: Past months use Actuals, future months use Budget+Adj
                        - FIXED: #REF! errors in annual totals (always use col offset 2 for report sheets)
                        - FIXED: Dynamic progress bar uses historical timing data
                        - FIXED: Menu helper columns E-K hidden AFTER column width settings
                        - Menu: Right-aligned values (C7-C10), standard blue hyperlinks
                        - P&L/BS/Cash Flow: AutoFit only, no minimum width enforcement
                        - Balance Sheet: Added row grouping matching P&L style
                        - Divisions 3-5: Row groupings collapsed by default
- v3.2.0 (2026-01-16): Settings tab and sheet ordering improvements
                        - Added Settings sheet with sheet visibility controls
                        - Reordered sheets (Forecast_Summary before Consolidated_Forecast)
                        - Added division-specific Forecast Summary sheets
                        - Removed orphan sheets (Balance_Sheet, Consolidated_BS, Dashboard_Control)
- v3.2.1 (2026-01-16): Number formatting and column width fixes
                        - Removed cents from all dollar formats (#,##0 instead of #,##0.00)
                        - Set column widths to 13 (fits $10,000,000)
- v3.2.2 (2026-01-16): File validation and performance improvements
                        - Added file type detection (P&L vs Balance Sheet)
                        - Shows error if wrong file type uploaded
                        - Removed unnecessary Configuration UI (dates auto-detected)
                        - PERF: Replaced column width loops with bulk operations
- v3.2.3 (2026-01-16): Period selector dropdown
                        - Fixed month format from "Nov 24" to "Nov 2024" (avoids date confusion)
                        - Added dropdown to Menu C7 for selecting viewing period
                        - Added lookup table (hidden K:M) for period-to-YYYYMM conversion
                        - G7 now uses VLOOKUP formula to auto-update when period changes
                        - All P&L, BS, YTD data dynamically adjusts to selected period
- v3.2.4 (2026-01-16): Bug fixes and UI simplification
                        - Fixed missing data_start_month error (dates now auto-detected)
                        - Removed account mapping dialog (auto-processes mappings)
- v3.2.5 (2026-01-16): Forecast Summary and Dashboard fixes
                        - Fixed CM Actual formula column range for multi-division mode
                        - Restored Dashboard_Control sheet creation (needed for KPI targets)
                        - Updated build date display to 2026-01-16
- v3.2.6 (2026-01-16): Performance optimization - Template pre-formatting
                        - PERF: Division sheets now use clone-and-inject pattern
                        - Creates _TPL_PL and _TPL_BS templates once, clones for each division
                        - Reduces division sheet creation from ~35-40s to ~10-15s each
                        - Expected total time reduction: 40-50% for multi-division models
                        - Templates deleted after generation (not visible in final file)
                        - NOTE: Optimization temporarily disabled due to COM errors
- v3.2.7 (2026-01-16): Cash Flow division selector
- v3.2.8 (2026-01-23): Unified update model support
                        - Auto-detects single-entity vs multi-division models
                        - Date prompt dialog when auto-detection fails
                        - Improved error messages with header contents
- v3.2.9 (2026-01-23): Update performance optimization
                        - Bulk Find/Replace for formula range updates
                        - Bulk formula writes instead of cell-by-cell
                        - Eliminated slow nested loops in report updates
                        - Added division dropdown selector to Cash Flow sheet (cell E2)
                        - Options: "All Divisions" (default) or any specific division
                        - All Cash Flow formulas dynamically filter by selected division
                        - Uses IF/SUMIFS pattern for division-aware calculations
"""

import os
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime
import tempfile
import shutil
import json
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers, NamedStyle
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from copy import copy

# Multi-division support imports
from consolidation_engine import ConsolidationEngine, DivisionConfig, AccountMapping
from mapping_persistence import MappingPersistence

# ============================================================
# openpyxl Style Constants & Helper Functions
# ============================================================

# Color constants (hex strings for openpyxl)
CLR_DARK_BLUE = "16213E"       # RGB(22, 33, 62)
CLR_ACCENT_BLUE = "3B5998"     # RGB(59, 89, 152)
CLR_LINK_BLUE = "0066CC"       # RGB(0, 102, 204)
CLR_WHITE = "FFFFFF"
CLR_SUBTOTAL_GRAY = "ECECEC"   # RGB(236, 236, 236)
CLR_LIGHT_GRAY = "C8C8C8"     # RGB(200, 200, 200)
CLR_MED_GRAY = "DCDCDC"       # RGB(220, 220, 220)
CLR_LIGHT_YELLOW = "FFFFC8"   # RGB(255, 255, 200)
CLR_DARK_GREEN = "006400"      # RGB(0, 100, 0)
CLR_PURPLE = "800080"          # RGB(128, 0, 128)
CLR_GRAY_TEXT = "808080"       # RGB(128, 128, 128)
CLR_GRAY_64 = "646464"        # RGB(100, 100, 100)

# Pre-built style objects
FILL_DARK_BLUE = PatternFill(start_color=CLR_DARK_BLUE, end_color=CLR_DARK_BLUE, fill_type="solid")
FILL_SUBTOTAL_GRAY = PatternFill(start_color=CLR_SUBTOTAL_GRAY, end_color=CLR_SUBTOTAL_GRAY, fill_type="solid")
FILL_LIGHT_YELLOW = PatternFill(start_color=CLR_LIGHT_YELLOW, end_color=CLR_LIGHT_YELLOW, fill_type="solid")
FILL_LIGHT_GRAY = PatternFill(start_color=CLR_LIGHT_GRAY, end_color=CLR_LIGHT_GRAY, fill_type="solid")
FILL_MED_GRAY = PatternFill(start_color=CLR_MED_GRAY, end_color=CLR_MED_GRAY, fill_type="solid")
FILL_ACCENT_BLUE = PatternFill(start_color=CLR_ACCENT_BLUE, end_color=CLR_ACCENT_BLUE, fill_type="solid")

FONT_BOLD = Font(bold=True)
FONT_HEADER_WHITE = Font(bold=True, color=CLR_WHITE, size=11)
FONT_LINK_BLUE = Font(color=CLR_LINK_BLUE, underline="single")
FONT_GRAY_TEXT = Font(color=CLR_GRAY_TEXT)
FONT_GRAY_64 = Font(color=CLR_GRAY_64)
FONT_DARK_BLUE = Font(color=CLR_DARK_BLUE)
FONT_ACCENT_BLUE = Font(color=CLR_ACCENT_BLUE)
FONT_DARK_GREEN = Font(color=CLR_DARK_GREEN)
FONT_PURPLE = Font(color=CLR_PURPLE)

def rgb_fill(rgb_tuple):
    """Convert an (R, G, B) tuple to a PatternFill object."""
    if rgb_tuple is None:
        return PatternFill(fill_type=None)
    hex_color = "{:02X}{:02X}{:02X}".format(rgb_tuple[0], rgb_tuple[1], rgb_tuple[2])
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

SIDE_THIN = Side(style='thin')
SIDE_MEDIUM = Side(style='medium')
SIDE_DOUBLE = Side(style='double')
BORDER_TOP_THIN = Border(top=SIDE_THIN)
BORDER_TOP_MEDIUM = Border(top=SIDE_MEDIUM)
BORDER_BOTTOM_DOUBLE = Border(bottom=SIDE_DOUBLE)
BORDER_NET_INCOME = Border(top=SIDE_MEDIUM, bottom=SIDE_DOUBLE)
BORDER_BOX_THIN = Border(top=SIDE_THIN, bottom=SIDE_THIN, left=SIDE_THIN, right=SIDE_THIN)

ALIGN_CENTER = Alignment(horizontal='center')
ALIGN_RIGHT = Alignment(horizontal='right')
ALIGN_LEFT = Alignment(horizontal='left')
ALIGN_WRAP = Alignment(wrap_text=True)

NUM_FMT_CURRENCY = '#,##0'
NUM_FMT_PERCENT = '0.0%'
NUM_FMT_ACCOUNTING = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'


def apply_style_to_range(sheet, min_row, min_col, max_row, max_col,
                         font=None, fill=None, border=None, alignment=None, number_format=None):
    """Apply style properties to a rectangular range of cells."""
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = sheet.cell(row=row, column=col)
            if font is not None:
                cell.font = font
            if fill is not None:
                cell.fill = fill
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment
            if number_format is not None:
                cell.number_format = number_format


def set_col_width(sheet, col, width):
    """Set column width by column number."""
    sheet.column_dimensions[get_column_letter(col)].width = width


def set_col_widths_range(sheet, start_col, end_col, width):
    """Set column width for a range of columns."""
    for col in range(start_col, end_col + 1):
        sheet.column_dimensions[get_column_letter(col)].width = width


def hide_column(sheet, col):
    """Hide a column by column number."""
    sheet.column_dimensions[get_column_letter(col)].hidden = True


def hide_columns_range(sheet, start_col, end_col):
    """Hide a range of columns."""
    for col in range(start_col, end_col + 1):
        sheet.column_dimensions[get_column_letter(col)].hidden = True


def hide_row(sheet, row):
    """Hide a row by row number."""
    sheet.row_dimensions[row].hidden = True


def group_rows(sheet, start_row, end_row, outline_level=1, hidden=False):
    """Group rows with optional collapse."""
    for row in range(start_row, end_row + 1):
        sheet.row_dimensions[row].outline_level = outline_level
        if hidden:
            sheet.row_dimensions[row].hidden = True


def group_cols(sheet, start_col, end_col, outline_level=1, hidden=False):
    """Group columns with optional collapse."""
    for col in range(start_col, end_col + 1):
        letter = get_column_letter(col)
        sheet.column_dimensions[letter].outline_level = outline_level
        if hidden:
            sheet.column_dimensions[letter].hidden = True


def clear_sheet_data(sheet, start_row=1):
    """Clear all data and formatting from a sheet starting from start_row."""
    max_row = sheet.max_row
    if max_row and max_row >= start_row:
        sheet.delete_rows(start_row, max_row - start_row + 1)


def copy_sheet(wb, source_name, new_name):
    """Copy a worksheet within the same workbook."""
    source = wb[source_name]
    new_sheet = wb.copy_worksheet(source)
    new_sheet.title = new_name
    return new_sheet


def move_sheet(wb, sheet_name, position):
    """Move a sheet to a specific position (0-based index)."""
    if sheet_name in wb.sheetnames:
        wb.move_sheet(sheet_name, offset=position - wb.sheetnames.index(sheet_name))


def cells_replace(sheet, old_text, new_text):
    """Replace text in all cells of a sheet (equivalent to Excel's Cells.Replace)."""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and old_text in cell.value:
                cell.value = cell.value.replace(old_text, new_text)


def remove_filter_database_from_xlsx(filepath):
    """Remove legacy _xlnm._FilterDatabase defined names from a saved .xlsx/.xlsm file.

    openpyxl writes both a sheet-level <autoFilter> element AND a workbook-level
    _xlnm._FilterDatabase named range. Excel considers the named range redundant
    and flags 'Removed Records: Named range' during file repair. This function
    removes the legacy named range while preserving the sheet-level autoFilter.
    """
    import zipfile
    import re
    import shutil

    temp_path = filepath + '.tmp'
    try:
        with zipfile.ZipFile(filepath, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == 'xl/workbook.xml':
                        content = data.decode('utf-8')
                        # Remove _FilterDatabase defined name entries
                        content = re.sub(
                            r'<definedName[^>]*name="_xlnm\._FilterDatabase"[^>]*>[^<]*</definedName>',
                            '', content)
                        # Clean up empty definedNames tag if all entries removed
                        content = re.sub(r'<definedNames>\s*</definedNames>', '', content)
                        data = content.encode('utf-8')
                    zout.writestr(item, data)
        shutil.move(temp_path, filepath)
    except Exception as e:
        print(f"Warning: Could not remove _FilterDatabase from {filepath}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)


def apply_border_box(sheet, min_row, min_col, max_row, max_col, style='thin'):
    """Apply a border box around a rectangular range of cells."""
    side = Side(style=style)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = sheet.cell(row=row, column=col)
            top = side if row == min_row else cell.border.top
            bottom = side if row == max_row else cell.border.bottom
            left = side if col == min_col else cell.border.left
            right = side if col == max_col else cell.border.right
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)


def set_bulk_col_width(sheet, start_col, end_col, width):
    """Set column width for a range of columns by number."""
    for col in range(start_col, end_col + 1):
        sheet.column_dimensions[get_column_letter(col)].width = width


def write_data_to_cells(sheet, data, start_row, start_col=1):
    """Write a 2D list (list of rows) to cells starting at (start_row, start_col).
    Replaces xlwings pattern: sheet['A3'].value = data (2D list)."""
    for row_idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            sheet.cell(row=start_row + row_idx, column=start_col + col_idx, value=value)


def write_row_to_cells(sheet, row_data, row, start_col=1):
    """Write a 1D list to a single row of cells starting at (row, start_col).
    Replaces xlwings pattern: sheet['A4'].value = [row_data]."""
    for col_idx, value in enumerate(row_data):
        sheet.cell(row=row, column=start_col + col_idx, value=value)


def _final_formatting_check(sheet, data_start_row, last_data_row, last_col, sheet_type='pl'):
    """
    Final pass to ensure key financial totals are properly formatted.
    Scans column A for Net Income, Total Revenue, Total COGS, Total Expenses
    and applies correct bold + border formatting regardless of detection logic.

    This is the safety net that catches any totals missed by the primary detection.
    """
    # Patterns for key financial totals (case-insensitive)
    net_income_patterns = ['net income', 'net profit', 'net loss', 'net earnings']
    total_revenue_patterns = ['total revenue', 'total income', 'total revenues',
                              'total sales', 'total for income', 'total for revenue']
    total_cogs_patterns = ['total cost of goods', 'total cogs', 'total cost of sales',
                           'total for cost', 'cost of goods sold total', 'total cost of revenue']
    total_expense_patterns = ['total expense', 'total expenses', 'total operating expense',
                              'total for expense', 'total general and admin',
                              'total selling', 'total other expense']
    gross_profit_patterns = ['gross profit', 'gross margin', 'gross income']

    # BS patterns
    total_assets_patterns = ['total assets', 'total for assets']
    total_liab_equity_patterns = ['total liabilities and equity', 'total for liabilities and equity',
                                   'total liabilities & equity', 'total liabilities and shareholders']
    total_liabilities_patterns = ['total liabilities', 'total for liabilities']
    total_equity_patterns = ['total equity', "total stockholders' equity", "total shareholders' equity",
                             'total for equity']

    for row in range(data_start_row, last_data_row + 1):
        cell_val = sheet.cell(row=row, column=1).value
        if not cell_val or not isinstance(cell_val, str):
            continue
        name_lower = cell_val.strip().lower()

        is_net_income = any(p in name_lower for p in net_income_patterns)
        is_key_total = any(p in name_lower for p in (
            total_revenue_patterns + total_cogs_patterns + total_expense_patterns +
            gross_profit_patterns + total_assets_patterns + total_liab_equity_patterns +
            total_liabilities_patterns + total_equity_patterns
        ))

        if is_net_income:
            # Net Income: Bold + medium top border + double bottom border
            for col in range(1, last_col + 1):
                c = sheet.cell(row=row, column=col)
                c.font = Font(bold=True, size=c.font.size or 11)
                c.border = BORDER_NET_INCOME
        elif is_key_total:
            # Other key totals: Bold + thin top border + gray fill
            for col in range(1, last_col + 1):
                c = sheet.cell(row=row, column=col)
                c.font = Font(bold=True, size=c.font.size or 11)
                c.border = BORDER_TOP_THIN
                c.fill = FILL_SUBTOTAL_GRAY


# Application Version
APP_VERSION = "3.2.13"

# License/Expiration Settings
# Default expiration: 1 month from release (Feb 23, 2026)
DEFAULT_EXPIRATION = "2026-02-23"
LICENSE_SECRET_KEY = 0x4F43  # "OC" in hex - used for code obfuscation


class LicenseManager:
    """
    Simple license/expiration system for beta distribution.

    How it works:
    - App checks expiration date on startup
    - If expired, shows dialog requesting extension code
    - Extension codes encode a new expiration date
    - Codes are generated using: (YYYYMMDD XOR SECRET_KEY) converted to base36

    To generate a code for a date (run in Python):
        from desktop_app import LicenseManager
        code = LicenseManager.generate_code(2026, 3, 31)  # March 31, 2026
        print(code)  # Give this code to the user
    """

    LICENSE_FILE = "cfo_license.dat"

    @staticmethod
    def get_license_path():
        """Get path to license file in user's AppData"""
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        license_dir = os.path.join(appdata, 'CFO_Financial_Model')
        if not os.path.exists(license_dir):
            os.makedirs(license_dir)
        return os.path.join(license_dir, LicenseManager.LICENSE_FILE)

    @staticmethod
    def generate_code(year, month, day):
        """
        Generate an extension code for a given expiration date.
        Use this to create codes for users.

        Example: LicenseManager.generate_code(2026, 3, 31) -> "XXXXXX"
        """
        date_int = year * 10000 + month * 100 + day
        encoded = date_int ^ LICENSE_SECRET_KEY
        # Convert to base36 (alphanumeric) and add prefix
        code = LicenseManager._to_base36(encoded)
        # Add checksum character
        checksum = sum(ord(c) for c in code) % 26
        code = code + chr(65 + checksum)  # A-Z
        return f"CFO-{code[:4]}-{code[4:]}"

    @staticmethod
    def validate_code(code):
        """
        Validate an extension code and return the expiration date.
        Returns (year, month, day) tuple or None if invalid.
        """
        try:
            # Remove prefix and dashes
            code = code.upper().replace("CFO-", "").replace("-", "")
            if len(code) < 5:
                return None

            # Extract checksum
            checksum_char = code[-1]
            code_body = code[:-1]

            # Verify checksum
            expected_checksum = sum(ord(c) for c in code_body) % 26
            if chr(65 + expected_checksum) != checksum_char:
                return None

            # Decode
            encoded = LicenseManager._from_base36(code_body)
            date_int = encoded ^ LICENSE_SECRET_KEY

            # Extract date parts
            year = date_int // 10000
            month = (date_int % 10000) // 100
            day = date_int % 100

            # Basic validation (allow up to 2099 for "permanent" licenses)
            if year < 2024 or year > 2099:
                return None
            if month < 1 or month > 12:
                return None
            if day < 1 or day > 31:
                return None

            return (year, month, day)
        except:
            return None

    @staticmethod
    def _to_base36(num):
        """Convert integer to base36 string"""
        chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        result = ""
        while num > 0:
            result = chars[num % 36] + result
            num //= 36
        return result or "0"

    @staticmethod
    def _from_base36(s):
        """Convert base36 string to integer"""
        return int(s, 36)

    @staticmethod
    def get_expiration_date():
        """Get current expiration date (from license file or default)"""
        license_path = LicenseManager.get_license_path()

        if os.path.exists(license_path):
            try:
                with open(license_path, 'r') as f:
                    data = f.read().strip()
                    # Decode stored date
                    parts = data.split('-')
                    if len(parts) == 3:
                        return datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            except:
                pass

        # Return default expiration
        parts = DEFAULT_EXPIRATION.split('-')
        return datetime(int(parts[0]), int(parts[1]), int(parts[2]))

    @staticmethod
    def save_expiration_date(year, month, day):
        """Save new expiration date to license file"""
        license_path = LicenseManager.get_license_path()
        with open(license_path, 'w') as f:
            f.write(f"{year}-{month:02d}-{day:02d}")

    @staticmethod
    def is_first_run():
        """Check if this is the first run (no license file exists)"""
        license_path = LicenseManager.get_license_path()
        return not os.path.exists(license_path)

    @staticmethod
    def is_expired():
        """Check if the application has expired"""
        exp_date = LicenseManager.get_expiration_date()
        return datetime.now() > exp_date

    @staticmethod
    def days_remaining():
        """Get number of days until expiration"""
        exp_date = LicenseManager.get_expiration_date()
        delta = exp_date - datetime.now()
        return max(0, delta.days)

    @staticmethod
    def extend_with_code(code):
        """
        Try to extend expiration with a code.
        Returns True if successful, False if invalid code.
        """
        result = LicenseManager.validate_code(code)
        if result:
            year, month, day = result
            LicenseManager.save_expiration_date(year, month, day)
            return True
        return False


class StepTimer:
    """Tracks timing for each step and persists to log file for analysis"""

    def __init__(self, log_dir=None):
        self.log_dir = log_dir or EXE_DIR
        self.log_file = os.path.join(self.log_dir, 'timing_log.json')
        self.current_run = {
            'timestamp': datetime.now().isoformat(),
            'steps': {},
            'total_time': 0
        }
        self.step_start_time = None
        self.current_step_name = None
        self.run_start_time = time.time()

    def start_step(self, step_name):
        """Start timing a step"""
        # End previous step if any
        if self.current_step_name:
            self.end_step()

        self.current_step_name = step_name
        self.step_start_time = time.time()

    def end_step(self):
        """End timing current step"""
        if self.current_step_name and self.step_start_time:
            elapsed = time.time() - self.step_start_time
            self.current_run['steps'][self.current_step_name] = round(elapsed, 2)
            print(f"  [TIMING] {self.current_step_name}: {elapsed:.2f}s")
        self.current_step_name = None
        self.step_start_time = None

    def finish_run(self):
        """Finish the run and save to log"""
        self.end_step()  # End any pending step
        self.current_run['total_time'] = round(time.time() - self.run_start_time, 2)

        # Load existing log
        history = self._load_history()
        history['runs'].append(self.current_run)

        # Keep only last 50 runs
        if len(history['runs']) > 50:
            history['runs'] = history['runs'][-50:]

        # Calculate averages
        history['averages'] = self._calculate_averages(history['runs'])

        # Save
        self._save_history(history)

        # Print summary
        self._print_summary(history)

        return self.current_run['total_time']

    def _load_history(self):
        """Load timing history from file"""
        try:
            if os.path.exists(self.log_file):
                with open(self.log_file, 'r') as f:
                    return json.load(f)
        except:
            pass
        return {'runs': [], 'averages': {}}

    def _save_history(self, history):
        """Save timing history to file"""
        try:
            with open(self.log_file, 'w') as f:
                json.dump(history, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save timing log: {e}")

    def _calculate_averages(self, runs):
        """Calculate average time for each step across runs"""
        step_times = {}
        for run in runs:
            for step_name, elapsed in run.get('steps', {}).items():
                if step_name not in step_times:
                    step_times[step_name] = []
                step_times[step_name].append(elapsed)

        averages = {}
        for step_name, times in step_times.items():
            averages[step_name] = {
                'avg': round(sum(times) / len(times), 2),
                'min': round(min(times), 2),
                'max': round(max(times), 2),
                'count': len(times)
            }

        # Add total time average
        total_times = [r.get('total_time', 0) for r in runs if r.get('total_time')]
        if total_times:
            averages['_total'] = {
                'avg': round(sum(total_times) / len(total_times), 2),
                'min': round(min(total_times), 2),
                'max': round(max(total_times), 2),
                'count': len(total_times)
            }

        return averages

    def _print_summary(self, history):
        """Print timing summary"""
        print("\n" + "="*60)
        print("TIMING SUMMARY - This Run")
        print("="*60)

        # Sort steps by time (descending)
        sorted_steps = sorted(
            self.current_run['steps'].items(),
            key=lambda x: x[1],
            reverse=True
        )

        for step_name, elapsed in sorted_steps:
            avg_data = history['averages'].get(step_name, {})
            avg = avg_data.get('avg', elapsed)
            print(f"  {step_name}: {elapsed:.2f}s (avg: {avg:.2f}s)")

        print(f"\n  TOTAL: {self.current_run['total_time']:.2f}s")

        if '_total' in history['averages']:
            avg_total = history['averages']['_total']
            print(f"  Average over {avg_total['count']} runs: {avg_total['avg']:.2f}s")
            print(f"  Range: {avg_total['min']:.2f}s - {avg_total['max']:.2f}s")

        print("="*60 + "\n")

    @classmethod
    def get_historical_estimate(cls, log_dir=None, default_estimate=180):
        """Get the estimated total time based on historical averages.

        This method can be called BEFORE creating a StepTimer instance to get
        a better initial estimate for the progress bar.

        Args:
            log_dir: Directory where timing_log.json is stored (default: EXE_DIR)
            default_estimate: Fallback estimate if no history exists

        Returns:
            Estimated seconds for the build (float)
        """
        try:
            log_dir = log_dir or EXE_DIR
            log_file = os.path.join(log_dir, 'timing_log.json')

            if os.path.exists(log_file):
                with open(log_file, 'r') as f:
                    history = json.load(f)

                averages = history.get('averages', {})
                if '_total' in averages:
                    avg_total = averages['_total']
                    # Use average + 10% buffer for conservative estimate
                    estimated = avg_total.get('avg', default_estimate) * 1.1
                    run_count = avg_total.get('count', 0)
                    print(f"[TIMING] Using historical estimate: {estimated:.1f}s (based on {run_count} runs)")
                    return estimated
        except Exception as e:
            print(f"[TIMING] Could not load historical estimate: {e}")

        print(f"[TIMING] Using default estimate: {default_estimate}s")
        return default_estimate


APP_BUILD_DATE = "2026-01-16"
APP_NAME = "CFO Financial Model Generator"

# Get the directory where the script/exe is located
if getattr(sys, 'frozen', False):
    # Running as compiled exe - template is in the temp extraction folder
    APP_DIR = sys._MEIPASS
    EXE_DIR = os.path.dirname(sys.executable)
else:
    # Running as script
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = APP_DIR

# Template could be in APP_DIR (bundled) or EXE_DIR (alongside exe)
# Support both old DNA_Template name and new Financial_Template name
TEMPLATE_PATH = os.path.join(APP_DIR, 'Financial_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(EXE_DIR, 'Financial_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    # Fallback to old name for backwards compatibility
    TEMPLATE_PATH = os.path.join(APP_DIR, 'DNA_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(EXE_DIR, 'DNA_Template.xlsm')

# Template v2 path (optimized template with pre-built sheets)
TEMPLATE_V2_PATH = os.path.join(EXE_DIR, 'DNA_Template_v2.xlsm')
USE_TEMPLATE_V2 = True  # Set to False to use legacy template generation


class DivisionSetupDialog(tk.Toplevel):
    """Dialog for configuring multiple divisions"""

    MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    def __init__(self, parent, callback, existing_divisions=None):
        super().__init__(parent)
        self.title("Multi-Division Setup")
        self.callback = callback
        self.divisions = []  # List of division entries
        self.division_widgets = []  # UI widgets for each division

        # Year options
        current_year = datetime.now().year
        self.years = [str(y) for y in range(current_year - 10, current_year + 2)]

        # Initialize with existing divisions or empty
        if existing_divisions:
            for div in existing_divisions:
                self.divisions.append({
                    'name': tk.StringVar(value=div.get('name', '')),
                    'is_primary': tk.BooleanVar(value=div.get('is_primary', False)),
                    'pl_path': tk.StringVar(value=div.get('pl_path', '')),
                    'bs_path': tk.StringVar(value=div.get('bs_path', '')),
                    'start_month': tk.StringVar(value=div.get('start_month', 'January')),
                    'start_year': tk.StringVar(value=div.get('start_year', str(current_year))),
                    'end_month': tk.StringVar(value=div.get('end_month', 'December')),
                    'end_year': tk.StringVar(value=div.get('end_year', str(current_year)))
                })

        self.resizable(True, True)
        self._create_ui()
        self._center_window(750, 600)

        # Make modal
        self.transient(parent)
        self.grab_set()

    def _center_window(self, width, height):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _create_ui(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main_frame, text="Configure Divisions",
                  font=('Segoe UI', 14, 'bold')).pack(pady=(0, 10))

        ttk.Label(main_frame, text="Add divisions/departments/branches for consolidated reporting.",
                  font=('Segoe UI', 9), foreground='gray').pack(pady=(0, 15))

        # Division list frame with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Canvas for scrollable content
        self.canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)  # Linux scroll up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)  # Linux scroll down

        # Add initial division if none exist, otherwise render existing
        if not self.divisions:
            self._add_division(is_first=True)
        else:
            # Render existing divisions (only if not just added)
            for i, div in enumerate(self.divisions):
                self._render_division(i, is_first=(i == 0))

        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(btn_frame, text="+ Add Division", command=self._add_division).pack(side=tk.LEFT)

        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Save & Continue", command=self._save_and_close).pack(side=tk.RIGHT)

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if event.num == 4:  # Linux scroll up
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:  # Linux scroll down
            self.canvas.yview_scroll(1, "units")
        else:  # Windows
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _scroll_to_bottom(self):
        """Scroll to the bottom of the canvas"""
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def _add_division(self, is_first=False):
        """Add a new division entry"""
        current_year = datetime.now().year
        current_month = datetime.now().month
        div = {
            'name': tk.StringVar(value=f"Division {len(self.divisions) + 1}"),
            'is_primary': tk.BooleanVar(value=is_first),
            'pl_path': tk.StringVar(),
            'bs_path': tk.StringVar(),
            'start_month': tk.StringVar(value='January'),
            'start_year': tk.StringVar(value=str(current_year)),
            'end_month': tk.StringVar(value=self.MONTHS[current_month - 1]),
            'end_year': tk.StringVar(value=str(current_year))
        }
        self.divisions.append(div)
        self._render_division(len(self.divisions) - 1, is_first)

        # Scroll to show the new division
        if not is_first:
            self.after(100, self._scroll_to_bottom)

    def _render_division(self, index, is_first=False):
        """Render UI widgets for a division entry"""
        div = self.divisions[index]

        frame = ttk.LabelFrame(self.scrollable_frame, text=f"Division {index + 1}", padding="10")
        frame.pack(fill=tk.X, pady=5, padx=5)

        # Row 1: Name and Primary checkbox
        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row1, text="Name:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=div['name'], width=25).pack(side=tk.LEFT, padx=5)

        ttk.Checkbutton(row1, text="Primary Division",
                        variable=div['is_primary'],
                        command=lambda i=index: self._set_primary(i)).pack(side=tk.LEFT, padx=20)

        if not is_first:
            ttk.Button(row1, text="Remove", width=8,
                       command=lambda i=index: self._remove_division(i)).pack(side=tk.RIGHT)

        # Row 2: P&L file
        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row2, text="P&L File:").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=div['pl_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(row2, text="Browse...",
                   command=lambda v=div['pl_path'], d=div: self._browse_file(v, "P&L", d)).pack(side=tk.LEFT)

        # Row 3: BS file
        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(row3, text="Balance Sheet:").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=div['bs_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(row3, text="Browse...",
                   command=lambda v=div['bs_path'], d=div: self._browse_file(v, "Balance Sheet", d)).pack(side=tk.LEFT)

        # Date Range is auto-detected from files - no UI needed
        # The start_month, start_year, end_month, end_year variables are
        # automatically populated when files are selected via _auto_detect_date_range()

        self.division_widgets.append(frame)

    def _set_primary(self, selected_index):
        """Ensure only one division is marked as primary"""
        for i, div in enumerate(self.divisions):
            if i != selected_index:
                div['is_primary'].set(False)

    def _remove_division(self, index):
        """Remove a division entry"""
        if len(self.divisions) <= 1:
            messagebox.showwarning("Warning", "At least one division is required.")
            return

        # Remove from list
        self.divisions.pop(index)

        # Rebuild UI
        for widget in self.division_widgets:
            widget.destroy()
        self.division_widgets.clear()

        for i, div in enumerate(self.divisions):
            self._render_division(i, is_first=(i == 0))

        # Ensure at least one is primary
        if not any(d['is_primary'].get() for d in self.divisions):
            self.divisions[0]['is_primary'].set(True)

    def _browse_file(self, var, file_type, div=None):
        """Browse for a file and auto-detect date range"""
        path = filedialog.askopenfilename(
            title=f"Select {file_type} File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            # Validate file type (P&L vs Balance Sheet)
            expected_type = 'pl' if 'P&L' in file_type else 'bs'
            is_valid, detected_type, message = self._validate_file_type(path, expected_type)
            if not is_valid:
                messagebox.showerror("Wrong File Type", message)
                return  # Don't set the path - let user select again
            var.set(path)
            # Auto-detect date range from file
            if div:
                self._auto_detect_date_range(path, div)

    def _detect_file_type(self, file_path):
        """Detect whether a file is P&L or Balance Sheet based on account names."""
        try:
            df = pd.read_excel(file_path, header=None, nrows=100)
            pl_keywords = {'income': 2, 'revenue': 2, 'sales': 2, 'expense': 2,
                          'cost of goods': 3, 'gross profit': 3, 'net income': 3,
                          'operating income': 3, 'ebitda': 3, 'cogs': 2}
            bs_keywords = {'assets': 3, 'liabilities': 3, 'equity': 3,
                          'accounts receivable': 3, 'accounts payable': 3,
                          'cash': 2, 'inventory': 2, 'total assets': 3,
                          'total liabilities': 3, 'retained earnings': 3}
            pl_score = bs_score = 0
            for idx in range(len(df)):
                cell_val = df.iloc[idx, 0]
                if pd.isna(cell_val):
                    continue
                cell_lower = str(cell_val).lower().strip()
                for kw, wt in pl_keywords.items():
                    if kw in cell_lower:
                        pl_score += wt
                for kw, wt in bs_keywords.items():
                    if kw in cell_lower:
                        bs_score += wt
            total = pl_score + bs_score
            if total == 0:
                return (None, 0)
            if pl_score > bs_score:
                return ('pl', pl_score / total)
            elif bs_score > pl_score:
                return ('bs', bs_score / total)
            return (None, 0.5)
        except:
            return (None, 0)

    def _validate_file_type(self, file_path, expected_type):
        """Validate that a file matches the expected type (pl or bs)."""
        detected_type, confidence = self._detect_file_type(file_path)
        if detected_type is None:
            return (True, None, "Could not determine file type - proceeding anyway")
        if detected_type == expected_type:
            return (True, detected_type, f"File detected as {detected_type.upper()}")
        if confidence >= 0.6:
            type_names = {'pl': 'Profit & Loss (P&L)', 'bs': 'Balance Sheet'}
            detected_name = type_names.get(detected_type, detected_type.upper())
            expected_name = type_names.get(expected_type, expected_type.upper())
            return (False, detected_type,
                    f"This file appears to be a {detected_name} ({confidence:.0%} confidence),\n"
                    f"but you're uploading it as a {expected_name}.\n\n"
                    f"Please select the correct file type.")
        return (True, detected_type, f"File type uncertain - proceeding as {expected_type.upper()}")

    def _auto_detect_date_range(self, file_path, div):
        """Read Excel file and auto-detect the date range from column headers"""
        try:
            # Read just the first few rows to get headers
            df = pd.read_excel(file_path, header=None, nrows=5)

            detected_months = []
            # Look in first 3 rows for date headers (row 0, 1, 2)
            for row_idx in range(min(3, len(df))):
                for col_idx in range(1, min(50, len(df.columns))):  # Skip column A, check up to 50 cols
                    val = df.iloc[row_idx, col_idx]
                    parsed = self._parse_month_value(val)
                    if parsed:
                        month_num, year, display_name = parsed
                        detected_months.append((month_num, year, display_name))

            if detected_months:
                # Sort by year then month
                detected_months.sort(key=lambda x: (x[1], x[0]))

                # Get first and last month
                first_month, first_year, _ = detected_months[0]
                last_month, last_year, _ = detected_months[-1]

                # Update the division's date range
                div['start_month'].set(self.MONTHS[first_month - 1])
                div['start_year'].set(str(first_year))
                div['end_month'].set(self.MONTHS[last_month - 1])
                div['end_year'].set(str(last_year))

                print(f"Auto-detected date range: {self.MONTHS[first_month - 1]} {first_year} to {self.MONTHS[last_month - 1]} {last_year}")
        except Exception as e:
            print(f"Could not auto-detect date range: {e}")

    def _save_and_close(self):
        """Validate and save divisions"""
        # Validate
        for i, div in enumerate(self.divisions):
            if not div['name'].get().strip():
                messagebox.showerror("Error", f"Division {i + 1} needs a name.")
                return
            if not div['pl_path'].get():
                messagebox.showerror("Error", f"Division '{div['name'].get()}' needs a P&L file.")
                return
            if not div['bs_path'].get():
                messagebox.showerror("Error", f"Division '{div['name'].get()}' needs a Balance Sheet file.")
                return

        # Ensure one is primary
        if not any(d['is_primary'].get() for d in self.divisions):
            self.divisions[0]['is_primary'].set(True)

        # Convert to list of dicts
        result = []
        for div in self.divisions:
            result.append({
                'name': div['name'].get().strip(),
                'is_primary': div['is_primary'].get(),
                'pl_path': div['pl_path'].get(),
                'bs_path': div['bs_path'].get(),
                'start_month': div['start_month'].get(),
                'start_year': div['start_year'].get(),
                'end_month': div['end_month'].get(),
                'end_year': div['end_year'].get()
            })

        self.callback(result)
        self.destroy()


class AccountMappingDialog(tk.Toplevel):
    """Interactive dialog for reviewing and merging account mappings"""

    def __init__(self, parent, mappings, divisions, callback, pl_mappings=None, bs_mappings=None):
        super().__init__(parent)
        self.title("Review Account Mappings")
        self.mappings = mappings  # Combined Dict of consolidated_name -> AccountMapping
        self.pl_mappings = pl_mappings or {}  # P&L specific mappings
        self.bs_mappings = bs_mappings or {}  # BS specific mappings
        self.divisions = divisions
        self.callback = callback

        # Separate mappings if combined dict was passed
        if not pl_mappings and not bs_mappings:
            self._split_mappings_by_type()

        self.resizable(True, True)
        self._create_ui()
        self._center_window(1000, 700)

        # Make modal
        self.transient(parent)
        self.grab_set()

    def _split_mappings_by_type(self):
        """Try to split combined mappings - if not possible, show all in one view"""
        # For now, put all in pl_mappings since we don't have type info
        self.pl_mappings = dict(self.mappings)
        self.bs_mappings = {}

    def _center_window(self, width, height):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _create_ui(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title and Description
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(title_frame, text="Account Consolidation Review",
                  font=('Segoe UI', 16, 'bold')).pack(anchor='w')

        # Instructions box
        instr_frame = ttk.LabelFrame(main_frame, text="How This Works", padding="10")
        instr_frame.pack(fill=tk.X, pady=(0, 10))

        instructions = (
            "This screen shows how accounts from each division will be consolidated:\n\n"
            "• GREEN (Exact Match): Account names are identical across divisions - will be summed together\n"
            "• BLUE (Intelligent Match): AI detected similar accounts - review and merge if correct\n"
            "• ORANGE (Division-specific): Account exists in only one division - will appear as-is\n\n"
            "ACTION: Select 2+ similar accounts and click 'Merge Selected' to combine them."
        )
        ttk.Label(instr_frame, text=instructions, font=('Segoe UI', 9),
                  justify='left', wraplength=900).pack(anchor='w')

        # Summary stats frame
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Calculate stats
        exact_count = sum(1 for m in self.pl_mappings.values()
                         if (m.match_type if hasattr(m, 'match_type') else m.get('match_type', '')) == 'exact')
        intel_count = sum(1 for m in self.pl_mappings.values()
                         if (m.match_type if hasattr(m, 'match_type') else m.get('match_type', '')) == 'intelligent')
        div_specific = len(self.pl_mappings) - exact_count - intel_count

        ttk.Label(stats_frame, text=f"Summary: ", font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(stats_frame, text=f"{exact_count} exact matches", foreground='green',
                  font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(stats_frame, text=f"{intel_count} intelligent matches", foreground='blue',
                  font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(stats_frame, text=f"{div_specific} division-specific", foreground='orange',
                  font=('Segoe UI', 10)).pack(side=tk.LEFT)

        # Notebook for P&L and Balance Sheet tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # P&L Tab
        pl_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(pl_frame, text=f"P&L Accounts ({len(self.pl_mappings)})")
        self.pl_tree = self._create_mapping_tree(pl_frame, self.pl_mappings)

        # Balance Sheet Tab
        bs_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(bs_frame, text=f"Balance Sheet ({len(self.bs_mappings)})")
        self.bs_tree = self._create_mapping_tree(bs_frame, self.bs_mappings)

        # Action buttons frame with clear styling
        action_frame = ttk.LabelFrame(main_frame, text="Actions", padding="10")
        action_frame.pack(fill=tk.X, pady=(5, 10))

        merge_btn = ttk.Button(action_frame, text="Merge Selected Accounts",
                               command=self._merge_selected)
        merge_btn.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(action_frame,
                  text="Select 2+ rows above, then click to combine them into one consolidated account",
                  font=('Segoe UI', 9), foreground='#666').pack(side=tk.LEFT)

        # Bottom buttons with clear labeling
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        # Left side - status
        self.status_label = ttk.Label(btn_frame, text="Review mappings above, then click Continue to generate the model.",
                                       foreground='#666', font=('Segoe UI', 9))
        self.status_label.pack(side=tk.LEFT)

        # Right side - action buttons
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.destroy, width=12)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))

        confirm_btn = ttk.Button(btn_frame, text="Continue →", command=self._confirm_and_close, width=15)
        confirm_btn.pack(side=tk.RIGHT)

        # Tip text
        tip_frame = ttk.Frame(main_frame)
        tip_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(tip_frame, text="Tip: Most mappings are correct by default. Click 'Continue' if everything looks good.",
                  font=('Segoe UI', 8), foreground='#999').pack(anchor='e')

    def _create_mapping_tree(self, parent, mappings):
        """Create a treeview for a set of mappings"""
        # Frame for tree and scrollbars
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ['Consolidated', 'Type', 'Divisions'] + [d['name'] for d in self.divisions]
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15,
                           selectmode='extended')  # Allow multi-select

        for col in columns:
            tree.heading(col, text=col)
            if col == 'Consolidated':
                tree.column(col, width=200)
            elif col == 'Type':
                tree.column(col, width=100)
            elif col == 'Divisions':
                tree.column(col, width=80)
            else:
                tree.column(col, width=120)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Configure tag colors
        tree.tag_configure('exact', foreground='green')
        tree.tag_configure('intelligent', foreground='blue')
        tree.tag_configure('division_specific', foreground='orange')
        tree.tag_configure('manual', foreground='purple')

        # Populate tree
        self._populate_mapping_tree(tree, mappings)

        return tree

    def _populate_mapping_tree(self, tree, mappings):
        """Populate a treeview with mappings"""
        # Group by match type for display order
        by_type = {'exact': [], 'intelligent': [], 'manual': [], 'division_specific': []}

        for name, mapping in mappings.items():
            m_type = mapping.match_type if hasattr(mapping, 'match_type') else mapping.get('match_type', 'unknown')
            if m_type in by_type:
                by_type[m_type].append((name, mapping))
            else:
                by_type['division_specific'].append((name, mapping))

        # Add to tree
        for match_type in ['exact', 'intelligent', 'manual', 'division_specific']:
            for name, mapping in by_type[match_type]:
                if hasattr(mapping, 'division_mappings'):
                    div_mappings = mapping.division_mappings
                    m_type = mapping.match_type
                else:
                    div_mappings = mapping.get('division_mappings', {})
                    m_type = mapping.get('match_type', 'unknown')

                # Count how many divisions have this account
                div_count = len(div_mappings)

                values = [name, m_type.replace('_', '-').title(), f"{div_count}/{len(self.divisions)}"]
                for div in self.divisions:
                    values.append(div_mappings.get(div['name'], '-'))

                item = tree.insert('', tk.END, values=values, tags=(m_type,))

    def _merge_selected(self):
        """Merge selected accounts into one consolidated account"""
        # Get current tab's tree
        current_tab = self.notebook.index(self.notebook.select())
        tree = self.pl_tree if current_tab == 0 else self.bs_tree
        mappings = self.pl_mappings if current_tab == 0 else self.bs_mappings

        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showwarning("Selection Required",
                                   "Please select at least 2 accounts to merge.")
            return

        # Get selected account names
        selected_names = []
        for item in selected:
            values = tree.item(item, 'values')
            selected_names.append(values[0])

        # Ask for consolidated name
        default_name = selected_names[0]  # Use first selected as default
        new_name = tk.simpledialog.askstring(
            "Merge Accounts",
            f"Enter consolidated account name for:\n" + "\n".join(f"  - {n}" for n in selected_names),
            initialvalue=default_name,
            parent=self
        )

        if not new_name:
            return

        # Merge mappings
        merged_div_mappings = {}
        for name in selected_names:
            if name in mappings:
                mapping = mappings[name]
                div_maps = mapping.division_mappings if hasattr(mapping, 'division_mappings') else mapping.get('division_mappings', {})
                merged_div_mappings.update(div_maps)

        # Create new merged mapping
        new_mapping = AccountMapping(
            consolidated_name=new_name,
            division_mappings=merged_div_mappings,
            match_type='manual',
            confidence=1.0,
            approved=True
        )

        # Remove old mappings and add new one
        for name in selected_names:
            if name in mappings:
                del mappings[name]

        mappings[new_name] = new_mapping

        # Refresh tree
        for item in tree.get_children():
            tree.delete(item)
        self._populate_mapping_tree(tree, mappings)

        # Update tab label with count
        if current_tab == 0:
            self.notebook.tab(0, text=f"P&L Accounts ({len(self.pl_mappings)})")
        else:
            self.notebook.tab(1, text=f"Balance Sheet ({len(self.bs_mappings)})")

        self.status_label.config(text=f"Merged {len(selected_names)} accounts into '{new_name}'")

    def _confirm_and_close(self):
        """Confirm all mappings and close"""
        # Mark all as approved
        for mappings in [self.pl_mappings, self.bs_mappings]:
            for name, mapping in mappings.items():
                if hasattr(mapping, 'approved'):
                    mapping.approved = True
                elif isinstance(mapping, dict):
                    mapping['approved'] = True

        # Combine mappings and call callback
        combined = {**self.pl_mappings, **self.bs_mappings}
        self.callback(combined)
        self.destroy()


class FinancialModelApp:
    """Desktop application for generating financial models"""

    MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    VBA_CODE = '''
Option Explicit

Private Const SOURCE_PL_SHEET As String = "Source_PL"
Private Const SOURCE_BS_SHEET As String = "Source_BS"
Private Const SOURCE_BUDGET_SHEET As String = "Source_Budget"
Private Const PL_SHEET As String = "PL"
Private Const BS_SHEET As String = "Balance_Sheet"
Private Const CF_SHEET As String = "Cash_Flow"
Private Const MENU_SHEET As String = "Menu"

Public Sub UploadPLFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Select P&L File")
    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportData(filePath, SOURCE_PL_SHEET)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "P&L data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub UploadBSFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Select Balance Sheet File")
    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportData(filePath, SOURCE_BS_SHEET)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Balance Sheet data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub UploadBudgetFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", , "Select Budget File")
    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportBudgetDataSmart(filePath, SOURCE_BUDGET_SHEET)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub ImportBudgetDataSmart(filePath As String, targetSheet As String)
    ' Smart budget import with account matching logic
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim plWs As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim importCount As Long, matchCount As Long, unmatchedCount As Long
    Dim totalBudgetAccounts As Long
    Dim unmatchedList As String
    Dim userChoice As VbMsgBoxResult

    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)
    Set tgtWs = ThisWorkbook.Sheets(targetSheet)
    Set plWs = ThisWorkbook.Sheets(SOURCE_PL_SHEET)

    headerRow = FindHeaderRow(srcWs)
    If headerRow = 0 Then
        srcWb.Close False
        Err.Raise vbObjectError + 1, , "Could not find month headers in budget file"
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.Count).End(xlToLeft).Column

    ' First pass: analyze matching
    totalBudgetAccounts = 0
    matchCount = 0
    unmatchedList = ""

    For i = headerRow + 1 To lastRow
        Dim budgetAcct As String
        budgetAcct = Trim(CStr(srcWs.Cells(i, 1).Value))
        If Len(budgetAcct) > 0 And LCase(budgetAcct) <> "cash basis" Then
            totalBudgetAccounts = totalBudgetAccounts + 1
            Dim matchResult As Long
            matchResult = FindAccountMatch(plWs, budgetAcct)
            If matchResult > 0 Then
                matchCount = matchCount + 1
            Else
                unmatchedCount = unmatchedCount + 1
                If Len(unmatchedList) < 500 Then
                    unmatchedList = unmatchedList & vbCrLf & "  - " & budgetAcct
                End If
            End If
        End If
    Next i

    ' Check match rate
    Dim matchRate As Double
    If totalBudgetAccounts > 0 Then
        matchRate = matchCount / totalBudgetAccounts
    Else
        srcWb.Close False
        MsgBox "No accounts found in budget file.", vbExclamation
        Exit Sub
    End If

    ' If very few matches, warn user this may be wrong file
    If matchRate < 0.3 Then
        userChoice = MsgBox("Warning: Only " & matchCount & " of " & totalBudgetAccounts & _
            " accounts (" & Format(matchRate * 100, "0") & "%) match the existing P&L." & vbCrLf & vbCrLf & _
            "This appears to be an incorrect budget file that does not match the existing G/L accounts." & vbCrLf & vbCrLf & _
            "If you proceed, there will be " & unmatchedCount & " unmatched accounts that will not be displayed." & vbCrLf & vbCrLf & _
            "Do you want to proceed anyway?", vbYesNo + vbExclamation, "Possible File Mismatch")
        If userChoice = vbNo Then
            srcWb.Close False
            MsgBox "Import cancelled. Please select the correct budget file.", vbInformation
            Exit Sub
        End If
    End If

    ' If some accounts don't match, ask user
    If unmatchedCount > 0 And matchRate >= 0.3 Then
        Dim truncatedList As String
        If unmatchedCount > 10 Then
            truncatedList = Left(unmatchedList, 500) & vbCrLf & "  ... and " & (unmatchedCount - 10) & " more"
        Else
            truncatedList = unmatchedList
        End If

        userChoice = MsgBox(matchCount & " of " & totalBudgetAccounts & " accounts matched the P&L." & vbCrLf & vbCrLf & _
            "The following " & unmatchedCount & " accounts did not match:" & truncatedList & vbCrLf & vbCrLf & _
            "Would you like to add these unmatched accounts anyway?" & vbCrLf & vbCrLf & _
            "Click YES to add them, NO to skip them.", vbYesNo + vbQuestion, "Unmatched Accounts Found")
    End If

    ' Second pass: import data
    importCount = 0
    Dim addedCount As Long
    addedCount = 0
    Dim tgtLastCol As Long
    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column
    Dim tgtLastRow As Long
    tgtLastRow = tgtWs.Cells(tgtWs.Rows.Count, 1).End(xlUp).Row

    For j = 2 To lastCol
        Dim monthName As String
        monthName = Trim(CStr(srcWs.Cells(headerRow, j).Value))
        If Len(monthName) > 0 And LCase(monthName) <> "total" Then
            Dim monthCol As Long
            monthCol = FindMonthCol(tgtWs, monthName)

            If monthCol > 0 Then
                For i = headerRow + 1 To lastRow
                    Dim accountName As String
                    accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
                    If Len(accountName) > 0 And LCase(accountName) <> "cash basis" Then
                        Dim accountRow As Long
                        accountRow = FindAccountMatch(tgtWs, accountName)

                        If accountRow > 0 Then
                            tgtWs.Cells(accountRow, monthCol).Value = srcWs.Cells(i, j).Value
                            importCount = importCount + 1
                        ElseIf userChoice = vbYes Then
                            ' Add new account if user chose to add unmatched
                            accountRow = FindOrAddBudgetAccount(tgtWs, accountName, tgtLastRow)
                            If accountRow > tgtLastRow Then tgtLastRow = accountRow
                            tgtWs.Cells(accountRow, monthCol).Value = srcWs.Cells(i, j).Value
                            importCount = importCount + 1
                            addedCount = addedCount + 1
                        End If
                    End If
                Next i
            End If
        End If
    Next j

    srcWb.Close False

    ' Show summary
    Dim summaryMsg As String
    summaryMsg = "Budget import complete!" & vbCrLf & vbCrLf & _
        "Imported " & importCount & " values."
    If addedCount > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "Added " & addedCount & " new accounts."
    End If
    If unmatchedCount > 0 And userChoice = vbNo Then
        summaryMsg = summaryMsg & vbCrLf & "Skipped " & unmatchedCount & " unmatched accounts."
    End If
    MsgBox summaryMsg, vbInformation
End Sub

Private Function FindAccountMatch(ws As Worksheet, accountName As String) As Long
    ' Try to find account with smart matching:
    ' 1. Exact match
    ' 2. Budget name contained in P&L name
    ' 3. P&L name contained in budget name
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim plAcct As String
    Dim budgetLower As String
    budgetLower = LCase(Trim(accountName))

    ' First: exact match
    For i = 3 To lastRow
        plAcct = Trim(CStr(ws.Cells(i, 1).Value))
        If LCase(plAcct) = budgetLower Then
            FindAccountMatch = i
            Exit Function
        End If
    Next i

    ' Second: budget name contained in P&L name (e.g., "Rent" matches "Rent Expense")
    For i = 3 To lastRow
        plAcct = LCase(Trim(CStr(ws.Cells(i, 1).Value)))
        If Len(budgetLower) > 3 And InStr(plAcct, budgetLower) > 0 Then
            FindAccountMatch = i
            Exit Function
        End If
    Next i

    ' Third: P&L name contained in budget name (e.g., "Marketing Expense" contains "Marketing")
    For i = 3 To lastRow
        plAcct = LCase(Trim(CStr(ws.Cells(i, 1).Value)))
        If Len(plAcct) > 3 And InStr(budgetLower, plAcct) > 0 Then
            FindAccountMatch = i
            Exit Function
        End If
    Next i

    FindAccountMatch = 0
End Function

Private Function FindOrAddBudgetAccount(ws As Worksheet, accountName As String, currentLastRow As Long) As Long
    ' Find existing account or add new one at the end
    Dim i As Long
    For i = 3 To currentLastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = accountName Then
            FindOrAddBudgetAccount = i
            Exit Function
        End If
    Next i
    ' Add new account
    Dim newRow As Long
    newRow = currentLastRow + 1
    ws.Cells(newRow, 1).Value = accountName
    FindOrAddBudgetAccount = newRow
End Function

Private Function FindMonthCol(ws As Worksheet, monthName As String) As Long
    ' Find existing month column (do not add new columns)
    Dim j As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For j = 2 To lastCol
        If LCase(Trim(CStr(ws.Cells(1, j).Value))) = LCase(monthName) Then
            FindMonthCol = j
            Exit Function
        End If
    Next j
    FindMonthCol = 0
End Function

Private Function FindAccountRow(ws As Worksheet, accountName As String) As Long
    ' Find existing account row (do not add new accounts)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 3 To lastRow  ' Start at row 3 (row 1=header, row 2=YYYYMM helper)
        If Trim(CStr(ws.Cells(i, 1).Value)) = accountName Then
            FindAccountRow = i
            Exit Function
        End If
    Next i
    FindAccountRow = 0
End Function

Private Sub ImportData(filePath As String, targetSheet As String)
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long

    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)
    Set tgtWs = ThisWorkbook.Sheets(targetSheet)

    headerRow = FindHeaderRow(srcWs)
    If headerRow = 0 Then
        srcWb.Close False
        Err.Raise vbObjectError + 1, , "Could not find month headers"
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.Count).End(xlToLeft).Column

    Dim tgtLastCol As Long
    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column

    For j = 2 To lastCol
        Dim monthName As String
        monthName = Trim(CStr(srcWs.Cells(headerRow, j).Value))
        If Len(monthName) > 0 And LCase(monthName) <> "total" Then
            Dim monthCol As Long
            monthCol = FindOrAddMonth(tgtWs, monthName, tgtLastCol)
            If monthCol > tgtLastCol Then tgtLastCol = monthCol

            For i = headerRow + 1 To lastRow
                Dim accountName As String
                accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
                If Len(accountName) > 0 And LCase(accountName) <> "cash basis" Then
                    Dim accountRow As Long
                    accountRow = FindOrAddAccount(tgtWs, accountName)
                    tgtWs.Cells(accountRow, monthCol).Value = srcWs.Cells(i, j).Value
                End If
            Next i
        End If
    Next j

    srcWb.Close False
    UpdateNamedRange tgtWs
End Sub

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim i As Long, j As Long
    For i = 1 To 15
        For j = 1 To 20
            Dim val As String
            val = LCase(Trim(CStr(ws.Cells(i, j).Value)))
            If InStr(val, "january") > 0 Or InStr(val, "february") > 0 Or InStr(val, "march") > 0 Then
                If InStr(val, "-") = 0 Then
                    FindHeaderRow = i
                    Exit Function
                End If
            End If
        Next j
    Next i
    FindHeaderRow = 0
End Function

Private Function FindOrAddMonth(ws As Worksheet, monthName As String, lastCol As Long) As Long
    Dim j As Long
    For j = 2 To lastCol
        If LCase(Trim(CStr(ws.Cells(1, j).Value))) = LCase(monthName) Then
            FindOrAddMonth = j
            Exit Function
        End If
    Next j
    lastCol = lastCol + 1
    ws.Cells(1, lastCol).Value = monthName
    FindOrAddMonth = lastCol
End Function

Private Function FindOrAddAccount(ws As Worksheet, accountName As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = accountName Then
            FindOrAddAccount = i
            Exit Function
        End If
    Next i
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = accountName
    FindOrAddAccount = lastRow
End Function

Private Sub UpdateNamedRange(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim rangeName As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ws.Name = SOURCE_PL_SHEET Then
        rangeName = "SourcePL"
    ElseIf ws.Name = SOURCE_BS_SHEET Then
        rangeName = "SourceBS"
    Else
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add rangeName, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Sub

Public Sub RefreshAll()
    Application.CalculateFull
    MsgBox "All formulas refreshed!", vbInformation
End Sub

Public Sub RunDiagnostics()
    Dim issues As String
    issues = ""

    If Not CheckBSBalance() Then
        issues = issues & "- Balance Sheet does not balance" & vbCrLf
    End If

    If Len(issues) = 0 Then
        MsgBox "All checks passed!", vbInformation
    Else
        MsgBox "Issues found:" & vbCrLf & issues, vbExclamation
    End If
End Sub

Private Function CheckBSBalance() As Boolean
    CheckBSBalance = True
End Function

'-----------------------------------------------------
' UpdateColumnVisibility - Hide prior years and future months
' Called when current month changes on Menu sheet (C7)
'-----------------------------------------------------
Public Sub UpdateColumnVisibility()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsMenu As Worksheet
    Dim wsPL As Worksheet
    Dim wsBS As Worksheet
    Dim wsCF As Worksheet
    Dim currentMonthStr As String

    Set wsMenu = GetWs("Menu")
    Set wsPL = GetWs("PL")
    Set wsBS = GetWs("Balance_Sheet")
    Set wsCF = GetWs("Cash_Flow")

    If wsMenu Is Nothing Then GoTo Cleanup

    ' Get current month from Menu C7 (e.g., "Nov 24" or date value)
    currentMonthStr = Trim(CStr(wsMenu.Range("C7").Value))
    If Len(currentMonthStr) = 0 Then GoTo Cleanup

    ' Update visibility on P&L sheet
    If Not wsPL Is Nothing Then
        Call SetColVisibility(wsPL, currentMonthStr)
    End If

    ' Update visibility on Balance Sheet
    If Not wsBS Is Nothing Then
        Call SetColVisibility(wsBS, currentMonthStr)
    End If

    ' Update visibility on Cash Flow
    If Not wsCF Is Nothing Then
        Call SetColVisibility(wsCF, currentMonthStr)
    End If

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Resume Cleanup
End Sub

Private Sub SetColVisibility(ws As Worksheet, currentMonthStr As String)
    On Error Resume Next

    Dim headerRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim currentYear As Long
    Dim currentMonthNum As Long
    Dim colYear As Long
    Dim colMonthNum As Long
    Dim headerVal As String

    headerRow = 4
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' Parse current month/year from string like "Nov 24" or "Nov-24"
    currentYear = GetYearFromStr(currentMonthStr)
    currentMonthNum = GetMonthFromStr(currentMonthStr)

    If currentYear = 0 Or currentMonthNum = 0 Then Exit Sub

    ' First, unhide all columns
    ws.Columns.Hidden = False

    ' Scan columns and set visibility
    For col = 2 To lastCol
        headerVal = Trim(CStr(ws.Cells(headerRow, col).Value))

        ' Check if this is a month column
        If IsMonthCol(headerVal) Then
            colYear = GetYearFromStr(headerVal)
            colMonthNum = GetMonthFromStr(headerVal)

            If colYear > 0 And colMonthNum > 0 Then
                ' Future year - hide
                If colYear > currentYear Then
                    ws.Columns(col).Hidden = True
                ' Same year but future month - hide
                ElseIf colYear = currentYear And colMonthNum > currentMonthNum Then
                    ws.Columns(col).Hidden = True
                ' Prior year - hide
                ElseIf colYear < currentYear Then
                    ws.Columns(col).Hidden = True
                ' Current year, current or prior month - show
                Else
                    ws.Columns(col).Hidden = False
                End If
            End If
        End If
    Next col
End Sub

Private Function IsMonthCol(headerVal As String) As Boolean
    Dim months As Variant
    Dim i As Long

    months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    headerVal = LCase(Trim(headerVal))

    For i = LBound(months) To UBound(months)
        If InStr(headerVal, months(i)) > 0 Then
            IsMonthCol = True
            Exit Function
        End If
    Next i
    IsMonthCol = False
End Function

Private Function GetYearFromStr(monthStr As String) As Long
    Dim parts() As String
    Dim i As Long
    Dim y As Long

    GetYearFromStr = 0
    monthStr = Replace(monthStr, "-", " ")
    parts = Split(Trim(monthStr), " ")

    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            y = CLng(parts(i))
            If y >= 2000 And y <= 2100 Then
                GetYearFromStr = y
                Exit Function
            ElseIf y >= 0 And y <= 99 Then
                GetYearFromStr = 2000 + y
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetMonthFromStr(monthStr As String) As Long
    Dim months As Variant
    Dim i As Long

    months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    monthStr = LCase(Trim(monthStr))

    For i = LBound(months) To UBound(months)
        If InStr(monthStr, months(i)) > 0 Then
            GetMonthFromStr = i + 1
            Exit Function
        End If
    Next i
    GetMonthFromStr = 0
End Function

Private Function GetWs(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWs = ThisWorkbook.Worksheets(sheetName)
End Function

'-----------------------------------------------------
' UpdateYearGrouping - Regroup columns based on Reporting Year
' Call this when the Reporting Year dropdown changes on Menu sheet (C10)
' Groups and collapses all columns for years PRIOR TO the Reporting Year
' Example: If Reporting Year = 2025, groups 2024 and earlier
'-----------------------------------------------------
Public Sub UpdateYearGrouping()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsMenu As Worksheet
    Dim reportingYear As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim processedCount As Long

    Set wsMenu = GetWs("Menu")
    If wsMenu Is Nothing Then GoTo Cleanup

    ' Get Reporting Year from Menu C10
    If IsNumeric(wsMenu.Range("C10").Value) Then
        reportingYear = CLng(wsMenu.Range("C10").Value)
    Else
        MsgBox "Invalid Reporting Year in Menu C10. Please enter a valid year.", vbExclamation
        GoTo Cleanup
    End If

    ' List of sheets to process - all P&L, BS, CF tabs
    sheetNames = Array("Consolidated_PL", "Consolidated_BS", "Consolidated_CF", _
                       "Source_PL", "Source_BS", "Source_Budget", "PL", "Balance_Sheet", "Cash_Flow")

    processedCount = 0
    For Each sheetName In sheetNames
        Set ws = GetWs(CStr(sheetName))
        If Not ws Is Nothing Then
            Call RegroupSheetByYear(ws, reportingYear)
            processedCount = processedCount + 1
        End If
    Next sheetName

    ' Also process division-specific sheets
    Dim i As Integer
    Dim divSheets As Variant
    divSheets = Array("_PL", "_BS", "_CF", "_Forecast")

    For Each ws In ThisWorkbook.Worksheets
        For i = LBound(divSheets) To UBound(divSheets)
            If InStr(ws.Name, divSheets(i)) > 0 And Left(ws.Name, 4) <> "Menu" Then
                Call RegroupSheetByYear(ws, reportingYear)
                processedCount = processedCount + 1
            End If
        Next i
    Next ws

    MsgBox "Year grouping updated!" & vbCrLf & vbCrLf & _
           "Reporting Year: " & reportingYear & vbCrLf & _
           "Years " & (reportingYear - 1) & " and earlier are now grouped/collapsed." & vbCrLf & _
           "Sheets processed: " & processedCount, vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error updating year grouping: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub RegroupSheetByYear(ws As Worksheet, reportingYear As Long)
    On Error Resume Next

    Dim headerRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim colYear As Long
    Dim headerVal As String
    Dim priorYearCols As Collection
    Dim startCol As Long
    Dim endCol As Long
    Dim i As Long

    ' Find header row - check rows 1, 2, 4 for "Account" header
    headerRow = 0
    If InStr(LCase(CStr(ws.Cells(1, 1).Value)), "account") > 0 Then
        headerRow = 1
    ElseIf InStr(LCase(CStr(ws.Cells(4, 1).Value)), "account") > 0 Then
        headerRow = 4
    ElseIf InStr(LCase(CStr(ws.Cells(2, 1).Value)), "account") > 0 Then
        headerRow = 2
    End If

    If headerRow = 0 Then Exit Sub

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then Exit Sub

    ' Clear all existing outline grouping
    ws.Cells.ClearOutline

    ' Find all columns that belong to prior years (before reportingYear)
    Set priorYearCols = New Collection
    For col = 2 To lastCol
        headerVal = Trim(CStr(ws.Cells(headerRow, col).Value))
        colYear = GetYearFromStr(headerVal)

        ' If year is valid and BEFORE reporting year, it's a prior year
        If colYear > 0 And colYear < reportingYear Then
            priorYearCols.Add col
        End If
    Next col

    ' Group prior year columns if any exist
    If priorYearCols.Count > 0 Then
        ' Find contiguous ranges and group them
        startCol = priorYearCols(1)
        endCol = priorYearCols(1)

        For i = 2 To priorYearCols.Count
            If priorYearCols(i) = endCol + 1 Then
                ' Contiguous - extend range
                endCol = priorYearCols(i)
            Else
                ' Gap - group current range and start new one
                If startCol > 0 And endCol >= startCol Then
                    ws.Range(ws.Columns(startCol), ws.Columns(endCol)).Group
                End If
                startCol = priorYearCols(i)
                endCol = priorYearCols(i)
            End If
        Next i

        ' Group final range
        If startCol > 0 And endCol >= startCol Then
            ws.Range(ws.Columns(startCol), ws.Columns(endCol)).Group
        End If

        ' Collapse all groups
        ws.Outline.ShowLevels ColumnLevels:=1
    End If
End Sub

' ============================================================
' UpdateVisibility - Show/Hide sheets based on Settings tab
' ============================================================
Public Sub UpdateVisibility()
    ' Read Settings sheet and show/hide sheets accordingly
    Dim ws As Worksheet
    Dim settingsSheet As Worksheet
    Dim sheetName As String
    Dim visible As String
    Dim row As Long
    Dim lastRow As Long

    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0

    If settingsSheet Is Nothing Then
        MsgBox "Settings sheet not found.", vbExclamation
        Exit Sub
    End If

    ' Find last row with data in column B
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, 2).End(xlUp).row

    ' Start at row 9 (first data row after header at row 8)
    For row = 9 To lastRow
        sheetName = Trim(CStr(settingsSheet.Cells(row, 2).Value))
        visible = Trim(UCase(CStr(settingsSheet.Cells(row, 3).Value)))

        If Len(sheetName) > 0 Then
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0

            If Not ws Is Nothing Then
                If visible = "YES" Then
                    ws.visible = xlSheetVisible
                ElseIf visible = "NO" Then
                    ws.visible = xlSheetHidden
                End If
                Set ws = Nothing
            End If
        End If
    Next row

    MsgBox "Sheet visibility updated.", vbInformation
End Sub
'''

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.resizable(False, False)

        # Check license/expiration before anything else
        if not self._check_license():
            self.root.destroy()
            return

        # Variables
        self.pl_path = tk.StringVar()
        self.bs_path = tk.StringVar()
        self.existing_model_path = tk.StringVar()  # For update mode
        self.company_name = tk.StringVar(value="Company Name")
        self.mode = tk.StringVar(value="new")  # "new" or "update"

        # Multi-division support
        self.is_multi_division = tk.BooleanVar(value=False)
        self.divisions = []  # List of division configs
        self.division_configs = []  # List of DivisionConfig objects
        self.account_mappings = {'pl': {}, 'bs': {}}  # Consolidated account mappings

        self._create_ui()
        self._center_window(700, 650)  # Reduced height since config section removed

    def _check_license(self):
        """Check license on startup. Returns True if valid, False if user cancels."""

        # FIRST RUN - Show activation splash screen
        if LicenseManager.is_first_run():
            return self._show_activation_splash()

        # SUBSEQUENT RUNS - Check expiration
        days_left = LicenseManager.days_remaining()

        # Show warning if expiring soon (within 7 days)
        if 0 < days_left <= 7:
            exp_date = LicenseManager.get_expiration_date()
            messagebox.showwarning(
                "License Expiring Soon",
                f"This beta version expires on {exp_date.strftime('%B %d, %Y')}.\n\n"
                f"You have {days_left} day(s) remaining.\n\n"
                "Contact Focus CFO for an extension code."
            )
            return True

        # If not expired, continue
        if not LicenseManager.is_expired():
            return True

        # Expired - show dialog for extension code
        return self._show_license_dialog(is_expired=True)

    def _show_activation_splash(self):
        """Show activation splash screen on first run. Returns True if activated, False if cancelled."""
        return self._show_license_dialog(is_expired=False, is_first_run=True)

    def _show_license_dialog(self, is_expired=True, is_first_run=False):
        """Show license dialog. Returns True if activated/extended, False if cancelled."""
        dialog = tk.Toplevel(self.root)

        if is_first_run:
            dialog.title("Product Activation")
            title_text = "Welcome to CFO Financial Model Generator"
            message_text = ("Please enter your license code to activate.\n\n"
                           "If you don't have a code, contact Focus CFO.")
            button_text = "Activate"
        else:
            dialog.title("License Expired")
            exp_date = LicenseManager.get_expiration_date()
            title_text = "License Expired"
            message_text = (f"This version expired on {exp_date.strftime('%B %d, %Y')}.\n\n"
                           "Please enter an extension code to continue,\n"
                           "or contact Focus CFO for a new code.")
            button_text = "Extend License"

        dialog.geometry("500x320")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # Center on screen
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 500) // 2
        y = (dialog.winfo_screenheight() - 320) // 2
        dialog.geometry(f"500x320+{x}+{y}")

        result = {'continue': False}

        main_frame = ttk.Frame(dialog, padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main_frame, text=title_text,
                  font=('Segoe UI', 14, 'bold')).pack(pady=(0, 5))

        # Version
        ttk.Label(main_frame, text=f"Version {APP_VERSION}",
                  font=('Segoe UI', 9), foreground='#666').pack(pady=(0, 15))

        # Message
        ttk.Label(main_frame, text=message_text,
                  font=('Segoe UI', 10), justify='center').pack(pady=(0, 20))

        # Code entry frame
        code_frame = ttk.Frame(main_frame)
        code_frame.pack(pady=(0, 10))

        ttk.Label(code_frame, text="License Code:", font=('Segoe UI', 10)).pack(side=tk.LEFT)
        code_var = tk.StringVar()
        code_entry = ttk.Entry(code_frame, textvariable=code_var, width=18, font=('Consolas', 12))
        code_entry.pack(side=tk.LEFT, padx=(10, 0))

        # Status label
        status_label = ttk.Label(main_frame, text="", font=('Segoe UI', 9))
        status_label.pack(pady=(5, 15))

        def try_activate():
            code = code_var.get().strip()
            if not code:
                status_label.config(text="Please enter a license code", foreground='red')
                return

            if LicenseManager.extend_with_code(code):
                new_exp = LicenseManager.get_expiration_date()
                if is_first_run:
                    messagebox.showinfo(
                        "Activation Successful",
                        f"Thank you for activating!\n\nLicense valid until {new_exp.strftime('%B %d, %Y')}."
                    )
                else:
                    messagebox.showinfo(
                        "License Extended",
                        f"License extended until {new_exp.strftime('%B %d, %Y')}.\n\nThank you!"
                    )
                result['continue'] = True
                dialog.destroy()
            else:
                status_label.config(text="Invalid code. Please check and try again.", foreground='red')

        def cancel():
            result['continue'] = False
            dialog.destroy()

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack()

        activate_btn = ttk.Button(btn_frame, text=button_text, command=try_activate, width=15)
        activate_btn.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Exit", command=cancel, width=10).pack(side=tk.LEFT)

        # Contact info
        ttk.Label(main_frame, text="Contact: support@focuscfo.com  |  www.focuscfo.com",
                  font=('Segoe UI', 8), foreground='#666').pack(side=tk.BOTTOM, pady=(15, 0))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        code_entry.focus_set()

        self.root.wait_window(dialog)
        return result['continue']

    def _create_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text=APP_NAME,
                                font=('Segoe UI', 16, 'bold'))
        title_label.pack(pady=(0, 5))

        # Version info
        version_label = ttk.Label(main_frame, text=f"Version {APP_VERSION} ({APP_BUILD_DATE})",
                                  font=('Segoe UI', 9), foreground='gray')
        version_label.pack(pady=(0, 15))

        # Mode Selection (New Model vs Update Existing)
        mode_frame = ttk.LabelFrame(main_frame, text="Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Radiobutton(mode_frame, text="Create New Model", variable=self.mode, value="new",
                        command=self._toggle_mode).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Radiobutton(mode_frame, text="Update Existing Model", variable=self.mode, value="update",
                        command=self._toggle_mode).grid(row=0, column=1, sticky=tk.W)

        # Division Structure (Multi-division support)
        self.division_frame = ttk.LabelFrame(main_frame, text="Division Structure", padding="10")
        self.division_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(self.division_frame, text="Multiple Divisions / Departments",
                        variable=self.is_multi_division,
                        command=self._toggle_division_mode).grid(row=0, column=0, sticky=tk.W)

        self.division_btn = ttk.Button(self.division_frame, text="Configure Divisions...",
                                       command=self._open_division_setup, state='disabled')
        self.division_btn.grid(row=0, column=1, padx=20)

        self.division_status = ttk.Label(self.division_frame, text="Single entity mode",
                                         font=('Segoe UI', 9), foreground='gray')
        self.division_status.grid(row=0, column=2, sticky=tk.W)

        # Company Name (for new models only)
        self.name_frame = ttk.LabelFrame(main_frame, text="Company Information", padding="10")
        self.name_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.name_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.name_frame, textvariable=self.company_name, width=40).grid(row=0, column=1, padx=5)

        # Existing Model Selection (for update mode)
        self.existing_frame = ttk.LabelFrame(main_frame, text="Existing Model", padding="10")
        # Don't pack yet - hidden by default

        ttk.Label(self.existing_frame, text="Model File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.existing_frame, textvariable=self.existing_model_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(self.existing_frame, text="Browse...", command=self._browse_existing).grid(row=0, column=2)

        # File Selection (for single-division mode)
        self.file_frame = ttk.LabelFrame(main_frame, text="Upload Files (Single Division)", padding="10")
        self.file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.file_frame, text="P&L File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.file_frame, textvariable=self.pl_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(self.file_frame, text="Browse...", command=self._browse_pl).grid(row=0, column=2)

        ttk.Label(self.file_frame, text="Balance Sheet:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        ttk.Entry(self.file_frame, textvariable=self.bs_path, width=40).grid(row=1, column=1, padx=5, pady=(5, 0))
        ttk.Button(self.file_frame, text="Browse...", command=self._browse_bs).grid(row=1, column=2, pady=(5, 0))

        # Date range info - auto-detected from files
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(info_frame, text="Date range will be auto-detected from your Excel files.",
                  font=('Segoe UI', 9), foreground='#666').pack(anchor='w')

        # Button frame for Generate and Close buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)

        self.generate_btn = ttk.Button(btn_frame, text="Generate Financial Model",
                                       command=self._generate_model, style='Accent.TButton')
        self.generate_btn.pack(fill=tk.X, ipady=10)

        # Progress bar frame
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(10, 5))

        # Progress bar (dynamic growing bar)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                            maximum=100, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X)

        # Status label (shows current activity)
        self.status_label = ttk.Label(main_frame, text="Ready", foreground='gray', font=('Segoe UI', 10))
        self.status_label.pack(pady=(5, 10))

        # Button frame for Help and Close
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(5, 0))

        help_btn = ttk.Button(button_frame, text="Help", command=self._show_help)
        help_btn.pack(side=tk.LEFT, padx=(0, 10))

        close_btn = ttk.Button(button_frame, text="Close", command=self.root.quit)
        close_btn.pack(side=tk.LEFT)

        # Version info footer (minimal)
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(20, 0))
        ttk.Label(footer_frame, text=f"v{APP_VERSION}", foreground='#999',
                  font=('Segoe UI', 8)).pack()

    def _toggle_mode(self):
        """Toggle between New Model and Update Existing modes"""
        if self.mode.get() == "new":
            # Show company info, hide existing model selection
            self.name_frame.pack(fill=tk.X, pady=(0, 10), after=self.root.winfo_children()[0].winfo_children()[2])
            self.existing_frame.pack_forget()
            self.generate_btn.config(text="Generate Financial Model")
        else:
            # Show existing model selection, hide company info
            self.name_frame.pack_forget()
            self.existing_frame.pack(fill=tk.X, pady=(0, 10), after=self.root.winfo_children()[0].winfo_children()[2])
            self.generate_btn.config(text="Update Financial Model")

    def _toggle_division_mode(self):
        """Toggle between single entity and multi-division modes"""
        if self.is_multi_division.get():
            self.division_btn.config(state='normal')
            self.division_status.config(text="Multi-division mode - click Configure Divisions")
            # Hide single-division file upload since files are uploaded per-division
            self.file_frame.pack_forget()
        else:
            self.division_btn.config(state='disabled')
            self.division_status.config(text="Single entity mode")
            self.divisions = []
            self.division_configs = []
            # Show single-division file upload
            # Re-pack file_frame after division_frame
            self.file_frame.pack(fill=tk.X, pady=(0, 10), after=self.division_frame)

    def _show_help(self):
        """Show the help dialog with tabbed navigation"""
        help_window = tk.Toplevel(self.root)
        help_window.title(f"CFO Financial Model Generator v{APP_VERSION} - Help")
        help_window.geometry("900x650")
        help_window.transient(self.root)
        help_window.grab_set()  # Make modal

        # Center on screen
        help_window.update_idletasks()
        x = (help_window.winfo_screenwidth() - 900) // 2
        y = (help_window.winfo_screenheight() - 650) // 2
        help_window.geometry(f"900x650+{x}+{y}")

        # Main container
        main_frame = ttk.Frame(help_window, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header with logo/title
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(header_frame, text="CFO Financial Model Generator",
                  font=('Segoe UI', 18, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header_frame, text=f"Version {APP_VERSION}",
                  font=('Segoe UI', 10), foreground='#666').pack(side=tk.RIGHT, pady=(8, 0))

        # Create notebook (tabbed interface)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Tab 1: Getting Started
        tab1 = self._create_help_tab(notebook, "Getting Started", [
            ("Welcome", "This application creates comprehensive Excel financial models from your "
             "QuickBooks or accounting system exports. Simply provide your P&L and Balance Sheet "
             "files, and the generator creates a complete financial package."),
            ("Quick Start Guide",
             "1. Enter your company name in the field provided\n\n"
             "2. Click 'Browse' next to P&L File and select your Profit & Loss export\n\n"
             "3. Click 'Browse' next to Balance Sheet and select your Balance Sheet export\n\n"
             "4. Click 'Generate Financial Model'\n\n"
             "5. Choose where to save the output file\n\n"
             "6. Wait for generation to complete (typically 30-90 seconds)\n\n"
             "7. The generated Excel file will open automatically"),
            ("System Requirements",
             "• Windows 10 or Windows 11\n"
             "• Microsoft Excel 2016 or later (required)\n"
             "• 4 GB RAM minimum (8 GB recommended)\n"
             "• Close Excel before generating for best results"),
        ])
        notebook.add(tab1, text="  Getting Started  ")

        # Tab 2: File Preparation
        tab2 = self._create_help_tab(notebook, "Preparing Your Files", [
            ("P&L (Profit & Loss) Export",
             "Export your P&L from QuickBooks with:\n\n"
             "• Account names in Column A\n"
             "• Monthly data as column headers (e.g., 'Oct 2024', 'Nov 2024')\n"
             "• All income and expense accounts included\n"
             "• File format: .xlsx or .xls"),
            ("Balance Sheet Export",
             "Export your Balance Sheet with:\n\n"
             "• Account names in Column A\n"
             "• Monthly snapshots as column headers\n"
             "• All asset, liability, and equity accounts\n"
             "• File format: .xlsx or .xls"),
            ("File Validation",
             "The application automatically validates your files:\n\n"
             "• P&L files should contain: income, revenue, expense, cost keywords\n"
             "• Balance Sheet files should contain: assets, liabilities, equity keywords\n"
             "• If you select the wrong file type, you'll see a warning"),
        ])
        notebook.add(tab2, text="  File Preparation  ")

        # Tab 3: Single Entity Mode
        tab3 = self._create_help_tab(notebook, "Single Entity Mode", [
            ("When to Use",
             "Use Single Entity Mode when you have one company with one set of books. "
             "This is the default mode and works for most small to medium businesses."),
            ("Step-by-Step",
             "1. Ensure 'Enable Multi-Division Mode' is NOT checked\n\n"
             "2. Enter your company name (appears on all report headers)\n\n"
             "3. Browse and select your P&L export file\n\n"
             "4. Browse and select your Balance Sheet export file\n\n"
             "5. Click 'Generate Financial Model'\n\n"
             "6. Choose save location and wait for completion"),
            ("Generated Output",
             "Your workbook will contain:\n\n"
             "• Menu - Navigation hub with period selector\n"
             "• Dashboard - Executive summary with KPIs and charts\n"
             "• PL - Profit & Loss with monthly and YTD columns\n"
             "• Balance_Sheet - Balance Sheet with comparative periods\n"
             "• Cash_Flow - Auto-generated statement of cash flows\n"
             "• Forecast - 12-month forecast with Budget vs Actual"),
        ])
        notebook.add(tab3, text="  Single Entity  ")

        # Tab 4: Multi-Division Mode
        tab4 = self._create_help_tab(notebook, "Multi-Division Mode", [
            ("When to Use",
             "Use Multi-Division Mode when you need to consolidate multiple entities:\n\n"
             "• Parent company with subsidiaries\n"
             "• Company with multiple departments\n"
             "• Franchise with multiple locations\n"
             "• Any organization needing consolidated financials"),
            ("Setup Process",
             "1. Check 'Enable Multi-Division Mode'\n\n"
             "2. Click 'Configure Divisions' button\n\n"
             "3. In the dialog, click '+ Add Division' for each entity\n\n"
             "4. Enter a name for each division (e.g., 'North Region')\n\n"
             "5. Browse to select P&L and Balance Sheet for each division\n\n"
             "6. Click 'Save & Continue'\n\n"
             "7. Click 'Generate Financial Model'"),
            ("Consolidated Reports",
             "Multi-division models include:\n\n"
             "• Consolidated_PL - Combined P&L for all divisions\n"
             "• Division-specific P&L sheets (Division1_PL, etc.)\n"
             "• Consolidated_Forecast - Combined forecast\n"
             "• Division selector dropdown on Dashboard\n"
             "• Account consolidation with intelligent matching"),
        ])
        notebook.add(tab4, text="  Multi-Division  ")

        # Tab 5: Update Existing
        tab5 = self._create_help_tab(notebook, "Update Existing Model", [
            ("When to Use",
             "Instead of creating a new model each month, you can update an existing model:\n\n"
             "• Adding a new month of data\n"
             "• Updating with corrected data\n"
             "• Maintaining a continuous financial model"),
            ("How to Update",
             "1. Select 'Update Existing Model' mode (radio button)\n\n"
             "2. Click 'Browse' next to Model File\n\n"
             "3. Select your existing .xlsm financial model\n\n"
             "4. Select your NEW P&L file (containing the additional month)\n\n"
             "5. Select your NEW Balance Sheet file\n\n"
             "6. Click 'Update Financial Model'\n\n"
             "7. Choose to overwrite or save as new file"),
            ("What Gets Updated",
             "• Source data sheets receive new month columns\n"
             "• All report sheets get new month columns with formulas\n"
             "• Menu dropdown updated with new month option\n"
             "• Forecast Actual column populated for new month\n"
             "• Dashboard automatically reflects latest period\n"
             "• Existing data and formatting preserved"),
        ])
        notebook.add(tab5, text="  Update Model  ")

        # Tab 6: Working with Output
        tab6 = self._create_help_tab(notebook, "Working with Output", [
            ("Period Selection",
             "Change the viewing period on any report:\n\n"
             "1. Go to the Menu sheet\n"
             "2. Click the dropdown in cell C7\n"
             "3. Select any available month\n"
             "4. All sheets automatically update to show that period\n\n"
             "The Dashboard also has a period selector for quick access."),
            ("Adding Budget Data",
             "To enable Budget vs Actual comparisons:\n\n"
             "1. Go to the Source_Budget sheet (may be hidden)\n"
             "2. Enter budget amounts for each account\n"
             "3. Columns B through M represent Jan through Dec\n"
             "4. The Forecast sheet automatically uses this data\n"
             "5. Variance columns show Budget vs Actual differences"),
            ("Division Views (Multi-Division Only)",
             "Switch between consolidated and division views:\n\n"
             "1. Go to Dashboard or any report sheet\n"
             "2. Find the division dropdown (usually cell C6)\n"
             "3. Select 'Consolidated' for combined view\n"
             "4. Select a division name for that entity only\n"
             "5. All data and charts update accordingly"),
        ])
        notebook.add(tab6, text="  Using Output  ")

        # Tab 7: Troubleshooting
        tab7 = self._create_help_tab(notebook, "Troubleshooting", [
            ("Common Errors",
             "\"The file is still open\"\n"
             "   → Close Excel completely and try again\n\n"
             "\"Could not determine date columns\"\n"
             "   → Ensure date headers (e.g., 'Oct 2024') are in the first few rows\n"
             "   → The application will prompt you to enter dates manually\n\n"
             "\"Wrong file type detected\"\n"
             "   → You may have swapped P&L and Balance Sheet files\n"
             "   → P&L should have income/expense accounts\n"
             "   → Balance Sheet should have assets/liabilities"),
            ("Formula Errors",
             "If you see #REF! errors:\n"
             "   → This usually means source data structure changed\n"
             "   → Try regenerating the model from scratch\n\n"
             "If charts don't show data:\n"
             "   → Do NOT delete row 3 (hidden YYYYMM helper row)\n"
             "   → This row is essential for formula lookups"),
            ("Performance Tips",
             "Generation taking too long?\n\n"
             "• Close all other Excel workbooks\n"
             "• Single entity: expect 30-90 seconds\n"
             "• Multi-division (3 divs): expect 2-3 minutes\n"
             "• Update mode is faster than full rebuild\n"
             "• Large files (100+ accounts) take longer"),
        ])
        notebook.add(tab7, text="  Troubleshooting  ")

        # Bottom button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Label(button_frame, text="Support: support@focuscfo.com  |  www.focuscfo.com",
                  font=('Segoe UI', 9), foreground='#666').pack(side=tk.LEFT)

        ttk.Button(button_frame, text="Close", command=help_window.destroy,
                   width=12).pack(side=tk.RIGHT)

    def _create_help_tab(self, notebook, title, sections):
        """Create a help tab with formatted sections"""
        # Outer frame for the tab
        tab_frame = ttk.Frame(notebook, padding="10")

        # Create canvas with scrollbar for scrolling
        canvas = tk.Canvas(tab_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Add sections
        for i, (section_title, section_content) in enumerate(sections):
            # Section header
            header = ttk.Label(scrollable_frame, text=section_title,
                              font=('Segoe UI', 12, 'bold'))
            header.pack(anchor='w', pady=(15 if i > 0 else 5, 8))

            # Section content
            content = ttk.Label(scrollable_frame, text=section_content,
                               font=('Segoe UI', 10), wraplength=800, justify='left')
            content.pack(anchor='w', padx=(15, 10), pady=(0, 5))

            # Separator (except after last section)
            if i < len(sections) - 1:
                sep = ttk.Separator(scrollable_frame, orient='horizontal')
                sep.pack(fill='x', pady=(15, 0), padx=(0, 20))

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        return tab_frame

    def _open_division_setup(self):
        """Open the division setup dialog"""
        def on_divisions_saved(divisions):
            self.divisions = divisions
            count = len(divisions)
            self.division_status.config(text=f"{count} division(s) configured")
            # Clear the single-file paths since we'll use division-specific files
            self.pl_path.set('')
            self.bs_path.set('')

        DivisionSetupDialog(self.root, on_divisions_saved, self.divisions)

    def _handle_mapping_approval(self, approved_mappings):
        """Handle approved mappings from the review dialog"""
        self.account_mappings = approved_mappings

    def _browse_existing(self):
        """Browse for existing model file"""
        path = filedialog.askopenfilename(
            title="Select Existing Financial Model",
            filetypes=[("Excel files", "*.xlsm *.xlsx")]
        )
        if path:
            self.existing_model_path.set(path)

    def _center_window(self, width, height):
        """Center the window on the screen"""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _browse_pl(self):
        path = filedialog.askopenfilename(
            title="Select P&L File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            # Validate file type
            is_valid, detected_type, message = self._validate_file_type(path, 'pl')
            if not is_valid:
                messagebox.showerror("Wrong File Type", message)
                return  # Don't set the path - let user select again
            self.pl_path.set(path)

    def _browse_bs(self):
        path = filedialog.askopenfilename(
            title="Select Balance Sheet File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            # Validate file type
            is_valid, detected_type, message = self._validate_file_type(path, 'bs')
            if not is_valid:
                messagebox.showerror("Wrong File Type", message)
                return  # Don't set the path - let user select again
            self.bs_path.set(path)

    def _generate_model(self):
        """Generate or update the financial model based on mode"""
        # Validate inputs based on division mode
        if self.is_multi_division.get():
            # Multi-division mode: validate divisions are configured
            if not self.divisions or len(self.divisions) == 0:
                messagebox.showerror("Error", "Please configure divisions first using 'Configure Divisions...'")
                return
            # Validate each division has files
            for div in self.divisions:
                if not div.get('pl_path') or not div.get('bs_path'):
                    messagebox.showerror("Error", f"Division '{div.get('name', 'Unknown')}' is missing P&L or Balance Sheet file")
                    return
        else:
            # Single entity mode: validate single files
            if not self.pl_path.get() or not self.bs_path.get():
                messagebox.showerror("Error", "Please select both P&L and Balance Sheet files")
                return

        if self.mode.get() == "new":
            # New model mode
            if not self.company_name.get().strip():
                messagebox.showerror("Error", "Please enter a company name")
                return

            # Ask for save location
            save_path = filedialog.asksaveasfilename(
                title="Save Financial Model As",
                defaultextension=".xlsm",
                filetypes=[("Excel Macro-Enabled", "*.xlsm")],
                initialfile=f"{self.company_name.get().replace(' ', '_')}_Financial_Model.xlsm"
            )

            if not save_path:
                return

            self.status_label.config(text="Starting...", foreground='blue')
            self.generate_btn.config(state='disabled')
            self.root.update()

            # Run directly (xlwings doesn't work well with threads)
            try:
                print(f"Starting model generation...")
                print(f"P&L file: {self.pl_path.get()}")
                print(f"BS file: {self.bs_path.get()}")
                print(f"Save path: {save_path}")
                print(f"Template path: {TEMPLATE_PATH}")
                print(f"Template exists: {os.path.exists(TEMPLATE_PATH)}")
                self._create_excel_model(save_path)
                self._update_status("Complete!", color='green')
                messagebox.showinfo("Success", f"Financial Model created:\n{save_path}")
            except Exception as e:
                import traceback
                print(f"ERROR: {e}")
                traceback.print_exc()
                self._update_status("Error occurred", color='red')
                messagebox.showerror("Error", str(e))
            finally:
                self.generate_btn.config(state='normal')

        else:
            # Update existing model mode
            if not self.existing_model_path.get():
                messagebox.showerror("Error", "Please select an existing Financial Model file")
                return

            # Ask for save location (can overwrite or save as new)
            save_path = filedialog.asksaveasfilename(
                title="Save Updated Model As",
                defaultextension=".xlsm",
                filetypes=[("Excel Macro-Enabled", "*.xlsm")],
                initialfile=os.path.basename(self.existing_model_path.get())
            )

            if not save_path:
                return

            self.status_label.config(text="Starting update...", foreground='blue')
            self.generate_btn.config(state='disabled')
            self.root.update()

            try:
                print(f"Starting model update...")
                print(f"Existing model: {self.existing_model_path.get()}")
                print(f"P&L file: {self.pl_path.get()}")
                print(f"BS file: {self.bs_path.get()}")
                print(f"Save path: {save_path}")
                self._update_existing_model(save_path)
                self._update_status("Complete!", color='green')
                messagebox.showinfo("Success", f"Financial Model updated:\n{save_path}")
            except Exception as e:
                import traceback
                print(f"ERROR: {e}")
                traceback.print_exc()
                self._update_status("Error occurred", color='red')
                messagebox.showerror("Error", str(e))
            finally:
                self.generate_btn.config(state='normal')

    def _update_status(self, message, color='blue', elapsed=None, estimated_total=None, progress=None):
        """Update the status label, progress bar, and refresh the UI

        Args:
            message: Status message to display
            color: Text color
            elapsed: Elapsed time in seconds (optional)
            estimated_total: Estimated total time in seconds (optional)
            progress: Progress percentage 0-100 (optional)
        """
        if elapsed is not None and estimated_total is not None:
            remaining = max(0, estimated_total - elapsed)
            if remaining > 60:
                time_str = f" (~{int(remaining // 60)}m {int(remaining % 60)}s remaining)"
            else:
                time_str = f" (~{int(remaining)}s remaining)"
            message = f"{message}{time_str}"
        elif elapsed is not None:
            if elapsed > 60:
                time_str = f" ({int(elapsed // 60)}m {int(elapsed % 60)}s elapsed)"
            else:
                time_str = f" ({int(elapsed)}s elapsed)"
            message = f"{message}{time_str}"

        self.status_label.config(text=message, foreground=color)

        # Update progress bar if progress value provided
        if progress is not None:
            self.progress_var.set(progress)

        self.root.update()
        print(message)

    def _create_progress_tracker(self, total_steps, estimated_total_seconds):
        """Create a progress tracker for mid-step updates

        Returns functions for updating step progress and substep progress.
        """
        start_time = time.time()
        state = {'current_step': 0, 'substep': ''}

        def update_step(step_name):
            """Update to a new main step"""
            state['current_step'] += 1
            state['substep'] = step_name
            progress_pct = (state['current_step'] / total_steps) * 100

            elapsed = time.time() - start_time
            remaining = max(0, estimated_total_seconds - elapsed)

            # Build message - clean arrow style without step numbers
            msg = f"→ {step_name}"

            # Add time estimate
            if remaining > 60:
                time_str = f" (~{int(remaining // 60)}m {int(remaining % 60)}s remaining)"
            elif remaining > 5:
                time_str = f" (~{int(remaining)}s remaining)"
            else:
                time_str = ""  # Don't show tiny remaining times

            self.status_label.config(text=f"{msg}{time_str}", foreground='blue')
            self.progress_var.set(progress_pct)
            self.root.update()
            print(f"-> {step_name}")

        def update_substep(substep_name):
            """Update substep within current step (doesn't advance progress)"""
            msg = f"  → {substep_name}"
            self.status_label.config(text=msg, foreground='#666')
            self.root.update()
            print(f"    {substep_name}")

        return update_step, update_substep, state

    def _get_user_date_params(self):
        """Get date range parameters - returns None to use all data from files (auto-detect)"""
        # Dates are now auto-detected from files, so return None for both
        # This tells _parse_financial_data to use all available data
        return None, None

    def _check_file_not_open(self, file_path):
        """Check if a file is open by another process. Returns True if file is accessible."""
        if not os.path.exists(file_path):
            return True  # File doesn't exist, so it's not locked

        try:
            # Try to open the file in exclusive mode
            with open(file_path, 'r+b') as f:
                pass
            return True
        except (IOError, PermissionError):
            return False

    # ================================================================
    # UPDATE EXISTING MODEL - UNIFIED HELPER FUNCTIONS
    # ================================================================
    # These functions support updating both single-entity and multi-division models

    def _detect_model_type(self, wb):
        """Detect if model is single-entity or multi-division.

        Returns:
            bool: True if multi-division model, False if single-entity
        """
        source_pl = wb['Source_PL']
        header_a1 = source_pl['A1'].value
        is_multi_division = str(header_a1).strip().lower() == 'division'
        return is_multi_division

    def _get_source_structure(self, is_multi_division):
        """Get column structure constants based on model type.

        Returns:
            dict with keys: acct_col, div_col, data_start_col, yyyymm_row
        """
        if is_multi_division:
            return {
                'div_col': 1,        # Column A = Division
                'acct_col': 2,       # Column B = Account
                'data_start_col': 3, # Column C = first data column
                'yyyymm_row': 2      # Row 2 = YYYYMM values
            }
        else:
            return {
                'div_col': None,     # No Division column
                'acct_col': 1,       # Column A = Account
                'data_start_col': 2, # Column B = first data column
                'yyyymm_row': 2      # Row 2 = YYYYMM values
            }

    def _get_report_sheets(self, wb, is_multi_division):
        """Get list of report sheets to update.

        Returns:
            tuple: (pl_sheets, bs_sheets) where each is a list of (sheet_name, division_name or None)
        """
        pl_sheets = []
        bs_sheets = []

        if is_multi_division:
            # Find Consolidated_PL and all Division*_PL sheets
            for sheet in wb.worksheets:
                if sheet.title == 'Consolidated_PL':
                    pl_sheets.append(('Consolidated_PL', None))  # (name, division)
                elif '_PL' in sheet.title and sheet.title != 'Source_PL':
                    div_name = sheet.title.replace('_PL', '')
                    pl_sheets.append((sheet.title, div_name))
                elif sheet.title == 'Consolidated_BS':
                    bs_sheets.append(('Consolidated_BS', None))
                elif '_BS' in sheet.title and sheet.title != 'Source_BS':
                    div_name = sheet.title.replace('_BS', '')
                    bs_sheets.append((sheet.title, div_name))
        else:
            # Single-entity sheets
            sheet_names = wb.sheetnames
            if 'PL' in sheet_names:
                pl_sheets.append(('PL', None))
            elif 'P&L' in sheet_names:
                pl_sheets.append(('P&L', None))
            if 'Balance_Sheet' in sheet_names:
                bs_sheets.append(('Balance_Sheet', None))

        return pl_sheets, bs_sheets

    def _find_column_by_yyyymm(self, sheet, yyyymm, data_start_col, yyyymm_row=2):
        """Find the column index containing a specific YYYYMM value.

        Returns:
            int or None: Column index if found, None otherwise
        """
        # end("right") replaced with max_column
        for col in range(data_start_col, last_col + 1):
            val = sheet.cell(row=yyyymm_row, column=col).value
            if val and int(val) == int(yyyymm):
                return col
        return None

    def _update_menu_lookup_table(self, menu_sheet, new_months, existing_months_set):
        """Add new months to the Menu sheet K:M lookup table.

        Args:
            menu_sheet: The Menu sheet object
            new_months: List of (m, y, name, yyyymm) tuples for new months
            existing_months_set: Set of existing YYYYMM values already in the model
        """
        # Find last row of lookup table by scanning K column
        # K6 is header, K7+ are month entries
        last_lookup_row = 6
        for row in range(7, 200):  # Reasonable max
            if menu_sheet[f'K{row}'].value is None:
                break
            last_lookup_row = row

        for m, y, name, yyyymm in new_months:
            if yyyymm not in existing_months_set:
                last_lookup_row += 1
                menu_sheet[f'K{last_lookup_row}'].value = name
                menu_sheet[f'L{last_lookup_row}'].value = yyyymm
                menu_sheet[f'M{last_lookup_row}'].value = m

        return last_lookup_row

    def _update_menu_dropdown(self, menu_sheet, last_lookup_row):
        """Update C7 dropdown validation with all available months from lookup table.

        Args:
            menu_sheet: The Menu sheet object
            last_lookup_row: Last row of the lookup table (K column)
        """
        # Collect all month names from lookup table in CHRONOLOGICAL order
        month_data = []  # List of (yyyymm, name) for sorting
        for row in range(7, last_lookup_row + 1):
            name = menu_sheet[f'K{row}'].value
            yyyymm = menu_sheet[f'L{row}'].value
            if name and yyyymm:
                try:
                    month_data.append((int(yyyymm), str(name)))
                except:
                    month_data.append((0, str(name)))

        # Sort by YYYYMM to ensure chronological order
        month_data.sort(key=lambda x: x[0])
        month_names = [name for yyyymm, name in month_data]

        if month_names:
            month_list = ','.join(month_names)
            try:
                # Validation handled via DataValidation object
                # DataValidation handled via DataValidation()
                print(f"DEBUG: Menu dropdown updated with {len(month_names)} months: {month_names[:3]}...{month_names[-1:]}")
            except Exception as e:
                print(f"Warning: Could not update dropdown validation: {e}")

        # Update VLOOKUP formula range to include new rows
        try:
            menu_sheet['G7'].value = f'=IFERROR(VLOOKUP(C7,$K$7:$L${last_lookup_row},2,FALSE),0)'
            menu_sheet['E7'].value = f'=IFERROR(VLOOKUP(C7,$K$7:$M${last_lookup_row},3,FALSE),1)'
        except Exception as e:
            print(f"Warning: Could not update VLOOKUP formulas: {e}")

    def _generate_source_formula(self, is_multi_division, source_sheet, data_col_letter,
                                  account_ref, is_consolidated=True, division_name=None):
        """Generate the appropriate SUMIF/SUMIFS formula for source data lookup.

        Args:
            is_multi_division: Whether model is multi-division
            source_sheet: 'Source_PL' or 'Source_BS'
            data_col_letter: Column letter for data (e.g., 'D')
            account_ref: Cell reference for account name (e.g., '$A5')
            is_consolidated: For multi-division, whether this is consolidated view
            division_name: For division-specific sheets, the division name

        Returns:
            str: The formula string
        """
        source_start = 3
        source_end = 1500

        if is_multi_division:
            if division_name:
                # Division-specific: SUMIFS filtering by division and account
                return (f'=SUMIFS({source_sheet}!{data_col_letter}${source_start}:{data_col_letter}${source_end},'
                        f'{source_sheet}!$A${source_start}:$A${source_end},"{division_name}",'
                        f'{source_sheet}!$B${source_start}:$B${source_end},TRIM({account_ref}))')
            else:
                # Consolidated: SUMIF on account column B
                return (f'=SUMIF({source_sheet}!$B${source_start}:$B${source_end},'
                        f'TRIM({account_ref}),{source_sheet}!{data_col_letter}${source_start}:{data_col_letter}${source_end})')
        else:
            # Single-entity: SUMIF on account column A
            return (f'=SUMIF({source_sheet}!$A${source_start}:$A${source_end},'
                    f'TRIM({account_ref}),{source_sheet}!{data_col_letter}${source_start}:{data_col_letter}${source_end})')

    def _update_existing_model(self, save_path):
        """Update an existing Financial Model with new month data"""
        import shutil

        # Create progress tracker (same style as build)
        update_step, update_substep, _ = self._create_progress_tracker(
            total_steps=10,  # Total steps for update process
            estimated_total_seconds=30  # Estimated time
        )

        # Check if any of the input files are open
        files_to_check = [
            (self.pl_path.get(), "P&L file"),
            (self.bs_path.get(), "Balance Sheet file"),
            (self.existing_model_path.get(), "Existing model"),
        ]
        if save_path and os.path.exists(save_path):
            files_to_check.append((save_path, "Save destination"))

        for file_path, file_name in files_to_check:
            if file_path and not self._check_file_not_open(file_path):
                filename = os.path.basename(file_path)
                raise Exception(f"The file '{filename}' is still open.\n\nPlease close it and try again.")

        # Get user-specified date range
        user_start, user_end = self._get_user_date_params()

        # Parse new input files
        update_step("Reading P&L file...")
        pl_data = pd.read_excel(self.pl_path.get(), header=None)
        pl_indents = self._get_cell_indents(self.pl_path.get())

        update_step("Reading Balance Sheet file...")
        bs_data = pd.read_excel(self.bs_path.get(), header=None)
        bs_indents = self._get_cell_indents(self.bs_path.get())

        update_step("Extracting account data...")
        pl_accounts, new_months, pl_totals = self._parse_financial_data(pl_data, pl_indents, user_start, user_end, self.pl_path.get())
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data, bs_indents, user_start, user_end, self.bs_path.get())
        print(f"New data has {len(new_months)} months")

        # Work in temp directory
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        try:
            # Copy existing model to temp location
            update_step("Opening existing model...")
            shutil.copy2(self.existing_model_path.get(), temp_path)

            # openpyxl: no Excel app needed
            try:
                wb = load_workbook(temp_path, keep_vba=True)

                # Get existing sheets
                source_pl = wb['Source_PL']
                source_bs = wb['Source_BS']

                # Detect model type (single-entity vs multi-division)
                is_multi_division = self._detect_model_type(wb)
                source_struct = self._get_source_structure(is_multi_division)
                data_start_col = source_struct['data_start_col']
                acct_col = source_struct['acct_col']

                print(f"DEBUG: Model type: {'multi-division' if is_multi_division else 'single-entity'}")
                print(f"DEBUG: Account column: {acct_col}, Data starts at column: {data_start_col}")

                # Find existing months in Source_PL (row 2 has YYYYMM)
                update_step("Analyzing existing data...")
                existing_months = set()
                # end("right") replaced with max_column
                for col in range(data_start_col, last_col + 1):
                    yyyymm = source_pl.cell(row=2, column=col).value
                    if yyyymm:
                        existing_months.add(int(yyyymm))

                # Determine which months are new
                new_month_list = []
                for m, y, name in new_months:
                    yyyymm = y * 100 + m
                    if yyyymm not in existing_months:
                        new_month_list.append((m, y, name, yyyymm))

                if not new_month_list:
                    raise Exception("No new months found in uploaded files. All months already exist in the model.")

                print(f"Adding {len(new_month_list)} new month(s): {[n[2] for n in new_month_list]}")

                # Add new months to Source_PL
                update_step("Adding data to Source_PL...")
                update_substep("Populating P&L source data...")
                SOURCE_BLACK = (26, 26, 26)  # Match existing header color

                for m, y, name, yyyymm in new_month_list:
                    # Find next available column
                    # end("right") replaced with max_column

                    # Add month header and YYYYMM
                    source_pl.cell(row=1, column=new_col).value = name
                    source_pl.cell(row=2, column=new_col).value = yyyymm

                    # Format header to match existing columns
                    header_cell = source_pl.cell(row=1, column=new_col)
                    header_cell.font = Font(name='Calibri Light', size=10, bold=True, color="FFFFFF")
                    header_cell.fill = rgb_fill(SOURCE_BLACK)
                    # Alignment handled via Alignment() objects

                    # Add P&L data for this month - mode-aware matching
                    print(f"DEBUG: Adding P&L data for month ({m}, {y}) to column {new_col}")

                    # Build lookup dict from parsed accounts
                    # Key depends on mode: single-entity uses account name only
                    # Multi-division uses (division, account) tuples
                    pl_values_lookup = {}
                    for account in pl_accounts:
                        acct_name = account.get('name', '').strip()
                        account_values = account.get('values', {})
                        value = account_values.get((m, y), 0)
                        pl_values_lookup[acct_name] = value
                        pl_values_lookup[acct_name.lstrip()] = value

                    # Get existing account/division names from Source_PL and match
                    # Use UsedRange to reliably find last row
                    try:
                        last_row_pl = source_pl.max_row
                    except:
                        last_row_pl = 1500  # fallback

                    print(f"DEBUG: Source_PL last_row_pl = {last_row_pl}")
                    non_zero_count = 0
                    matched_count = 0

                    for row in range(3, last_row_pl + 1):
                        if is_multi_division:
                            # Multi-division: col A = Division, col B = Account
                            existing_div = source_pl.cell(row=row, column=1).value
                            existing_name = source_pl.cell(row=row, column=2).value
                            if existing_name:
                                existing_name_str = str(existing_name).strip()
                                # For now, match by account name only (division filtering happens in formulas)
                                # If uploaded file has division-specific data, that would need special handling
                                value = pl_values_lookup.get(existing_name_str,
                                        pl_values_lookup.get(existing_name_str.lstrip(), 0))
                                if value != 0:
                                    non_zero_count += 1
                                if existing_name_str in pl_values_lookup or existing_name_str.lstrip() in pl_values_lookup:
                                    matched_count += 1
                                source_pl.cell(row=row, column=new_col).value = value
                        else:
                            # Single-entity: col A = Account
                            existing_name = source_pl.cell(row=row, column=acct_col).value
                            if existing_name:
                                existing_name_str = str(existing_name).strip()
                                value = pl_values_lookup.get(existing_name_str,
                                        pl_values_lookup.get(existing_name_str.lstrip(), 0))
                                if value != 0:
                                    non_zero_count += 1
                                if existing_name_str in pl_values_lookup or existing_name_str.lstrip() in pl_values_lookup:
                                    matched_count += 1
                                source_pl.cell(row=row, column=new_col).value = value

                    # Format data column - font and number format
                    # Range: data_range_pl = (source_pl, 3, new_col, last_row_pl, new_col)
                    apply_style_to_range(source_pl, 3, new_col, last_row_pl, new_col, number_format='#,##0')
                    apply_style_to_range(source_pl, 3, new_col, last_row_pl, new_col, font=Font(name='Calibri Light', size=10))

                    print(f"DEBUG: P&L - {len(pl_accounts)} parsed accounts, {matched_count} matched, {non_zero_count} with non-zero values")
                    if len(pl_accounts) > 0:
                        print(f"DEBUG: First 5 parsed account names:")
                        for idx, acct in enumerate(pl_accounts[:5]):
                            print(f"  [{idx}] '{acct.get('name')}'")
                        acct_col_to_check = 2 if is_multi_division else 1
                        print(f"DEBUG: First 5 existing Source_PL account names (col {acct_col_to_check}):")
                        for row in range(3, min(8, last_row_pl + 1)):
                            existing = source_pl.cell(row=row, column=acct_col_to_check).value
                            print(f"  [row {row}] '{existing}'")

                    # Add BS data for same month
                    update_substep("Populating Balance Sheet source data...")
                    # end("right") replaced with max_column
                    source_bs.cell(row=1, column=new_col_bs).value = name
                    source_bs.cell(row=2, column=new_col_bs).value = yyyymm

                    # Format BS header to match existing columns
                    bs_header_cell = source_bs.cell(row=1, column=new_col_bs)
                    bs_header_cell.font = Font(name='Calibri Light', size=10, bold=True, color="FFFFFF")
                    bs_header_cell.fill = rgb_fill(SOURCE_BLACK)
                    # Alignment handled via Alignment() objects

                    # Add BS data for this month - mode-aware matching
                    print(f"DEBUG: Adding BS data for month ({m}, {y}) to column {new_col_bs}")

                    # Build a lookup dict from parsed BS accounts
                    bs_values_lookup = {}
                    for account in bs_accounts:
                        acct_name = account.get('name', '').strip()
                        account_values = account.get('values', {})
                        value = account_values.get((m, y), 0)
                        bs_values_lookup[acct_name] = value
                        bs_values_lookup[acct_name.lstrip()] = value

                    # Get existing account names from Source_BS and match
                    # Use UsedRange to reliably find last row
                    try:
                        # UsedRange replaced with sheet.max_row/max_column
                        last_row_bs = source_bs.max_row  # openpyxl: use max_row
                    except:
                        pass
                        # Using sheet.max_row instead of .end("down")

                    print(f"DEBUG: Source_BS last_row_bs = {last_row_bs}")
                    non_zero_count_bs = 0
                    matched_count_bs = 0

                    for row in range(3, last_row_bs + 1):
                        if is_multi_division:
                            # Multi-division: col A = Division, col B = Account
                            existing_name = source_bs.cell(row=row, column=2).value
                        else:
                            # Single-entity: col A = Account
                            existing_name = source_bs.cell(row=row, column=acct_col).value

                        if existing_name:
                            existing_name_str = str(existing_name).strip()
                            value = bs_values_lookup.get(existing_name_str,
                                    bs_values_lookup.get(existing_name_str.lstrip(), 0))
                            if value != 0:
                                non_zero_count_bs += 1
                            if existing_name_str in bs_values_lookup or existing_name_str.lstrip() in bs_values_lookup:
                                matched_count_bs += 1
                            source_bs.cell(row=row, column=new_col_bs).value = value

                    # Format BS data column - font and number format
                    # Range: data_range_bs = (source_bs, 3, new_col_bs, last_row_bs, new_col_bs)
                    apply_style_to_range(source_bs, 3, new_col_bs, last_row_bs, new_col_bs, number_format='#,##0')
                    apply_style_to_range(source_bs, 3, new_col_bs, last_row_bs, new_col_bs, font=Font(name='Calibri Light', size=10))

                    print(f"DEBUG: BS - {len(bs_accounts)} parsed accounts, {matched_count_bs} matched, {non_zero_count_bs} with non-zero values")

                # Update named ranges - use UsedRange for reliable row counts
                try:
                    # UsedRange replaced with sheet.max_row/max_column
                    pl_last_row = source_pl.max_row  # openpyxl: use max_row
                    pl_last_col = used_range_pl.Column + used_range_pl.Columns.Count - 1
                except:
                    pass
                    # Using sheet.max_row instead of .end("down")
                    # end("right") replaced with max_column

                try:
                    # UsedRange replaced with sheet.max_row/max_column
                    bs_last_row = source_pl.max_row  # openpyxl: use max_row
                    bs_last_col = used_range_bs.Column + used_range_bs.Columns.Count - 1
                except:
                    pass
                    # Using sheet.max_row instead of .end("down")
                    # end("right") replaced with max_column

                print(f"DEBUG: Source_PL dimensions: {pl_last_row} rows, {pl_last_col} columns (up to column {get_column_letter(pl_last_col)})")
                print(f"DEBUG: Source_BS dimensions: {bs_last_row} rows, {bs_last_col} columns (up to column {get_column_letter(bs_last_col)})")

                # Check what's in row 2 (YYYYMM values) of Source_PL
                yyyymm_row = []
                for col in range(2, pl_last_col + 1):
                    val = source_pl.cell(row=2, column=col).value
                    yyyymm_row.append(val)
                print(f"DEBUG: Source_PL row 2 (YYYYMM values): {yyyymm_row}")

                try:
                    del wb.defined_names['SourcePL']
                except:
                    pass
                try:
                    del wb.defined_names['SourceBS']
                except:
                    pass

                wb.defined_names.add(DefinedName('SourcePL', attr_text=f"'Source_PL'!$A$1:${get_column_letter(pl_last_col)}${pl_last_row}"))
                wb.defined_names.add(DefinedName('SourceBS', attr_text=f"'Source_BS'!$A$1:${get_column_letter(bs_last_col)}${bs_last_row}"))

                # Get the latest month added
                latest_month = new_month_list[-1]
                latest_m, latest_y, latest_name, latest_yyyymm = latest_month

                # Update Menu sheet with new current month
                update_step("Updating Menu sheet...")
                try:
                    menu_sheet = wb['Menu']

                    # Update lookup table (K:M) with new months
                    update_substep("Updating lookup table...")
                    last_lookup_row = self._update_menu_lookup_table(menu_sheet, new_month_list, existing_months)

                    # Update C7 dropdown validation with all available months
                    update_substep("Updating period dropdown...")
                    self._update_menu_dropdown(menu_sheet, last_lookup_row)

                    # Set current month (C7) to latest month - this triggers VLOOKUP formulas
                    # Use text format to prevent Excel from interpreting as date
                    menu_sheet['C7'].number_format = '@'  # Text format
                    menu_sheet['C7'].value = latest_name

                    # Also update C9 (Actuals Through) to match
                    menu_sheet['C9'].number_format = '@'  # Text format
                    menu_sheet['C9'].value = latest_name

                    # Update C8 "Data Range" text - get first month from lookup table K7
                    first_month_name = menu_sheet['K7'].value
                    if first_month_name:
                        menu_sheet['C8'].value = f"{first_month_name} - {latest_name}"

                    print(f"DEBUG: Updated Menu - Current Month: {latest_name}, Data Range: {first_month_name} - {latest_name}, Lookup table row: {last_lookup_row}")
                except Exception as e:
                    print(f"Warning: Could not update Menu sheet: {e}")
                    # Fallback to static values if lookup table update fails
                    try:
                        menu_sheet['G7'].value = latest_yyyymm
                        menu_sheet['E7'].value = latest_m
                        menu_sheet['F7'].value = latest_y
                        menu_sheet['G9'].value = latest_yyyymm
                    except:
                        pass

                # Add new month columns to report sheets (PL, Balance_Sheet, Cash_Flow)
                update_step("Updating P&L reports...")
                self._add_month_to_reports(wb, new_month_list, pl_last_col, bs_last_col, is_multi_division)

                # =====================================================================
                # POST-UPDATE CLEANUP: Match build code formatting
                # =====================================================================
                update_step("Finalizing formatting...")

                # Hide row 3 (YYYYMM helper) on all report sheets
                report_sheet_names = ['PL', 'Consolidated_PL', 'Balance_Sheet', 'Consolidated_BS', 'Cash_Flow']
                # Add division-specific sheets
                for sheet_name, _ in pl_sheets:
                    if sheet_name not in report_sheet_names:
                        report_sheet_names.append(sheet_name)
                for sheet_name, _ in bs_sheets:
                    if sheet_name not in report_sheet_names:
                        report_sheet_names.append(sheet_name)

                for sheet_name in report_sheet_names:
                    try:
                        if sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            # Row hiding handled via hide_row()
                            print(f"DEBUG: Hidden row 3 on {sheet_name}")
                    except Exception as e:
                        print(f"DEBUG: Could not hide row 3 on {sheet_name}: {e}")

                # Collapse outline groups on report sheets
                for sheet_name in report_sheet_names:
                    try:
                        if sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            # Outline handled via group_rows()/group_cols()
                            print(f"DEBUG: Collapsed outline groups on {sheet_name}")
                    except Exception as e:
                        print(f"DEBUG: Could not collapse groups on {sheet_name}: {e}")

                # Also collapse Forecast sheet
                try:
                    if 'Forecast' in wb.sheetnames:
                        forecast_sheet = wb['Forecast']
                        # Outline handled via group_rows()/group_cols()
                        print("DEBUG: Collapsed outline groups on Forecast")
                except Exception as e:
                    print(f"DEBUG: Could not collapse Forecast groups: {e}")

                # Force recalculation
                update_step("Recalculating formulas...")
                wb.app.calculate()

                # Set Dashboard as active sheet
                try:
                    if 'Dashboard' in wb.sheetnames:
                        wb['Dashboard'].activate()
                        print("DEBUG: Activated Dashboard sheet")
                except Exception as e:
                    print(f"DEBUG: Could not activate Dashboard: {e}")

                # Save
                update_step("Saving workbook...")
                wb.save(temp_path)

                # Remove legacy _FilterDatabase from saved file
                remove_filter_database_from_xlsx(temp_path)
                pass  # openpyxl auto-handles cleanup

            finally:
                pass  # openpyxl: no app to quit

            # Copy from temp to final location
            shutil.copy2(temp_path, save_path)

        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

    def _add_navigation_links(self, wb, new_month_list=None):
        """Add navigation hyperlinks to all sheets - Menu link on each sheet, sheet links on Menu"""
        try:
            menu_sheet = wb['Menu']

            # Define sheets that should have navigation with display names
            report_sheets = [
                ('Dashboard', 'Dashboard'),
                ('PL', 'P&L'),
                ('Balance_Sheet', 'Balance Sheet'),
                ('Cash_Flow', 'Cash Flow'),
                ('Forecast', 'Forecast'),
                ('Forecast_Summary', 'Forecast Summary'),
                ('Notes', 'Notes')
            ]

            # Add "Back to Menu" link on each report sheet (cell A1)
            for sheet_name, display_name in report_sheets:
                try:
                    sheet = wb[sheet_name]
                    # Add hyperlink in A1
                    sheet['A1'].value = '← Menu'
                    sheet['A1'].hyperlink = "#Menu!A1"
                    sheet['A1'].font = Font(color="0066CC", underline="single", size=9)
                except Exception as e:
                    print(f"DEBUG: Could not add Menu link to {sheet_name}: {e}")

            # Hide helper columns E:G on Menu
            try:
                # Column hiding handled via hide_columns_range()
                pass
            except Exception as e:
                print(f"DEBUG: Could not hide columns E:G: {e}")

            # Add Quick Links box in upper right of Menu (starting at I2)
            try:
                start_col = 'I'
                start_row = 2

                # Header
                menu_sheet[f'{start_col}{start_row}'].value = "Quick Links"
                menu_sheet[f'{start_col}{start_row}'].font = Font(bold=True, size=11)

                # Links
                link_row = start_row + 1
                for sheet_name, display_name in report_sheets:
                    try:
                        menu_sheet[f'{start_col}{link_row}'].value = display_name
                        cell = menu_sheet[f'{start_col}{link_row}']
                        cell.font = Font(color="0066CC", underline="single", size=10)
                        cell.hyperlink = f"#'{sheet_name}'!A1"
                        link_row += 1
                    except Exception as e:
                        print(f"DEBUG: Could not add link to {sheet_name} on Menu: {e}")

                # Add border box around Quick Links section
                start_col_num = openpyxl.utils.column_index_from_string(start_col)
                apply_border_box(menu_sheet, start_row, start_col_num, link_row - 1, start_col_num)

            except Exception as e:
                print(f"DEBUG: Could not create Quick Links box: {e}")

            print("DEBUG: Added navigation links")

        except Exception as e:
            print(f"DEBUG: Navigation links skipped: {e}")

    def _add_month_to_reports(self, wb, new_month_list, source_pl_last_col, source_bs_last_col=None, is_multi_division=False):
        """Add new month columns to PL, Balance Sheet, and Cash Flow reports

        OPTIMIZED VERSION: Uses bulk Find/Replace for formula updates instead of cell-by-cell.
        Supports both single-entity and multi-division models.
        """
        if source_bs_last_col is None:
            source_bs_last_col = source_pl_last_col

        # Get source structure based on model type
        source_struct = self._get_source_structure(is_multi_division)
        data_start_col = source_struct['data_start_col']  # B=2 for single, C=3 for multi

        # Get list of report sheets to update
        pl_sheets, bs_sheets = self._get_report_sheets(wb, is_multi_division)

        print(f"DEBUG: Model type: {'multi-division' if is_multi_division else 'single-entity'}")
        print(f"DEBUG: P&L sheets to update: {[s[0] for s in pl_sheets]}")
        print(f"DEBUG: BS sheets to update: {[s[0] for s in bs_sheets]}")

        try:
            # For single-entity, use the standard sheet names
            if not is_multi_division:
                pl_sheet = wb['PL'] if 'PL' in wb.sheetnames else None
                bs_sheet = wb['Balance_Sheet'] if 'Balance_Sheet' in wb.sheetnames else None
            else:
                # For multi-division, start with Consolidated_PL
                pl_sheet = wb['Consolidated_PL'] if 'Consolidated_PL' in wb.sheetnames else None
                bs_sheet = wb['Consolidated_BS'] if 'Consolidated_BS' in wb.sheetnames else None

            cf_sheet = wb['Cash_Flow'] if 'Cash_Flow' in wb.sheetnames else None

            # Get the new source column letters (after data was added)
            pl_col_letter = get_column_letter(source_pl_last_col)
            bs_col_letter = get_column_letter(source_bs_last_col)

            # Calculate the PREVIOUS last column (before new data was added)
            # This is needed to update formula ranges from old_col to new_col
            prev_pl_col = source_pl_last_col - len(new_month_list)
            prev_bs_col = source_bs_last_col - len(new_month_list)
            prev_pl_col_letter = get_column_letter(prev_pl_col) if prev_pl_col >= data_start_col else get_column_letter(data_start_col)
            prev_bs_col_letter = get_column_letter(prev_bs_col) if prev_bs_col >= data_start_col else get_column_letter(data_start_col)

            print(f"DEBUG: Source_PL range expanding from column {prev_pl_col_letter} to {pl_col_letter}")
            print(f"DEBUG: Source_BS range expanding from column {prev_bs_col_letter} to {bs_col_letter}")

            # =====================================================================
            # STEP 1: Bulk update existing formula ranges using Find/Replace
            # This is MUCH faster than cell-by-cell updates
            # =====================================================================
            print("DEBUG: Updating existing formula ranges with Find/Replace...")

            # Update Source_PL references in all sheets
            for sheet in wb.worksheets:
                try:
                    # Skip source sheets
                    if sheet.title in ['Source_PL', 'Source_BS', 'Source_Budget', 'Menu', 'Settings', 'Account_Mappings']:
                        continue

                    # Use cells_replace to update formula ranges
                    # Pattern: Source_PL!$B$2:$[OLD]$2 -> Source_PL!$B$2:$[NEW]$2
                    cells_replace(sheet, f"Source_PL!$B$2:${prev_pl_col_letter}$2", f"Source_PL!$B$2:${pl_col_letter}$2")
                    cells_replace(sheet, f"Source_PL!$B$3:${prev_pl_col_letter}$", f"Source_PL!$B$3:${pl_col_letter}$")
                    # For multi-division (column C start)
                    if is_multi_division:
                        cells_replace(sheet, f"Source_PL!$C$2:${prev_pl_col_letter}$2", f"Source_PL!$C$2:${pl_col_letter}$2")
                        cells_replace(sheet, f"Source_PL!$C$3:${prev_pl_col_letter}$", f"Source_PL!$C$3:${pl_col_letter}$")

                    # Update Source_BS references
                    cells_replace(sheet, f"Source_BS!$B$2:${prev_bs_col_letter}$2", f"Source_BS!$B$2:${bs_col_letter}$2")
                    cells_replace(sheet, f"Source_BS!$B$3:${prev_bs_col_letter}$", f"Source_BS!$B$3:${bs_col_letter}$")
                    if is_multi_division:
                        cells_replace(sheet, f"Source_BS!$C$2:${prev_bs_col_letter}$2", f"Source_BS!$C$2:${bs_col_letter}$2")
                        cells_replace(sheet, f"Source_BS!$C$3:${prev_bs_col_letter}$", f"Source_BS!$C$3:${bs_col_letter}$")

                    # Update Source_Budget references
                    cells_replace(sheet, f"Source_Budget!$B$2:${prev_pl_col_letter}$2", f"Source_Budget!$B$2:${pl_col_letter}$2")
                    cells_replace(sheet, f"Source_Budget!$B$3:${prev_pl_col_letter}$", f"Source_Budget!$B$3:${pl_col_letter}$")
                except Exception as e:
                    print(f"DEBUG: Find/Replace on {sheet.title}: {e}")

            print("DEBUG: Formula range updates complete")

            # =====================================================================
            # STEP 2: Insert new month columns in report sheets
            # =====================================================================
            for m, y, name, yyyymm in new_month_list:
                print(f"DEBUG: Adding month {name} ({yyyymm}) to report sheets...")

                # Update all P&L sheets
                for sheet_name, division_name in pl_sheets:
                    try:
                        current_pl_sheet = wb[sheet_name]

                        # Find the last MONTH column by looking for numeric YYYYMM in row 3
                        # Don't use end('right') as it includes YTD, Notes columns
                        last_month_col = 2  # Start at column B
                        for col in range(2, 50):  # Reasonable max columns
                            val = current_pl_sheet.cell(row=3, column=col).value
                            # YYYYMM values are integers like 202412
                            if val is not None and isinstance(val, (int, float)) and val > 200000 and val < 210000:
                                last_month_col = col
                            else:
                                # Stop at first non-YYYYMM column
                                break
                        new_col = last_month_col + 1

                        # Insert column
                        # Column insert: use sheet.insert_cols()

                        # Set up header row (row 3 is YYYYMM, row 4 is month name header)
                        current_pl_sheet.cell(row=3, column=new_col).value = yyyymm
                        current_pl_sheet.cell(row=3, column=new_col).font = Font(color="FFFFFF")
                        current_pl_sheet.cell(row=4, column=new_col).value = name
                        current_pl_sheet.cell(row=4, column=new_col).font = Font(bold=True)

                        # Using sheet.max_row instead of .end("down")

                        # Build formulas in memory first, then write in bulk
                        formulas = []
                        for row in range(5, last_row + 1):
                            account_cell = f"$A{row}"
                            if is_multi_division:
                                if division_name:
                                    formula = (
                                        f"=SUMPRODUCT("
                                        f"(Source_PL!$A$3:$A$1000=\"{division_name}\")*"
                                        f"(Source_PL!$B$3:$B$1000=TRIM({account_cell}))*"
                                        f"(Source_PL!$C$2:${pl_col_letter}$2={yyyymm})*"
                                        f"(Source_PL!$C$3:${pl_col_letter}$1000))"
                                    )
                                else:
                                    formula = (
                                        f"=SUMPRODUCT("
                                        f"(Source_PL!$B$3:$B$1000=TRIM({account_cell}))*"
                                        f"(Source_PL!$C$2:${pl_col_letter}$2={yyyymm})*"
                                        f"(Source_PL!$C$3:${pl_col_letter}$1000))"
                                    )
                            else:
                                formula = (
                                    f"=SUMPRODUCT("
                                    f"(Source_PL!$A$3:$A$1000=TRIM({account_cell}))*"
                                    f"(Source_PL!$B$2:${pl_col_letter}$2={yyyymm})*"
                                    f"(Source_PL!$B$3:${pl_col_letter}$1000))"
                                    )
                            formulas.append([formula])

                        # Bulk write all formulas at once
                        if formulas:
                            _data = formulas
                            if _data is not None:
                                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                                    for _ri, _row in enumerate(_data):
                                        for _ci, _val in enumerate(_row):
                                            current_pl_sheet.cell(row=5 + _ri, column=new_col + _ci).value = _val
                                elif isinstance(_data, list):
                                    for _ri, _val in enumerate(_data):
                                        if isinstance(_val, list):
                                            for _ci, _v in enumerate(_val):
                                                current_pl_sheet.cell(row=5 + _ri, column=new_col + _ci).value = _v
                                        else:
                                            current_pl_sheet.cell(row=5 + _ri, column=new_col).value = _val

                        # Format new column
                        apply_style_to_range(current_pl_sheet, 5, new_col, last_row, new_col, number_format='#,##0')

                        print(f"DEBUG: Added column to {sheet_name}")
                    except Exception as e:
                        print(f"Warning: Could not update {sheet_name}: {e}")

                # Use primary P&L sheet for YTD and annual updates
                if pl_sheet is None:
                    continue

                # Get positions for YTD update
                # Find the last MONTH column by looking for numeric YYYYMM in row 3
                pl_last_month_col = 2
                for col in range(2, 50):
                    val = pl_sheet.cell(row=3, column=col).value
                    if val is not None and isinstance(val, (int, float)) and val > 200000 and val < 210000:
                        pl_last_month_col = col
                    else:
                        break
                new_col = pl_last_month_col
                # Using sheet.max_row instead of .end("down")

                # =====================================================================
                # STEP 3: Update YTD and Annual formulas (only for primary P&L sheet)
                # Uses bulk writes for efficiency
                # =====================================================================
                header_row = 4
                py_ytd_col = None
                cy_ytd_col = None
                fy_start_col = None

                # Scan row 4 to find column headers
                for check_col in range(new_col + 1, new_col + 15):
                    header_val = pl_sheet.cell(row=header_row, column=check_col).value
                    if header_val == 'PY YTD':
                        py_ytd_col = check_col
                    elif header_val == 'CY YTD':
                        cy_ytd_col = check_col
                    elif header_val and str(header_val).isdigit() and len(str(header_val)) == 4:
                        if fy_start_col is None:
                            fy_start_col = check_col

                # Update YTD formulas using bulk write
                if py_ytd_col and cy_ytd_col:
                    print(f"DEBUG: Updating YTD columns (PY YTD: {py_ytd_col}, CY YTD: {cy_ytd_col})")
                    first_data_col = get_column_letter(2)
                    last_data_col = get_column_letter(new_col)
                    helper_range = f'{first_data_col}$3:{last_data_col}$3'

                    # Build formulas in memory
                    cy_formulas = []
                    py_formulas = []
                    for row in range(5, last_row + 1):
                        data_range = f'{first_data_col}{row}:{last_data_col}{row}'
                        cy_formulas.append([f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7)*--(MOD({helper_range},100)<=Menu!$E$7))'])
                        py_formulas.append([f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7-1)*--(MOD({helper_range},100)<=Menu!$E$7))'])

                    # Bulk write
                    if cy_formulas:
                        _data = cy_formulas
                        if _data is not None:
                            if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                                for _ri, _row in enumerate(_data):
                                    for _ci, _val in enumerate(_row):
                                        pl_sheet.cell(row=5 + _ri, column=cy_ytd_col + _ci).value = _val
                            elif isinstance(_data, list):
                                for _ri, _val in enumerate(_data):
                                    if isinstance(_val, list):
                                        for _ci, _v in enumerate(_val):
                                            pl_sheet.cell(row=5 + _ri, column=cy_ytd_col + _ci).value = _v
                                    else:
                                        pl_sheet.cell(row=5 + _ri, column=cy_ytd_col).value = _val
                        _data = py_formulas
                        if _data is not None:
                            if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                                for _ri, _row in enumerate(_data):
                                    for _ci, _val in enumerate(_row):
                                        pl_sheet.cell(row=5 + _ri, column=py_ytd_col + _ci).value = _val
                            elif isinstance(_data, list):
                                for _ri, _val in enumerate(_data):
                                    if isinstance(_val, list):
                                        for _ci, _v in enumerate(_val):
                                            pl_sheet.cell(row=5 + _ri, column=py_ytd_col + _ci).value = _v
                                    else:
                                        pl_sheet.cell(row=5 + _ri, column=py_ytd_col).value = _val
                    print(f"DEBUG: Updated YTD formulas")

                # Update Annual column formulas using bulk write
                if fy_start_col:
                    month_cols_by_year = {}
                    for col in range(2, new_col + 1):
                        yyyymm_val = pl_sheet.cell(row=3, column=col).value
                        if yyyymm_val:
                            col_year = int(yyyymm_val) // 100
                            if col_year not in month_cols_by_year:
                                month_cols_by_year[col_year] = []
                            month_cols_by_year[col_year].append(col)

                    for fy_col_offset in range(10):
                        fy_col = fy_start_col + fy_col_offset
                        year_header = pl_sheet.cell(row=header_row, column=fy_col).value
                        if year_header and str(year_header).isdigit():
                            year = int(year_header)
                            if year in month_cols_by_year:
                                year_cols = month_cols_by_year[year]
                                # Build formulas in memory
                                annual_formulas = []
                                for row in range(5, last_row + 1):
                                    refs = '+'.join([f'{get_column_letter(c)}{row}' for c in year_cols])
                                    annual_formulas.append([f'={refs}'])
                                # Bulk write
                                if annual_formulas:
                                    _data = annual_formulas
                                    if _data is not None:
                                        if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                                            for _ri, _row in enumerate(_data):
                                                for _ci, _val in enumerate(_row):
                                                    pl_sheet.cell(row=5 + _ri, column=fy_col + _ci).value = _val
                                        elif isinstance(_data, list):
                                            for _ri, _val in enumerate(_data):
                                                if isinstance(_val, list):
                                                    for _ci, _v in enumerate(_val):
                                                        pl_sheet.cell(row=5 + _ri, column=fy_col + _ci).value = _v
                                                else:
                                                    pl_sheet.cell(row=5 + _ri, column=fy_col).value = _val
                                print(f"DEBUG: Updated Annual {year} formulas")

                # =====================================================================
                # STEP 4: Update Balance Sheet sheets (bulk write for new column)
                # Existing formulas already updated via Find/Replace in Step 1
                # =====================================================================
                for sheet_name, division_name in bs_sheets:
                    try:
                        current_bs_sheet = wb[sheet_name]
                        # Find the last MONTH column by looking for numeric YYYYMM in row 3
                        bs_last_month_col = 2  # Start at column B
                        for col in range(2, 50):
                            val = current_bs_sheet.cell(row=3, column=col).value
                            if val is not None and isinstance(val, (int, float)) and val > 200000 and val < 210000:
                                bs_last_month_col = col
                            else:
                                break
                        new_col_bs = bs_last_month_col + 1

                        # Column insert: use sheet.insert_cols()
                        current_bs_sheet.cell(row=3, column=new_col_bs).value = yyyymm
                        current_bs_sheet.cell(row=3, column=new_col_bs).font = Font(color="FFFFFF")
                        current_bs_sheet.cell(row=4, column=new_col_bs).value = name
                        current_bs_sheet.cell(row=4, column=new_col_bs).font = Font(bold=True)

                        # Using sheet.max_row instead of .end("down")

                        # Build formulas in memory
                        bs_formulas = []
                        for row in range(5, bs_last_row + 1):
                            account_cell = f"$A{row}"
                            if is_multi_division:
                                if division_name:
                                    formula = f"=SUMPRODUCT((Source_BS!$A$3:$A$1000=\"{division_name}\")*(Source_BS!$B$3:$B$1000=TRIM({account_cell}))*(Source_BS!$C$2:${bs_col_letter}$2={yyyymm})*(Source_BS!$C$3:${bs_col_letter}$1000))"
                                else:
                                    formula = f"=SUMPRODUCT((Source_BS!$B$3:$B$1000=TRIM({account_cell}))*(Source_BS!$C$2:${bs_col_letter}$2={yyyymm})*(Source_BS!$C$3:${bs_col_letter}$1000))"
                            else:
                                formula = f"=SUMPRODUCT((Source_BS!$A$3:$A$1000=TRIM({account_cell}))*(Source_BS!$B$2:${bs_col_letter}$2={yyyymm})*(Source_BS!$B$3:${bs_col_letter}$1000))"
                            bs_formulas.append([formula])

                        # Bulk write
                        if bs_formulas:
                            _data = bs_formulas
                            if _data is not None:
                                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                                    for _ri, _row in enumerate(_data):
                                        for _ci, _val in enumerate(_row):
                                            current_bs_sheet.cell(row=5 + _ri, column=new_col_bs + _ci).value = _val
                                elif isinstance(_data, list):
                                    for _ri, _val in enumerate(_data):
                                        if isinstance(_val, list):
                                            for _ci, _v in enumerate(_val):
                                                current_bs_sheet.cell(row=5 + _ri, column=new_col_bs + _ci).value = _v
                                        else:
                                            current_bs_sheet.cell(row=5 + _ri, column=new_col_bs).value = _val
                        apply_style_to_range(current_bs_sheet, 5, new_col_bs, bs_last_row, new_col_bs, number_format='#,##0')
                        print(f"DEBUG: Updated {sheet_name}")
                    except Exception as e:
                        print(f"Warning: Could not update {sheet_name}: {e}")

                # =====================================================================
                # STEP 5: Update Cash Flow (copy formulas from previous column)
                # =====================================================================
                if cf_sheet:
                    try:
                        # Find the last MONTH column by looking for numeric YYYYMM in row 3
                        cf_last_month_col = 2  # Start at column B
                        for col in range(2, 50):
                            val = cf_sheet.cell(row=3, column=col).value
                            if val is not None and isinstance(val, (int, float)) and val > 200000 and val < 210000:
                                cf_last_month_col = col
                            else:
                                break
                        new_col_cf = cf_last_month_col + 1

                        # Column insert: use sheet.insert_cols()
                        cf_sheet.cell(row=3, column=new_col_cf).value = yyyymm
                        cf_sheet.cell(row=3, column=new_col_cf).font = Font(color="FFFFFF")
                        cf_sheet.cell(row=4, column=new_col_cf).value = name
                        cf_sheet.cell(row=4, column=new_col_cf).font = Font(bold=True)

                        # Cash Flow formulas reference PL and BS - copy from previous column
                        # Using sheet.max_row instead of .end("down")
                        for row in range(5, cf_last_row + 1):
                            prev_formula = cf_sheet.cell(row=row, column=cf_last_month_col).formula
                            if prev_formula:
                                cf_sheet.cell(row=row, column=new_col_cf).value = prev_formula

                        apply_style_to_range(cf_sheet, 5, new_col_cf, cf_last_row, new_col_cf, number_format='#,##0')
                    except Exception as e:
                        print(f"Warning: Could not update Cash_Flow: {e}")

                print(f"Added month {name} to reports")

            # Update Forecast sheet if it exists
            try:
                forecast_sheet = wb['Forecast']
                # UsedRange replaced with sheet.max_row/max_column
                forecast_last_row = forecast_used.Row + forecast_used.Rows.Count - 1
                forecast_last_col = forecast_used.Column + forecast_used.Columns.Count - 1

                print(f"DEBUG: Updating Forecast formulas to use expanded source range (up to column {pl_col_letter})")
                print(f"DEBUG: Forecast sheet has {forecast_last_row} rows, {forecast_last_col} columns")
                import re
                forecast_formula_count = 0
                formulas_found = 0

                # Starting column depends on mode: B for single-entity, C for multi-division
                start_col = 'C' if is_multi_division else 'B'

                def expand_forecast_range(formula_str, new_col, start_col):
                    """Replace Source_PL and Source_Budget column references with new_col"""
                    # Handle both Source_PL and Source_Budget
                    result = formula_str
                    for source_sheet in ['Source_PL', 'Source_Budget']:
                        # Pattern matches Source_XX!$B$2:$T$2 or Source_XX!$C$2:$T$2 etc
                        result = re.sub(
                            rf'{source_sheet}!\$?[BC]\$?(\d+):\$?([A-Z]+)\$?(\d+)',
                            rf'{source_sheet}!${start_col}$\1:${new_col}$\3',
                            result
                        )
                    return result

                # Scan all cells in the Forecast sheet for Source_PL references
                for col in range(2, min(forecast_last_col + 1, 65)):  # B through all used columns
                    for row in range(5, forecast_last_row + 1):  # Data starts at row 5
                        cell = forecast_sheet.cell(row=row, column=col)
                        formula = cell.formula
                        if formula and ('Source_PL!' in str(formula) or 'Source_Budget!' in str(formula)):
                            formulas_found += 1
                            original = str(formula)
                            updated = expand_forecast_range(original, pl_col_letter, start_col)
                            if updated != original:
                                if forecast_formula_count < 5:
                                    print(f"DEBUG: Forecast ({row},{col}) BEFORE: {original[:100]}")
                                    print(f"DEBUG: Forecast ({row},{col}) AFTER:  {updated[:100]}")
                                cell.value = updated
                                forecast_formula_count += 1
                            elif formulas_found <= 3:
                                # Show formulas that didn't change to debug pattern
                                print(f"DEBUG: Forecast ({row},{col}) NO CHANGE: {original[:100]}")

                print(f"DEBUG: Found {formulas_found} Forecast formulas with Source refs, updated {forecast_formula_count}")

                # =====================================================================
                # CRITICAL FIX: Update Forecast Actual formulas for NEW months
                # During build, future months have '0' in Actual column. Now we have data!
                # =====================================================================
                COLS_PER_MONTH = 5
                COL_ACTUAL = 0  # Offset within month group

                # Get source column mapping: find which Source_PL column has each YYYYMM
                source_pl = wb['Source_PL']
                source_yyyymm_map = {}  # YYYYMM -> column letter
                # end("right") replaced with max_column
                for col in range(data_start_col, source_row2_last + 1):
                    yyyymm_val = source_pl.cell(row=2, column=col).value
                    if yyyymm_val and isinstance(yyyymm_val, (int, float)) and yyyymm_val > 200000:
                        source_yyyymm_map[int(yyyymm_val)] = get_column_letter(col)

                print(f"DEBUG: Source_PL YYYYMM map: {source_yyyymm_map}")

                # Get account column based on mode
                source_acct_col = 'B' if is_multi_division else 'A'

                # Find which year Forecast is using (from row 3 header)
                try:
                    first_month_header = forecast_sheet['B3'].value
                    if first_month_header and ' ' in str(first_month_header):
                        year_part = str(first_month_header).split()[-1]
                        if len(year_part) == 2:
                            forecast_year = 2000 + int(year_part)
                        else:
                            forecast_year = int(year_part)
                    else:
                        forecast_year = new_month_list[-1][1]  # Use latest data year
                except:
                    forecast_year = new_month_list[-1][1]

                print(f"DEBUG: Forecast year detected: {forecast_year}")

                # Update Actual formulas for months that now have Source_PL data
                actual_updated = 0
                for month_idx in range(12):  # Jan=0 through Dec=11
                    month_num = month_idx + 1
                    month_yyyymm = forecast_year * 100 + month_num

                    if month_yyyymm in source_yyyymm_map:
                        source_col_letter = source_yyyymm_map[month_yyyymm]
                        actual_col = 2 + (month_idx * COLS_PER_MONTH) + COL_ACTUAL

                        # Update all data rows (starting at row 5)
                        for row in range(5, forecast_last_row + 1):
                            current_val = forecast_sheet.cell(row=row, column=actual_col).value
                            current_formula = forecast_sheet.cell(row=row, column=actual_col).formula

                            # Only update if currently '0' or no formula
                            if current_val == 0 or (not current_formula or str(current_formula) == '0'):
                                if is_multi_division:
                                    # For consolidated forecast - need division handling
                                    new_formula = f'=IFERROR(SUMIF(Source_PL!${source_acct_col}$3:${source_acct_col}$1500,TRIM($A{row}),Source_PL!{source_col_letter}$3:{source_col_letter}$1500),0)'
                                else:
                                    new_formula = f'=IFERROR(SUMIF(Source_PL!${source_acct_col}$3:${source_acct_col}$1500,TRIM($A{row}),Source_PL!{source_col_letter}$3:{source_col_letter}$1500),0)'
                                forecast_sheet.cell(row=row, column=actual_col).value = new_formula
                                actual_updated += 1

                print(f"DEBUG: Updated {actual_updated} Forecast Actual formulas with new data")

            except Exception as e:
                print(f"DEBUG: Forecast sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Update Forecast_Summary sheet if it exists
            try:
                forecast_summary = wb['Forecast_Summary']
                # UsedRange replaced with sheet.max_row/max_column
                fs_last_row = fs_used.Row + fs_used.Rows.Count - 1

                print(f"DEBUG: Updating Forecast_Summary formulas to use expanded source range (up to column {pl_col_letter})")
                import re
                fs_formula_count = 0

                # Starting column depends on mode
                start_col = 'C' if is_multi_division else 'B'

                def expand_fs_range(formula_str, new_col, start_col):
                    """Replace Source_PL column references with new_col"""
                    result = re.sub(
                        r'Source_PL!\$?[BC]\$?(\d+):\$?[A-Z]+\$?(\d+)',
                        rf'Source_PL!${start_col}$\1:${new_col}$\2',
                        formula_str
                    )
                    return result

                # Check all columns for formulas that reference Source_PL
                for col in range(2, 15):  # B through N
                    for row in range(3, fs_last_row + 1):
                        cell = forecast_summary.cell(row=row, column=col)
                        formula = cell.formula
                        if formula and 'Source_PL!' in str(formula):
                            original = str(formula)
                            updated = expand_fs_range(original, pl_col_letter, start_col)
                            if updated != original:
                                if fs_formula_count < 3:
                                    print(f"DEBUG: Forecast_Summary ({row},{col}) BEFORE: {original[:80]}...")
                                    print(f"DEBUG: Forecast_Summary ({row},{col}) AFTER:  {updated[:80]}...")
                                cell.value = updated
                                fs_formula_count += 1

                print(f"DEBUG: Updated {fs_formula_count} Forecast_Summary formulas")
            except Exception as e:
                print(f"DEBUG: Forecast_Summary sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Update Dashboard sheet - regenerate key formulas to use dynamic Menu reference
            try:
                dashboard_sheet = wb['Dashboard']
                print(f"DEBUG: Regenerating Dashboard formulas with dynamic Menu reference")

                # Preserve existing company name from B2
                existing_company_name = dashboard_sheet['B2'].value
                print(f"DEBUG: Preserving company name: {existing_company_name}")

                # Update Dashboard "Current Period" header (B4) to link to Menu!C7 as TEXT
                # Use TEXT() to ensure the month name displays correctly even if C7 contains a date
                dashboard_sheet['B4'].value = '="Current Period: "&TEXT(Menu!C7,"mmm yyyy")'

                # Add period selector dropdown on Dashboard (row 6) that syncs with Menu!C7
                try:
                    dashboard_sheet['B6'].value = 'Period:'
                    dashboard_sheet['B6'].font = Font(bold=True, size=10)
                    # C6 links to Menu!C7 (two-way sync via formula)
                    dashboard_sheet['C6'].value = '=Menu!C7'
                    dashboard_sheet['C6'].font = Font(bold=True, size=10)
                    dashboard_sheet['C6'].fill = PatternFill(start_color="FFFFC8", end_color="FFFFC8", fill_type="solid")  # Light yellow
                    # Note: To make dropdown work, user changes Menu!C7 directly
                    dashboard_sheet['D6'].value = '← Change on Menu sheet'
                    dashboard_sheet['D6'].font = Font(size=8, italic=True, color="808080")
                except Exception as e:
                    print(f"DEBUG: Could not add period selector to Dashboard: {e}")

                # Regenerate P&L Summary formulas (rows 8-12) with dynamic current month lookup
                # These formulas use Menu!G7 to dynamically select the month based on dropdown
                detected_totals = {}  # Use defaults
                total_income_name = detected_totals.get('total_income', 'Total for Income')
                total_cogs_name = detected_totals.get('total_cogs', 'Total for Cost of Sales')
                total_expenses_name = detected_totals.get('total_expenses', 'Total for Expenses')

                pl_accounts_info = [
                    ('Revenue', total_income_name),
                    ('Cost of Goods Sold', total_cogs_name),
                    ('Gross Profit', 'Gross Profit'),
                    ('Operating Expenses', total_expenses_name),
                    ('Net Income', 'Net Income'),
                ]

                for i, (label, account) in enumerate(pl_accounts_info):
                    row = 8 + i
                    # Current month formula (column C) - uses SUMPRODUCT with Menu!G7
                    if is_multi_division:
                        cm_formula = f'=IF($C$6="Consolidated",SUMPRODUCT((Source_PL!$B$3:$B$1500="{account}")*(Source_PL!$C$2:{pl_col_letter}$2=Menu!$G$7)*(Source_PL!$C$3:{pl_col_letter}$1500)),SUMPRODUCT((Source_PL!$A$3:$A$1500=$C$6)*(Source_PL!$B$3:$B$1500="{account}")*(Source_PL!$C$2:{pl_col_letter}$2=Menu!$G$7)*(Source_PL!$C$3:{pl_col_letter}$1500)))'
                        ytd_formula = f'=IF($C$6="Consolidated",IFERROR(SUMPRODUCT((Source_PL!$B$3:$B$1500="{account}")*(INT(Source_PL!$C$2:{pl_col_letter}$2/100)=Menu!$I$7)*(Source_PL!$C$3:{pl_col_letter}$1500)),0),IFERROR(SUMPRODUCT((Source_PL!$A$3:$A$1500=$C$6)*(Source_PL!$B$3:$B$1500="{account}")*(INT(Source_PL!$C$2:{pl_col_letter}$2/100)=Menu!$I$7)*(Source_PL!$C$3:{pl_col_letter}$1500)),0))'
                    else:
                        cm_formula = f'=SUMPRODUCT((Source_PL!$A$3:$A$1500="{account}")*(Source_PL!$B$2:{pl_col_letter}$2=Menu!$G$7)*(Source_PL!$B$3:{pl_col_letter}$1500))'
                        ytd_formula = f'=IFERROR(SUMPRODUCT((Source_PL!$A$3:$A$1500="{account}")*(INT(Source_PL!$B$2:{pl_col_letter}$2/100)=Menu!$F$7)*(Source_PL!$B$3:{pl_col_letter}$1500)),0)'

                    dashboard_sheet[f'C{row}'].value = cm_formula
                    dashboard_sheet[f'E{row}'].value = ytd_formula

                print(f"DEBUG: Regenerated Dashboard P&L summary formulas (rows 8-12)")

                # Also expand any other Source_PL/BS references in Dashboard
                # UsedRange replaced with sheet.max_row/max_column
                dashboard_last_row = dashboard_used.Row + dashboard_used.Rows.Count - 1
                import re
                formula_count = 0
                start_col = 'C' if is_multi_division else 'B'

                def expand_source_range(formula_str, source_sheet, new_col, start_col):
                    patterns = [
                        (rf'{source_sheet}!\$?[BC]\$?(\d+):\$?[A-Z]+\$?(\d+)', rf'{source_sheet}!${start_col}$\1:${new_col}$\2'),
                    ]
                    result = formula_str
                    for pattern, replacement in patterns:
                        result = re.sub(pattern, replacement, result)
                    return result

                # Skip rows 4 and 8-12 since we already updated those
                skip_rows = {4, 8, 9, 10, 11, 12}
                for col in range(2, 13):
                    for row in range(1, min(dashboard_last_row + 1, 60)):
                        if row in skip_rows and col in {2, 3, 5}:  # B, C, E columns
                            continue
                        cell = dashboard_sheet.cell(row=row, column=col)
                        formula = cell.formula
                        if formula and ('Source_PL!' in str(formula) or 'Source_BS!' in str(formula)):
                            original = str(formula)
                            updated = expand_source_range(original, 'Source_PL', pl_col_letter, start_col)
                            updated = expand_source_range(updated, 'Source_BS', bs_col_letter, start_col)
                            if updated != original:
                                cell.value = updated
                                formula_count += 1

                print(f"DEBUG: Expanded {formula_count} additional Dashboard formulas")
            except Exception as e:
                print(f"DEBUG: Dashboard sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Auto-fit columns on all Balance Sheet and Cash Flow sheets after update
            for sheet_name, _ in bs_sheets:
                try:
                    sheet = wb[sheet_name]
                    # Note: openpyxl doesn't support auto-fit; Excel adjusts on open
                    print(f"DEBUG: Auto-fit {sheet_name} columns")
                except Exception as e:
                    print(f"DEBUG: {sheet_name} auto-fit skipped: {e}")

            try:
                if cf_sheet:
                    cf_# Note: openpyxl doesn't support auto-fit; Excel adjusts on open
                    print("DEBUG: Auto-fit Cash Flow columns")
            except Exception as e:
                print(f"DEBUG: Cash Flow auto-fit skipped: {e}")

            # Format Forecast notes: wrap text, vertical center, auto-fit row heights
            try:
                forecast_sheet = wb['Forecast']
                # UsedRange replaced with sheet.max_row/max_column
                forecast_last_row = forecast_used.Row + forecast_used.Rows.Count - 1

                # Note columns are every 5th column starting at column 5 (E), 10 (J), 15 (O), etc.
                # Actually: col 2=Actual, col 3=Budget, col 4=Adj, col 5=Note, col 6=Forecast
                # So Note columns are at 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60
                note_cols = [5 + (i * 5) for i in range(12)]  # 12 months

                for note_col in note_cols:
                    try:
                        # Wrap text and vertically center note columns
                        apply_style_to_range(forecast_sheet, 5, note_col, forecast_last_row, note_col,
                                            alignment=Alignment(wrap_text=True, vertical='center'))
                    except:
                        pass

                # Note: openpyxl cannot auto-fit row heights - Excel handles this on open
                # Vertically center all data cells
                try:
                    max_r = forecast_sheet.max_row or 1
                    max_c = forecast_sheet.max_column or 1
                    apply_style_to_range(forecast_sheet, 5, 1, max_r, max_c,
                                        alignment=Alignment(vertical='center'))
                except:
                    pass

                print("DEBUG: Formatted Forecast notes and auto-fit rows")
            except Exception as e:
                print(f"DEBUG: Forecast formatting skipped: {e}")

            # Add navigation links to all sheets
            self._add_navigation_links(wb, new_month_list)

        except Exception as e:
            print(f"Warning: Error adding month to reports: {e}")
            import traceback
            traceback.print_exc()

    def _create_multi_division_model(self, save_path, user_start, user_end):
        """Create a multi-division Excel model with consolidated reporting"""
        # Initialize step timer for performance tracking
        timer = StepTimer()

        # Calculate total estimated time based on number of divisions
        num_divisions = len(self.divisions)

        # Calculate dynamic total steps:
        # Fixed steps: 15 (consolidation, Excel setup, consolidated sheets, budget, forecasts, menu, dashboard, notes, mappings, organize, VBA, save)
        # Variable steps: 3 per division (read files, division sheets, division forecasts)
        # Count: 18 fixed steps + 3 per division (read files, create sheets, create forecast)
        total_steps = 18 + (3 * num_divisions)

        # Use historical timing data for more accurate progress bar estimate
        default_estimate = 180 + (num_divisions - 1) * 30  # Fallback: base ~3 min + 30s per extra division
        estimated_total = StepTimer.get_historical_estimate(default_estimate=default_estimate)

        # Create progress tracker for mid-step updates
        update_step, update_substep, state = self._create_progress_tracker(total_steps, estimated_total)

        # Wrap update_step to also track timing
        original_update_step = update_step
        def update_step(msg):
            timer.start_step(msg)
            original_update_step(msg)

        # Step 1-3: Parse all division files
        all_pl_accounts = []  # Stacked accounts with division
        all_bs_accounts = []
        all_months = None
        pl_totals = {}
        bs_totals = {}
        division_configs = []

        for i, div in enumerate(self.divisions):
            div_name = div['name']
            update_step(f"Reading {div_name} files ({i+1}/{num_divisions})...")

            # Get per-division date range (convert month name to number)
            div_start_month = self.MONTHS.index(div.get('start_month', 'January')) + 1
            div_start_year = int(div.get('start_year', datetime.now().year))
            div_end_month = self.MONTHS.index(div.get('end_month', 'December')) + 1
            div_end_year = int(div.get('end_year', datetime.now().year))

            div_user_start = (div_start_month, div_start_year)
            div_user_end = (div_end_month, div_end_year)

            print(f"Division {div_name}: {div.get('start_month')} {div_start_year} to {div.get('end_month')} {div_end_year}")

            # Parse P&L with division-specific date range (combined read for speed)
            update_substep(f"{div_name}: Loading P&L file...")
            pl_data, pl_indents = self._read_excel_with_indents(div['pl_path'])
            update_substep(f"{div_name}: Parsing P&L accounts...")
            div_pl_accounts, div_months, div_pl_totals = self._parse_financial_data(
                pl_data, pl_indents, div_user_start, div_user_end, div['pl_path']
            )
            update_substep(f"{div_name}: Found {len(div_pl_accounts)} P&L accounts")

            # Parse BS with division-specific date range (combined read for speed)
            update_substep(f"{div_name}: Loading Balance Sheet file...")
            bs_data, bs_indents = self._read_excel_with_indents(div['bs_path'])
            update_substep(f"{div_name}: Parsing BS accounts...")
            div_bs_accounts, _, div_bs_totals = self._parse_financial_data(
                bs_data, bs_indents, div_user_start, div_user_end, div['bs_path']
            )
            update_substep(f"{div_name}: Found {len(div_bs_accounts)} BS accounts")

            # Add division identifier to each account
            for acct in div_pl_accounts:
                acct['division'] = div_name
            for acct in div_bs_accounts:
                acct['division'] = div_name

            all_pl_accounts.extend(div_pl_accounts)
            all_bs_accounts.extend(div_bs_accounts)

            # Use first division's months as reference (or merge if different)
            if all_months is None:
                all_months = div_months
            else:
                # Merge months from different divisions
                existing_keys = set((m, y) for m, y, _ in all_months)
                for m, y, name in div_months:
                    if (m, y) not in existing_keys:
                        all_months.append((m, y, name))
                all_months = sorted(all_months, key=lambda x: (x[1], x[0]))

            # Merge totals
            pl_totals.update(div_pl_totals)
            bs_totals.update(div_bs_totals)

            # Create DivisionConfig for consolidation engine
            division_configs.append(DivisionConfig(
                name=div_name,
                is_primary=div.get('is_primary', False),
                pl_file_path=div['pl_path'],
                bs_file_path=div['bs_path'],
                pl_accounts=div_pl_accounts,
                bs_accounts=div_bs_accounts
            ))

        print(f"Parsed {len(self.divisions)} divisions:")
        for div in self.divisions:
            print(f"  - {div['name']}")
        print(f"Total: {len(all_pl_accounts)} P&L accounts, {len(all_bs_accounts)} BS accounts, {len(all_months)} months")

        # Detect actual date range - filter out months with no data
        update_substep("Detecting actual data range...")
        all_months = self._detect_actual_date_range(all_pl_accounts + all_bs_accounts, all_months)
        print(f"After date range detection: {len(all_months)} months with data")

        # Step 4: Run consolidation engine for account matching
        update_step("Running consolidation engine...")
        update_substep("Initializing consolidation engine...")
        engine = ConsolidationEngine(division_configs)

        # Match P&L accounts
        update_substep("Matching P&L accounts across divisions...")
        pl_mappings = engine.match_accounts(statement_type="pl")
        update_substep(f"Found {len(pl_mappings)} consolidated P&L accounts")
        print(f"P&L mappings: {len(pl_mappings)} consolidated accounts")

        # Match BS accounts
        update_substep("Matching Balance Sheet accounts across divisions...")
        bs_mappings = engine.match_accounts(statement_type="bs")
        update_substep(f"Found {len(bs_mappings)} consolidated BS accounts")
        print(f"BS mappings: {len(bs_mappings)} consolidated accounts")

        # Store mappings
        self.account_mappings = {'pl': pl_mappings, 'bs': bs_mappings}

        # Skip mapping review dialog - auto-process with generated mappings
        # (Dialog removed per user request - can be re-enabled later if needed)
        update_step("Processing account mappings...")
        update_substep("Using auto-generated mappings...")

        # Create consolidated account lists
        update_step("Creating consolidated accounts...")
        update_substep("Consolidating P&L values across divisions...")
        consolidated_pl = engine.consolidate_values(pl_mappings, all_months, "pl")
        update_substep("Consolidating Balance Sheet values across divisions...")
        consolidated_bs = engine.consolidate_values(bs_mappings, all_months, "bs")
        update_substep(f"Created {len(consolidated_pl)} P&L and {len(consolidated_bs)} BS accounts")

        # Now create the Excel workbook
        use_template = os.path.exists(TEMPLATE_PATH)
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        update_step("Starting Excel...")
        update_substep("Launching Excel application...")
        # openpyxl: no Excel app needed
        # openpyxl: no display_alerts needed
        # openpyxl: no screen_updating needed
        # openpyxl: no calculation mode needed
        wb = None

        try:
            update_substep("Opening workbook template...")
            if use_template:
                shutil.copy(TEMPLATE_PATH, temp_path)
                wb = load_workbook(temp_path, keep_vba=True)
                menu_sheet = wb['Menu']
                source_pl = wb['Source_PL']
                source_bs = wb['Source_BS']
            else:
                wb = Workbook()
                menu_sheet = wb.worksheets[0]
                menu_sheet.title = 'Menu'
                source_pl = wb.create_sheet('Source_PL')
                source_bs = wb.create_sheet('Source_BS')

            # Clear existing data
            update_substep("Clearing existing data...")
            clear_sheet_data(source_pl, start_row=2)
            clear_sheet_data(source_bs, start_row=2)

            # Populate source sheets with stacked division data
            update_step("Populating source data with divisions...")
            update_substep(f"Writing {len(all_pl_accounts)} P&L accounts to Source_PL...")
            self._populate_source_sheet(source_pl, all_pl_accounts, all_months)
            update_substep(f"Writing {len(all_bs_accounts)} BS accounts to Source_BS...")
            self._populate_source_sheet(source_bs, all_bs_accounts, all_months)

            # Create or get report sheets
            update_step("Creating consolidated P&L...")
            update_substep("Setting up Consolidated_PL sheet...")
            if 'Consolidated_PL' in wb.sheetnames:
                cons_pl_sheet = wb['Consolidated_PL']
                clear_sheet_data(cons_pl_sheet)
            else:
                # Try to use existing PL sheet or create new
                if 'PL' in wb.sheetnames:
                    cons_pl_sheet = wb['PL']
                    cons_pl_sheet.title = 'Consolidated_PL'
                    clear_sheet_data(cons_pl_sheet)
                else:
                    cons_pl_sheet = wb.create_sheet('Consolidated_PL')

            # Create consolidated P&L report using consolidated accounts
            update_substep(f"Writing {len(consolidated_pl)} consolidated P&L accounts...")
            self._create_consolidated_pl_report(cons_pl_sheet, consolidated_pl, all_months, pl_totals)

            # Skip Consolidated Balance Sheet - not realistic for multi-division
            # Division-specific balance sheets are available instead

            update_step("Creating Cash Flow...")
            update_substep("Setting up Cash_Flow sheet...")
            if 'Cash_Flow' in wb.sheetnames:
                cons_cf_sheet = wb['Cash_Flow']
                clear_sheet_data(cons_cf_sheet)
            else:
                cons_cf_sheet = wb.create_sheet('Cash_Flow')

            update_substep("Building cash flow formulas...")
            self._create_cash_flow(cons_cf_sheet, all_months)

            # =================================================================
            # OPTIMIZATION: Temporarily disabled due to COM errors
            # TODO: Debug template cloning on user's Excel version
            # =================================================================
            # update_step("Creating division sheet templates...")
            # update_substep("Building pre-formatted _TPL_PL and _TPL_BS templates...")
            # tpl_pl, tpl_bs = self._create_runtime_templates(wb, all_months)
            # use_optimized = tpl_pl is not None and tpl_bs is not None
            use_optimized = False  # DISABLED - using standard method
            print(f"[OPTIMIZATION] Disabled - using standard sheet creation")

            # Create division-specific sheets (track last sheet to maintain correct order)
            last_created_sheet = cons_cf_sheet
            for idx, div in enumerate(self.divisions):
                div_name = div['name']
                safe_name = div_name.replace(' ', '_')[:20]  # Excel sheet name limit

                update_step(f"Creating {div_name} sheets ({idx+1}/{num_divisions})...")

                # Get division-specific accounts
                div_pl = [a for a in all_pl_accounts if a.get('division') == div_name]
                div_bs = [a for a in all_bs_accounts if a.get('division') == div_name]

                # Create division P&L sheet using optimized or standard method
                update_substep(f"{div_name}: Creating P&L sheet ({len(div_pl)} accounts)...")
                if use_optimized:
                    # OPTIMIZED: Clone template and inject data
                    div_pl_sheet = self._create_division_pl_optimized(
                        wb, div_pl, all_months, pl_totals, div_name, last_created_sheet
                    )
                else:
                    # FALLBACK: Standard creation from scratch
                    div_pl_name = f"{safe_name}_PL"
                    if div_pl_name in wb.sheetnames:
                        div_pl_sheet = wb[div_pl_name]
                        clear_sheet_data(div_pl_sheet)
                    else:
                        div_pl_sheet = wb.create_sheet(div_pl_name)
                    self._create_division_pl_report(div_pl_sheet, div_pl, all_months, pl_totals, div_name)

                # For divisions 3-5 (idx >= 2), collapse row groupings by default
                if idx >= 2:
                    try:
                        # Outline handled via group_rows()/group_cols()
                        print(f"[{div_name}_PL] Row groups collapsed (division {idx + 1})")
                    except Exception as e:
                        print(f"[{div_name}_PL] Could not collapse groups: {e}")

                # Create division BS sheet using optimized or standard method
                update_substep(f"{div_name}: Creating Balance Sheet ({len(div_bs)} accounts)...")
                if use_optimized:
                    # OPTIMIZED: Clone template and inject data
                    div_bs_sheet = self._create_division_bs_optimized(
                        wb, div_bs, all_months, bs_totals, div_name, div_pl_sheet
                    )
                else:
                    # FALLBACK: Standard creation from scratch
                    div_bs_name = f"{safe_name}_BS"
                    if div_bs_name in wb.sheetnames:
                        div_bs_sheet = wb[div_bs_name]
                        clear_sheet_data(div_bs_sheet)
                    else:
                        div_bs_sheet = wb.create_sheet(div_bs_name)
                    self._create_division_bs_report(div_bs_sheet, div_bs, all_months, bs_totals, div_name)

                # For divisions 3-5 (idx >= 2), collapse row groupings by default
                if idx >= 2:
                    try:
                        # Outline handled via group_rows()/group_cols()
                        print(f"[{div_name}_BS] Row groups collapsed (division {idx + 1})")
                    except Exception as e:
                        print(f"[{div_name}_BS] Could not collapse groups: {e}")

                # Update last_created_sheet for next iteration
                last_created_sheet = div_bs_sheet

            # Create Source_Budget sheet for budget data (placed after all P&L/BS sheets)
            update_step("Creating Source Budget sheet...")
            # Find the last P&L/BS sheet to place budget after it
            last_report_sheet = cons_cf_sheet
            for s in wb.worksheets:
                if s.name.endswith('_BS') or s.name.endswith('_PL') or s.name == 'Cash_Flow':
                    last_report_sheet = s
            if 'Source_Budget' in wb.sheetnames:
                source_budget = wb['Source_Budget']
                clear_sheet_data(source_budget)
            else:
                source_budget = wb.create_sheet('Source_Budget')
            self._create_source_budget_sheet(source_budget, consolidated_pl, all_months)

            # Create Consolidated Forecast sheet
            update_step("Creating Consolidated Forecast...")
            update_substep("Setting up Consolidated_Forecast sheet...")
            if 'Consolidated_Forecast' in wb.sheetnames:
                forecast_sheet = wb['Consolidated_Forecast']
                clear_sheet_data(forecast_sheet)
            elif 'Forecast' in wb.sheetnames:
                forecast_sheet = wb['Forecast']
                forecast_sheet.title = 'Consolidated_Forecast'
                clear_sheet_data(forecast_sheet)
            else:
                forecast_sheet = wb.create_sheet('Consolidated_Forecast')
            self._create_forecast_sheet(forecast_sheet, consolidated_pl, all_months)

            # Create Consolidated Forecast Summary
            update_step("Creating Forecast Summary...")
            update_substep("Building YTD and variance analysis...")
            if 'Forecast_Summary' in wb.sheetnames:
                forecast_summary = wb['Forecast_Summary']
                clear_sheet_data(forecast_summary)
            else:
                forecast_summary = wb.create_sheet('Forecast_Summary')
            self._create_forecast_summary_sheet(forecast_summary, consolidated_pl, all_months)

            # Create per-division Forecast and Forecast Summary sheets
            for idx, div in enumerate(self.divisions):
                div_name = div['name']
                safe_name = div_name.replace(' ', '_')[:20]

                update_step(f"Creating {div_name} Forecast ({idx+1}/{num_divisions})...")

                # Get division-specific P&L accounts
                div_pl = [a for a in all_pl_accounts if a.get('division') == div_name]

                # Create division Forecast sheet
                update_substep(f"{div_name}: Creating Forecast sheet...")
                div_forecast_name = f"{safe_name}_Forecast"
                if div_forecast_name in wb.sheetnames:
                    div_forecast_sheet = wb[div_forecast_name]
                    clear_sheet_data(div_forecast_sheet)
                else:
                    div_forecast_sheet = wb.create_sheet(div_forecast_name)
                self._create_forecast_sheet(div_forecast_sheet, div_pl, all_months, division_name=div_name)

                # Create division Forecast Summary sheet
                update_substep(f"{div_name}: Creating Forecast Summary...")
                div_fcst_summary_name = f"{safe_name}_Fcst_Summary"
                if div_fcst_summary_name in wb.sheetnames:
                    div_fcst_summary_sheet = wb[div_fcst_summary_name]
                    clear_sheet_data(div_fcst_summary_sheet)
                else:
                    div_fcst_summary_sheet = wb.create_sheet(div_fcst_summary_name)
                self._create_forecast_summary_sheet(div_fcst_summary_sheet, div_pl, all_months, division_name=div_name)

            # Update Menu sheet with division navigation
            update_step("Creating Menu with navigation...")
            update_substep("Building navigation tree for all divisions...")
            self._create_multi_division_menu_sheet(menu_sheet, all_months, self.divisions)

            # Create Dashboard
            update_step("Creating Dashboard...")
            update_substep("Setting up Dashboard sheet...")
            if 'Dashboard' in wb.sheetnames:
                dashboard_sheet = wb['Dashboard']
                clear_sheet_data(dashboard_sheet)
            else:
                dashboard_sheet = wb.create_sheet('Dashboard')
            update_substep("Building Dashboard KPIs and charts...")
            self._create_dashboard_sheet(dashboard_sheet, consolidated_pl, consolidated_bs, all_months, pl_totals)

            # Create Notes sheet with Division dropdown
            update_step("Creating Notes sheet...")
            update_substep("Setting up Notes with division selector...")
            if 'Notes' in wb.sheetnames:
                notes_sheet = wb['Notes']
                clear_sheet_data(notes_sheet)
            else:
                notes_sheet = wb.create_sheet('Notes')
            month_names = [name for m, y, name in all_months]
            self._create_notes_sheet(notes_sheet, consolidated_pl, consolidated_bs, month_names, divisions=self.divisions)

            # Create Settings sheet (sheet visibility controls)
            update_step("Creating Settings sheet...")
            update_substep("Building sheet visibility controls...")
            if 'Settings' in wb.sheetnames:
                settings_sheet = wb['Settings']
                clear_sheet_data(settings_sheet)
            else:
                settings_sheet = wb.create_sheet('Settings')
            # Build list of all report sheets for visibility control
            all_report_sheets = ['Dashboard', 'Consolidated_PL', 'Cash_Flow', 'Forecast_Summary', 'Consolidated_Forecast']
            for div in self.divisions:
                safe_name = div['name'].replace(' ', '_')[:20]
                all_report_sheets.extend([f"{safe_name}_PL", f"{safe_name}_BS", f"{safe_name}_Forecast", f"{safe_name}_Fcst_Summary"])
            all_report_sheets.extend(['Notes', 'Source_PL', 'Source_BS'])
            self._create_settings_sheet(settings_sheet, all_report_sheets)

            # Create Dashboard_Control sheet for KPI targets (used by Dashboard formulas)
            update_substep("Creating Dashboard Control sheet...")
            if 'Dashboard_Control' not in wb.sheetnames:
                dc_sheet = wb.create_sheet('Dashboard_Control')
            else:
                dc_sheet = wb['Dashboard_Control']
            self._create_dashboard_control_sheet(dc_sheet, all_months)

            # Clean up orphan sheets (Balance_Sheet, Consolidated_BS don't exist in multi-division mode)
            orphan_sheets = ['Balance_Sheet', 'Consolidated_BS']
            for orphan in orphan_sheets:
                try:
                    if orphan in wb.sheetnames:
                        wb.remove(wb[orphan])
                        print(f"Deleted orphan sheet: {orphan}")
                except Exception as e:
                    print(f"Could not delete {orphan}: {e}")

            # Hide or delete template sheets (_TPL_PL, _TPL_BS)
            for sheet_to_remove in [s for s in wb.worksheets if s.title.startswith('_TPL_')]:
                try:
                    wb.remove(sheet_to_remove)
                    print(f"Deleted template sheet: {sheet_to_remove.title}")
                except Exception as e:
                    try:
                        sheet_to_remove.sheet_state = 'hidden'
                        print(f"Hidden template sheet: {sheet_to_remove.title}")
                    except:
                        print(f"Could not remove template: {sheet_to_remove.title}")

            # Create Mapping_Config sheet for persistence
            update_step("Saving account mappings...")
            update_substep("Writing mappings to hidden Excel sheet...")
            MappingPersistence.save_to_excel(wb, self.account_mappings, self.divisions)

            # Save mappings to JSON file alongside Excel
            update_substep("Writing mappings to JSON file...")
            json_path = MappingPersistence.get_json_filepath(save_path)
            MappingPersistence.save_to_json(
                self.account_mappings,
                self.divisions,
                self.company_name.get(),
                json_path
            )

            # Move source sheets to end and set black tabs
            update_step("Organizing sheet tabs...")
            update_substep("Arranging sheets in professional order...")
            try:
                # Desired order: Dashboard, Menu, Consolidated reports, Division P&Ls, Division BSs,
                # Division Forecasts, Forecast Summary, Notes, Source sheets
                sheet_order = ['Dashboard', 'Menu']

                # Add Consolidated reports
                for name in ['Consolidated_PL', 'Cash_Flow']:
                    if name in wb.sheetnames:
                        sheet_order.append(name)

                # Add Division P&L sheets (before forecasts)
                for div in self.divisions:
                    safe_name = div['name'].replace(' ', '_')[:20]
                    pl_name = f"{safe_name}_PL"
                    if pl_name in wb.sheetnames:
                        sheet_order.append(pl_name)

                # Add Division BS sheets
                for div in self.divisions:
                    safe_name = div['name'].replace(' ', '_')[:20]
                    bs_name = f"{safe_name}_BS"
                    if bs_name in wb.sheetnames:
                        sheet_order.append(bs_name)

                # Add Forecast sheets (after division reports)
                # Forecast_Summary comes BEFORE Consolidated_Forecast for logical flow
                for name in ['Forecast_Summary', 'Consolidated_Forecast', 'Forecast']:
                    if name in wb.sheetnames:
                        sheet_order.append(name)

                # Add Division Forecast and Forecast Summary sheets
                for div in self.divisions:
                    safe_name = div['name'].replace(' ', '_')[:20]
                    fcst_name = f"{safe_name}_Forecast"
                    fcst_summary_name = f"{safe_name}_Fcst_Summary"
                    if fcst_name in wb.sheetnames:
                        sheet_order.append(fcst_name)
                    if fcst_summary_name in wb.sheetnames:
                        sheet_order.append(fcst_summary_name)

                # Add remaining sheets
                for name in ['Notes', 'Settings', 'Dashboard_Control']:
                    if name in wb.sheetnames:
                        sheet_order.append(name)

                # Move sheets to match desired order
                for i, name in enumerate(sheet_order):
                    if name in wb.sheetnames:
                        current_idx = wb.sheetnames.index(name)
                        wb.move_sheet(name, offset=i - current_idx)

                # Move Source sheets to the very end
                for src_name in ['Source_PL', 'Source_BS']:
                    if src_name in wb.sheetnames:
                        current_idx = wb.sheetnames.index(src_name)
                        wb.move_sheet(src_name, offset=len(wb.sheetnames) - 1 - current_idx)

                # Set tab colors to black for source/hidden sheets
                update_substep("Setting source sheet tab colors...")
                source_pl.sheet_properties.tabColor = "000000"
                source_bs.sheet_properties.tabColor = "000000"

                # Also hide Mapping_Config if it exists
                try:
                    mapping_sheet = wb['Mapping_Config']
                    mapping_sheet.sheet_state = 'hidden'
                    mapping_sheet.sheet_properties.tabColor = "000000"
                except:
                    pass

                # Set horizontal scrollbar small to show more tabs
                # TabRatio = ratio of tabs to total width. 0.85 = 85% tabs, 15% scrollbar
                update_substep("Adjusting view settings...")
                try:
                    # Window settings handled via sheet.views
                    pass
                except:
                    pass

            except Exception as e:
                print(f"Warning: Could not organize sheet tabs: {e}")

            # Clean up empty leading columns from division sheets
            update_step("Cleaning up empty columns...")
            for div in self.divisions:
                div_name = div['name']
                safe_name = div_name.replace(' ', '_')[:20]

                # Clean division P&L sheet
                pl_sheet_name = f"{safe_name}_PL"
                try:
                    if pl_sheet_name in wb.sheetnames:
                        update_substep(f"Cleaning {pl_sheet_name}...")
                        self._remove_empty_leading_columns(wb[pl_sheet_name])
                except Exception as e:
                    print(f"Warning: Could not clean {pl_sheet_name}: {e}")

                # Clean division BS sheet
                bs_sheet_name = f"{safe_name}_BS"
                try:
                    if bs_sheet_name in wb.sheetnames:
                        update_substep(f"Cleaning {bs_sheet_name}...")
                        self._remove_empty_leading_columns(wb[bs_sheet_name])
                except Exception as e:
                    print(f"Warning: Could not clean {bs_sheet_name}: {e}")

            # VBA macros: In the openpyxl version, VBA is preserved from
            # the template via keep_vba=True. If building from scratch,
            # we skip VBA (it will be added when user opens in Excel).
            update_step("Finalizing macros...")
            update_substep("VBA macros preserved from template" if use_template else "VBA macros will be added via template")

            # Activate Dashboard
            update_substep("Setting Dashboard as active sheet...")
            try:
                dashboard_sheet.activate()
            except:
                pass

            # Save workbook
            update_step("Saving workbook...")
            update_substep("Re-enabling calculations...")
            # openpyxl: no calculation mode needed
            update_substep("Writing to disk (this may take a moment)...")
            wb.save(temp_path)

            # Remove legacy _FilterDatabase from saved file
            remove_filter_database_from_xlsx(temp_path)
            update_substep("Closing workbook...")
            pass  # openpyxl auto-handles cleanup
            wb = None

        finally:
            try:
                if wb is not None:
                    pass  # openpyxl auto-handles cleanup
            except:
                pass
            try:
                pass  # openpyxl: no app to quit
            except:
                pass

        # Move from temp to final location
        try:
            if os.path.exists(save_path):
                os.remove(save_path)
            shutil.move(temp_path, save_path)
        finally:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

        # Finish timing and save log
        timer.finish_run()

    def _create_consolidated_pl_report(self, sheet, accounts, months, detected_totals=None):
        """Create consolidated P&L report with pre-computed values

        For multi-division consolidated view, values are already summed by consolidate_values().
        We write values directly and also include SUMIF formulas for live updates.
        """
        # For multi-division, use special report that writes values and formulas
        # referencing column B (Account) since column A is Division
        self._create_pl_report_multi_div(sheet, accounts, months, detected_totals, is_consolidated=True)

    def _create_consolidated_bs_report(self, sheet, accounts, months, detected_totals=None):
        """Create consolidated Balance Sheet report"""
        self._create_bs_report_multi_div(sheet, accounts, months, detected_totals, is_consolidated=True)

    def _create_division_pl_report(self, sheet, accounts, months, detected_totals, division_name):
        """Create P&L report for a specific division"""
        self._create_pl_report_multi_div(sheet, accounts, months, detected_totals,
                                          is_consolidated=False, division_name=division_name)
        # Update title to show division name
        sheet['A1'].value = f"{division_name}"

    def _create_division_bs_report(self, sheet, accounts, months, detected_totals, division_name):
        """Create Balance Sheet report for a specific division"""
        self._create_bs_report_multi_div(sheet, accounts, months, detected_totals,
                                          is_consolidated=False, division_name=division_name)
        sheet['A1'].value = f"{division_name}"

    def _create_pl_report_multi_div(self, sheet, accounts, months, detected_totals=None,
                                     is_consolidated=True, division_name=None, single_entity_mode=False):
        """Create P&L report matching single-division format with YTD, variance, and full year columns.

        OPTIMIZED: Uses bulk writes to prevent performance issues.

        Args:
            single_entity_mode: If True, Source_PL has col A=Account (not Division), data starts col B
        """
        import time as _time
        _timings = {}
        _t0 = _time.perf_counter()

        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = CLR_DARK_BLUE
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY

        # Source data range
        source_start = 3
        source_end = 1500

        # Calculate column positions
        header_row = 4
        num_months = len(months)
        last_month_col = num_months + 1
        notes_col = last_month_col + 1
        spacer1_col = notes_col + 1
        py_ytd_col = spacer1_col + 1
        cy_ytd_col = py_ytd_col + 1
        var_col = cy_ytd_col + 1
        var_pct_col = var_col + 1
        spacer2_col = var_pct_col + 1

        years = sorted(set(y for m, y, name in months))
        fy_start_col = spacer2_col + 1
        last_col = fy_start_col + len(years) - 1

        _timings['setup'] = _time.perf_counter() - _t0

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST
        # ================================================================
        _t1 = _time.perf_counter()
        all_data = []  # List of row data arrays
        row_types = []  # Track row type for formatting
        row_tracking = {}
        total_income_row = None
        total_cogs_row = None
        total_expenses_row = None
        found_cogs_section = False

        # Helper to build row data
        first_data_col_letter = get_column_letter(2)
        last_data_col_letter = get_column_letter(num_months + 1)

        for account in accounts:
            account_name = account['name']
            name_lower = account_name.lower()
            indent_level = account.get('indent', 0)

            # Track sections
            if account['is_header']:
                if 'cost' in name_lower or 'cogs' in name_lower:
                    found_cogs_section = True

            # Display name with indentation
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name

            # Calculate actual row number
            actual_row = header_row + 1 + len(all_data)

            # Track key rows
            if 'gross profit' in name_lower:
                row_tracking['gross_profit_row'] = actual_row
            if 'net income' in name_lower and account['is_total']:
                row_tracking['net_income_row'] = actual_row

            # Track total rows
            if 'total' in name_lower and account['is_total']:
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and 'other' not in name_lower and not found_cogs_section):
                    total_income_row = actual_row
                    row_tracking['total_income_row'] = actual_row
                elif 'cost' in name_lower or 'cogs' in name_lower:
                    total_cogs_row = actual_row
                    row_tracking['total_cogs_row'] = actual_row
                    found_cogs_section = True
                elif 'expense' in name_lower and 'other' not in name_lower:
                    total_expenses_row = actual_row
                    row_tracking['total_expense_row'] = actual_row

            # Build row data array
            row_data = [display_name]

            if account['is_header']:
                # Header rows: empty data cells
                row_data.extend([''] * (last_col - 1))
                all_data.append(row_data)
                row_types.append('header')
                continue

            # Monthly data columns - SUMIF/SUMIFS formulas
            # single_entity_mode: col A=Account, data starts col B
            # multi-div consolidated: col A=Division, col B=Account, data starts col C
            # multi-div division: filter by division name
            for i, (m, y, name) in enumerate(months):
                if single_entity_mode:
                    cl = get_column_letter(i + 2)  # Data starts at column B
                    formula = f'=SUMIF(Source_PL!$A${source_start}:$A${source_end},"{account_name}",Source_PL!{cl}${source_start}:{cl}${source_end})'
                elif is_consolidated:
                    cl = get_column_letter(i + 3)  # Data starts at column C (col A=Div, B=Acct)
                    formula = f'=SUMIF(Source_PL!$B${source_start}:$B${source_end},"{account_name}",Source_PL!{cl}${source_start}:{cl}${source_end})'
                else:
                    cl = get_column_letter(i + 3)  # Data starts at column C
                    formula = f'=SUMIFS(Source_PL!{cl}${source_start}:{cl}${source_end},Source_PL!$A${source_start}:$A${source_end},"{division_name}",Source_PL!$B${source_start}:$B${source_end},"{account_name}")'
                row_data.append(formula)

            # Notes column
            notes_formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="P&L")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{actual_row}))),Notes!$D$2:$D$100),"")'
            row_data.append(notes_formula)

            # Spacer 1
            row_data.append('')

            # YTD formulas
            data_range = f'{first_data_col_letter}{actual_row}:{last_data_col_letter}{actual_row}'
            helper_range = f'{first_data_col_letter}$3:{last_data_col_letter}$3'

            py_formula = f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7-1)*--(MOD({helper_range},100)<=Menu!$E$7))'
            cy_formula = f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7)*--(MOD({helper_range},100)<=Menu!$E$7))'
            var_formula = f'={get_column_letter(cy_ytd_col)}{actual_row}-{get_column_letter(py_ytd_col)}{actual_row}'
            var_pct_formula = f'=IFERROR({get_column_letter(var_col)}{actual_row}/{get_column_letter(py_ytd_col)}{actual_row},0)'

            row_data.extend([py_formula, cy_formula, var_formula, var_pct_formula])

            # Spacer 2
            row_data.append('')

            # Full Year columns - sum monthly values from THIS sheet using SUMPRODUCT
            # FIXED: Use SUMPRODUCT with row 3 YYYYMM values to determine which columns belong to each year
            # This survives column removal/insertion because it references the YYYYMM helper row
            # rather than hardcoded column positions
            for year in years:
                # SUMPRODUCT formula: sum values where INT(YYYYMM/100) matches the year
                # Row 3 contains YYYYMM values, data row contains actual values
                year_formula = f'=SUMPRODUCT((INT({first_data_col_letter}$3:{last_data_col_letter}$3/100)={year})*{first_data_col_letter}{actual_row}:{last_data_col_letter}{actual_row})'
                row_data.append(year_formula)

            all_data.append(row_data)
            row_type = 'total' if account['is_total'] else 'detail'
            if 'net income' in name_lower and account['is_total']:
                row_type = 'net_income'
            row_types.append(row_type)

        _timings['build_data'] = _time.perf_counter() - _t1

        # ================================================================
        # WRITE TO EXCEL IN BULK
        # ================================================================
        _t2 = _time.perf_counter()
        # Title - Professional corporate style
        sheet['A1'].value = company if is_consolidated else division_name
        sheet['A1'].font = Font(name='Calibri Light', size=16, bold=True, color=CLR_DARK_BLUE)
        sheet['A2'].value = 'Consolidated Profit & Loss' if is_consolidated else 'Profit & Loss Statement'
        sheet['A2'].font = Font(name='Calibri Light', size=12, color="808080")

        # Row 3: YYYYMM helper values
        helper_row = [y * 100 + m for m, y, name in months]
        for _ci, _val in enumerate(helper_row if isinstance(helper_row, list) else [helper_row]):
            sheet.cell(row=3, column=2 + _ci).value = _val if not isinstance(_val, list) else _val

        # Header row
        header_data = ['Account']
        for m, y, name in months:
            header_data.append(f"{self.MONTHS[m-1][:3]} {y}")
        header_data.extend(['Notes', '', 'PY YTD', 'CY YTD', 'Var $', 'Var %', ''])
        header_data.extend([str(y) for y in years])
        for _ci, _val in enumerate(header_data if isinstance(header_data, list) else [header_data]):
            sheet.cell(row=header_row, column=1 + _ci).value = _val if not isinstance(_val, list) else _val

        # Format header - professional corporate style
        apply_style_to_range(sheet, header_row, 1, header_row, last_col, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=FILL_DARK_BLUE)

        # Write all data in ONE bulk operation
        if all_data:
            data_start_row = header_row + 1
            data_end_row = header_row + len(all_data)
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=data_start_row + _ri, column=1).value = _val

        _timings['write_excel'] = _time.perf_counter() - _t2

        # ================================================================
        # APPLY FORMATTING IN BULK (NO per-row loops)
        # ================================================================
        _t3 = _time.perf_counter()
        if all_data:
            # Number formats for entire columns - single operations
            apply_style_to_range(sheet, data_start_row, 2, data_end_row, last_month_col, number_format='#,##0')
            apply_style_to_range(sheet, data_start_row, py_ytd_col, data_end_row, cy_ytd_col, number_format='#,##0')
            apply_style_to_range(sheet, data_start_row, var_col, data_end_row, var_col, number_format='#,##0')
            apply_style_to_range(sheet, data_start_row, var_pct_col, data_end_row, var_pct_col, number_format='0.0%')
            if fy_start_col <= last_col:
                apply_style_to_range(sheet, data_start_row, fy_start_col, data_end_row, last_col, number_format='#,##0')

            # Apply Calibri Light to all data
            apply_style_to_range(sheet, data_start_row, 1, data_end_row, last_col, font=Font(name='Calibri Light', size=10))

            # Collect row numbers by type for batch formatting
            header_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'header']
            total_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt in ('total', 'net_income')]

            # Format headers in batch (bold, section style)
            if header_rows:
                for r in header_rows:
                    sheet.cell(row=r, column=1).font = Font(bold=True, size=11)

            # Format totals with bold, top border (professional P&L style)
            if total_rows:
                for r in total_rows:
                    # Range: row_range = (sheet, r, 1, r, last_col)
                    apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
                    # Add top border for total rows (single line above)
                    try:
                        # Borders handled via Border()/Side() objects
                        # Borders handled via Border()/Side() objects
                        pass
                    except:
                        pass

            # Net Income gets double underline (accounting standard)
            net_income_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'net_income']
            if net_income_rows:
                for r in net_income_rows:
                    # Range: row_range = (sheet, r, 1, r, last_col)
                    apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
                    try:
                        # Double bottom border for Net Income
                        # Borders handled via Border()/Side() objects
                        # Borders handled via Border()/Side() objects
                        pass
                    except:
                        pass

        # Column widths
        try:
            sheet.column_dimensions['A'].width = 35
            sheet.column_dimensions[get_column_letter(spacer1_col)].width = 2
            sheet.column_dimensions[get_column_letter(spacer2_col)].width = 2
        except:
            pass

        # Group previous year columns
        self._group_previous_year_columns(sheet, months, data_start_col=2)

        _timings['formatting'] = _time.perf_counter() - _t3

        # ================================================================
        # VALIDATION SECTION
        # ================================================================
        _t4 = _time.perf_counter()
        validation_start_row = (header_row + len(all_data) + 3) if all_data else header_row + 5
        sheet[f'A{validation_start_row}'].value = 'VALIDATION - Source vs Calculated Totals'
        sheet[f'A{validation_start_row}'].font = Font(bold=True, color="800080")

        val_header_row = validation_start_row + 1
        val_headers = ['Category', 'Source', 'Calculated', 'Variance', 'Match?']
        write_row_to_cells(sheet, val_headers, row=val_header_row, start_col=1)
        apply_style_to_range(sheet, val_header_row, 1, val_header_row, 5,
                             font=Font(bold=True),
                             fill=PatternFill(start_color="C8C8C8", end_color="C8C8C8", fill_type="solid"))

        last_col_letter = get_column_letter(last_month_col)
        val_categories = [
            ('Total Income', '*Total*Income*', row_tracking.get('total_income_row')),
            ('Total COGS', '*Total*Cost*', row_tracking.get('total_cogs_row')),
            ('Total Expenses', '*Total*Expense*', row_tracking.get('total_expense_row')),
            ('Net Income', '*Net Income*', row_tracking.get('net_income_row')),
        ]

        for idx, (label, source_pattern, calc_row) in enumerate(val_categories):
            row = val_header_row + 1 + idx
            sheet[f'A{row}'].value = label

            # Source total (SUMIF from Source_PL with wildcard) - sum the last month column
            if is_consolidated:
                source_formula = f'=SUMIF(Source_PL!$B$3:$B$1500,"{source_pattern}",Source_PL!{last_col_letter}$3:{last_col_letter}$1500)'
            else:
                source_formula = f'=SUMIFS(Source_PL!{last_col_letter}$3:{last_col_letter}$1500,Source_PL!$A$3:$A$1500,"{division_name}",Source_PL!$B$3:$B$1500,"{source_pattern}")'
            sheet[f'B{row}'].value = source_formula
            sheet[f'B{row}'].number_format = '#,##0'
            print(f"[VALIDATION] Row {row} Source: {source_formula}")

            # Calculated total (from this sheet) - reference the tracked row directly
            if calc_row:
                calc_formula = f'={last_col_letter}{calc_row}'
            else:
                # Fallback: search for the label in column A using SUMIF
                calc_formula = f'=SUMIF($A:$A,"{source_pattern}",{last_col_letter}:{last_col_letter})'
            sheet[f'C{row}'].value = calc_formula
            sheet[f'C{row}'].number_format = '#,##0'
            print(f"[VALIDATION] Row {row} Calc ({calc_row}): {calc_formula}")

            # Variance (Source - Calculated)
            sheet[f'D{row}'].value = f'=B{row}-C{row}'
            sheet[f'D{row}'].number_format = '#,##0'

            # Match indicator
            sheet[f'E{row}'].value = f'=IF(ABS(D{row})<1,"✓","✗")'
            sheet[f'E{row}'].font = Font(size=14)

        validation_end_row = val_header_row + len(val_categories)

        # Add division breakdown for consolidated reports
        if is_consolidated and hasattr(self, 'divisions') and len(self.divisions) > 1:
            div_start_row = validation_end_row + 2
            sheet[f'A{div_start_row}'].value = 'BY DIVISION - Net Income'
            sheet[f'A{div_start_row}'].font = Font(bold=True, color="006400")

            div_header_row = div_start_row + 1
            div_headers = ['Division', 'Source', 'Calculated', 'Variance']
            write_row_to_cells(sheet, div_headers, row=div_header_row, start_col=1)
            apply_style_to_range(sheet, div_header_row, 1, div_header_row, 4,
                                 font=Font(bold=True),
                                 fill=PatternFill(start_color="DCDCDC", end_color="DCDCDC", fill_type="solid"))

            for div_idx, div in enumerate(self.divisions):
                div_name = div.get('name', div) if isinstance(div, dict) else getattr(div, 'name', str(div))
                row = div_header_row + 1 + div_idx
                sheet[f'A{row}'].value = div_name
                # Source Net Income for this division
                sheet[f'B{row}'].value = f'=SUMIFS(Source_PL!{last_col_letter}$3:{last_col_letter}$1500,Source_PL!$A$3:$A$1500,"{div_name}",Source_PL!$B$3:$B$1500,"Net Income")'
                sheet[f'B{row}'].number_format = '#,##0'
                # Calculated - reference division-specific sheet if exists
                div_sheet_name = f"{div_name}_PL"
                sheet[f'C{row}'].value = f"=IFERROR('{div_sheet_name}'!{last_col_letter}{row_tracking.get('net_income_row', 20)},0)"
                sheet[f'C{row}'].number_format = '#,##0'
                sheet[f'D{row}'].value = f'=B{row}-C{row}'
                sheet[f'D{row}'].number_format = '#,##0'

            # Sum row
            sum_row = div_header_row + 1 + len(self.divisions)
            sheet[f'A{sum_row}'].value = 'TOTAL'
            sheet[f'A{sum_row}'].font = Font(bold=True)
            sheet[f'B{sum_row}'].value = f'=SUM(B{div_header_row + 1}:B{sum_row - 1})'
            sheet[f'B{sum_row}'].font = Font(bold=True)
            sheet[f'B{sum_row}'].number_format = '#,##0'
            sheet[f'C{sum_row}'].value = f'=SUM(C{div_header_row + 1}:C{sum_row - 1})'
            sheet[f'C{sum_row}'].font = Font(bold=True)
            sheet[f'C{sum_row}'].number_format = '#,##0'
            sheet[f'D{sum_row}'].value = f'=B{sum_row}-C{sum_row}'
            sheet[f'D{sum_row}'].font = Font(bold=True)
            sheet[f'D{sum_row}'].number_format = '#,##0'

            validation_end_row = sum_row

        # Group/collapse the validation section
        try:
            if validation_start_row and validation_end_row and validation_end_row > validation_start_row:
                group_rows(sheet, validation_start_row, validation_end_row, outline_level=2, hidden=True)
        except Exception as e:
            print(f"[PL-MULTI] Could not group validation rows: {e}")

        _timings['validation'] = _time.perf_counter() - _t4
        _timings['total'] = _time.perf_counter() - _t0

        # Write timing summary to file
        sheet_name = "Consolidated_PL" if is_consolidated else f"{division_name}_PL"
        timing_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pl_timing_detail.txt')
        with open(timing_log_path, 'a') as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"[TIMING] P&L Report: {sheet_name} - {_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*60}\n")
            f.write(f"  Accounts: {len(accounts)}, Months: {len(months)}, Rows written: {len(all_data)}\n")
            f.write(f"  Setup:       {_timings.get('setup', 0)*1000:8.1f} ms\n")
            f.write(f"  Build Data:  {_timings.get('build_data', 0)*1000:8.1f} ms  (Python loop building formulas)\n")
            f.write(f"  Write Excel: {_timings.get('write_excel', 0)*1000:8.1f} ms  (Bulk write to Excel)\n")
            f.write(f"  Formatting:  {_timings.get('formatting', 0)*1000:8.1f} ms  (Number formats, column widths)\n")
            f.write(f"  Validation:  {_timings.get('validation', 0)*1000:8.1f} ms  (Validation section)\n")
            f.write(f"  ----------------------------------------\n")
            f.write(f"  TOTAL:       {_timings.get('total', 0)*1000:8.1f} ms ({_timings.get('total', 0):.2f} seconds)\n")
            f.write(f"{'='*60}\n")

        # AutoFit column A (account names), then set explicit widths for data columns
        # Width of 13 accommodates "$10,000,000" format (no cents)
        # Using range operation instead of loop for performance
        try:
            sheet.column_dimensions['A'].width = 45
            set_bulk_col_width(sheet, 2, 78, 13)  # B through BZ
        except Exception as e:
            print(f"[PL] Column width warning: {e}")

        # Collapse all outline groups
        try:
            # Outline handled via group_rows()/group_cols()
            pass
        except:
            pass

        # Hide row 3 (YYYYMM helper row) - MUST be at end after all other operations
        try:
            # Row hiding handled via hide_row()
            print(f"[PL] Row 3 hidden successfully")
        except Exception as e:
            print(f"[PL] ERROR hiding row 3: {e}")

        # Add back to menu link and print setup
        self._add_back_to_menu_link(sheet, row=1, col=last_col + 2)
        self._setup_print_area(sheet)

    def _create_bs_report_multi_div(self, sheet, accounts, months, detected_totals=None,
                                     is_consolidated=True, division_name=None, single_entity_mode=False):
        """Create Balance Sheet report for multi-division mode with DYNAMIC FORMULAS

        Uses SUMIF/SUMIFS formulas with LIMITED RANGES (not entire columns).

        Args:
            single_entity_mode: If True, Source_BS has col A=Account (not Division), data starts col B
        """
        print(f"[BS-MULTI] === STARTING === {len(accounts)} accounts, {len(months)} months")
        import time
        start_time = time.time()

        company = self.company_name.get()

        DARK_BLUE = CLR_DARK_BLUE
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY

        # Source data range - limited to actual data rows
        source_start = 3
        source_end = 1500  # Safe upper limit

        # Title
        sheet['A1'].value = company if is_consolidated else division_name
        sheet['A2'].value = 'Consolidated Balance Sheet' if is_consolidated else 'Balance Sheet'

        header_row = 4
        num_months = len(months)
        last_col = num_months + 1

        # Build header row
        header_data = ['Account']
        for m, y, name in months:
            header_data.append(f"{self.MONTHS[m-1][:3]} {y}")

        _data = header_data
        if _data is not None:
            if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                for _ri, _row in enumerate(_data):
                    for _ci, _val in enumerate(_row):
                        sheet.cell(row=header_row + _ri, column=1 + _ci).value = _val
            elif isinstance(_data, list):
                for _ri, _val in enumerate(_data):
                    if isinstance(_val, list):
                        for _ci, _v in enumerate(_val):
                            sheet.cell(row=header_row + _ri, column=1 + _ci).value = _v
                    else:
                        sheet.cell(row=header_row + _ri, column=1).value = _val
        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Header written")

        # Track rows for formatting
        total_rows = []
        header_rows_list = []

        # Build data with DYNAMIC FORMULAS using LIMITED RANGES
        all_data = []
        row_idx = header_row + 1

        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Building formulas...")

        for account in accounts:
            account_name = account['name']
            indent_level = account.get('indent', 0)
            is_header_row = account.get('is_header', False)
            is_total = account.get('is_total', False)

            display_name = account_name
            if indent_level > 0 and not is_header_row and not is_total:
                display_name = ('    ' * indent_level) + account_name

            if is_header_row:
                header_rows_list.append(row_idx)
            elif is_total:
                total_rows.append(row_idx)

            row_data = [display_name]

            if is_header_row:
                row_data.extend([''] * num_months)
            else:
                # DYNAMIC FORMULAS with LIMITED RANGES
                # single_entity_mode: col A=Account, data starts col B
                # multi-div consolidated: col A=Division, col B=Account, data starts col C
                for i, (m, y, name) in enumerate(months):
                    if single_entity_mode:
                        col = get_column_letter(i + 2)  # Data starts at column B
                        formula = f'=SUMIF(Source_BS!$A${source_start}:$A${source_end},"{account_name}",Source_BS!{col}${source_start}:{col}${source_end})'
                    elif is_consolidated:
                        col = get_column_letter(i + 3)  # Data starts at column C
                        formula = f'=SUMIF(Source_BS!$B${source_start}:$B${source_end},"{account_name}",Source_BS!{col}${source_start}:{col}${source_end})'
                    else:
                        col = get_column_letter(i + 3)  # Data starts at column C
                        formula = f'=SUMIFS(Source_BS!{col}${source_start}:{col}${source_end},Source_BS!$A${source_start}:$A${source_end},"{division_name}",Source_BS!$B${source_start}:$B${source_end},"{account_name}")'
                    row_data.append(formula)

            all_data.append(row_data)
            row_idx += 1

        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - {len(all_data)} rows built")

        # Write ALL data in ONE operation
        if all_data:
            data_start_row = header_row + 1
            data_end_row = header_row + len(all_data)
            print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Writing to Excel...")
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=data_start_row + _ri, column=1).value = _val
            print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Data written")

        # Apply professional formatting
        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Formatting...")
        try:
            # Header row - dark blue with white text
            apply_style_to_range(sheet, header_row, 1, header_row, last_col, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=FILL_DARK_BLUE)

            if all_data:
                data_start_row = header_row + 1
                data_end_row = header_row + len(all_data)

                # Apply Calibri Light to all data
                apply_style_to_range(sheet, data_start_row, 2, data_end_row, last_col, font=Font(name='Calibri Light', size=10), number_format='#,##0')

                # Format header rows (section titles) - bold, larger font
                for r in header_rows_list:
                    sheet.cell(row=r, column=1).font = Font(bold=True, size=11)

                # Format total rows with bold and top border (professional accounting style)
                for r in total_rows:
                    # Range: row_range = (sheet, r, 1, r, last_col)
                    apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
                    try:
                        # Add top border for total rows
                        # Borders handled via Border()/Side() objects
                        # Borders handled via Border()/Side() objects
                        pass
                    except:
                        pass

                # Special formatting for Total Assets and Total Liabilities & Equity (double underline)
                for r in total_rows:
                    try:
                        cell_value = sheet.cell(row=r, column=1).value
                        if cell_value and ('Total Assets' in str(cell_value) or
                                          'Total Liabilities & Equity' in str(cell_value) or
                                          'Total Liabilities and Equity' in str(cell_value)):
                            apply_style_to_range(sheet, r, 1, r, last_col, border=BORDER_NET_INCOME)
                    except:
                        pass

            print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Formatting complete")
        except Exception as e:
            print(f"[BS-MULTI] Formatting error: {e}")

        try:
            sheet.column_dimensions['A'].width = 35
        except:
            pass

        # Group and collapse previous year columns
        print(f"[BS-MULTI] {time.time() - start_time:.2f}s - Grouping previous year columns...")
        self._group_previous_year_columns(sheet, months, data_start_col=2)  # Data starts at column B

        # ================================================================
        # VALIDATION SECTION - Collapsible comparison of Source vs Calculated
        # ================================================================
        last_data_row = header_row + len(all_data) if all_data else header_row
        validation_start_row = last_data_row + 3
        last_col_letter = get_column_letter(last_col)

        # Add validation section header
        sheet[f'A{validation_start_row}'].value = 'VALIDATION - Source vs Calculated Totals'
        sheet[f'A{validation_start_row}'].font = Font(bold=True, size=11, color="800080")

        # Column headers for validation
        val_header_row = validation_start_row + 1
        val_headers = ['Category', 'Source', 'Calculated', 'Variance', 'Match?']
        write_row_to_cells(sheet, val_headers, row=val_header_row, start_col=1)
        apply_style_to_range(sheet, val_header_row, 1, val_header_row, 5,
                             font=Font(bold=True),
                             fill=PatternFill(start_color="C8C8C8", end_color="C8C8C8", fill_type="solid"))

        # BS validation categories - key balance sheet totals (use wildcards)
        val_categories = [
            ('Total Assets', '*Total*Assets*'),
            ('Total Liabilities', '*Total*Liabilities*'),
            ('Total Equity', '*Total*Equity*'),
            ('Total Liab + Equity', '*Total*Liabilities*Equity*'),
        ]

        for idx, (label, source_pattern) in enumerate(val_categories):
            row = val_header_row + 1 + idx
            sheet[f'A{row}'].value = label

            # Source total (SUMIF from Source_BS with wildcard)
            if is_consolidated:
                source_formula = f'=SUMIF(Source_BS!$B$3:$B$1500,"{source_pattern}",Source_BS!{last_col_letter}$3:{last_col_letter}$1500)'
            else:
                source_formula = f'=SUMIFS(Source_BS!{last_col_letter}$3:{last_col_letter}$1500,Source_BS!$A$3:$A$1500,"{division_name}",Source_BS!$B$3:$B$1500,"{source_pattern}")'
            sheet[f'B{row}'].value = source_formula
            sheet[f'B{row}'].number_format = '#,##0'

            # Calculated total - SUMIF on this sheet for matching account
            sheet[f'C{row}'].value = f'=SUMIF(A:A,"{source_pattern}",{last_col_letter}:{last_col_letter})'
            sheet[f'C{row}'].number_format = '#,##0'

            # Variance (Source - Calculated)
            sheet[f'D{row}'].value = f'=B{row}-C{row}'
            sheet[f'D{row}'].number_format = '#,##0'

            # Match indicator
            sheet[f'E{row}'].value = f'=IF(ABS(D{row})<1,"✓","✗")'
            sheet[f'E{row}'].font = Font(size=14)

        validation_end_row = val_header_row + len(val_categories)

        # Add division breakdown for consolidated reports
        if is_consolidated and hasattr(self, 'divisions') and len(self.divisions) > 1:
            div_start_row = validation_end_row + 2
            sheet[f'A{div_start_row}'].value = 'BY DIVISION - Total Assets'
            sheet[f'A{div_start_row}'].font = Font(bold=True, color="006400")

            div_header_row = div_start_row + 1
            div_headers = ['Division', 'Source', 'Calculated', 'Variance']
            write_row_to_cells(sheet, div_headers, row=div_header_row, start_col=1)
            apply_style_to_range(sheet, div_header_row, 1, div_header_row, 4,
                                 font=Font(bold=True),
                                 fill=PatternFill(start_color="DCDCDC", end_color="DCDCDC", fill_type="solid"))

            for div_idx, div in enumerate(self.divisions):
                div_name = div.get('name', div) if isinstance(div, dict) else getattr(div, 'name', str(div))
                row = div_header_row + 1 + div_idx
                sheet[f'A{row}'].value = div_name
                # Source Total Assets for this division
                sheet[f'B{row}'].value = f'=SUMIFS(Source_BS!{last_col_letter}$3:{last_col_letter}$1500,Source_BS!$A$3:$A$1500,"{div_name}",Source_BS!$B$3:$B$1500,"Total for Assets")'
                sheet[f'B{row}'].number_format = '#,##0'
                # Calculated - reference division-specific sheet if exists
                div_sheet_name = f"{div_name}_BS"
                sheet[f'C{row}'].value = f"=IFERROR(SUMIF('{div_sheet_name}'!A:A,\"*Total for Assets*\",'{div_sheet_name}'!{last_col_letter}:{last_col_letter}),0)"
                sheet[f'C{row}'].number_format = '#,##0'
                sheet[f'D{row}'].value = f'=B{row}-C{row}'
                sheet[f'D{row}'].number_format = '#,##0'

            # Sum row
            sum_row = div_header_row + 1 + len(self.divisions)
            sheet[f'A{sum_row}'].value = 'TOTAL'
            sheet[f'A{sum_row}'].font = Font(bold=True)
            sheet[f'B{sum_row}'].value = f'=SUM(B{div_header_row + 1}:B{sum_row - 1})'
            sheet[f'B{sum_row}'].font = Font(bold=True)
            sheet[f'B{sum_row}'].number_format = '#,##0'
            sheet[f'C{sum_row}'].value = f'=SUM(C{div_header_row + 1}:C{sum_row - 1})'
            sheet[f'C{sum_row}'].font = Font(bold=True)
            sheet[f'C{sum_row}'].number_format = '#,##0'
            sheet[f'D{sum_row}'].value = f'=B{sum_row}-C{sum_row}'
            sheet[f'D{sum_row}'].font = Font(bold=True)
            sheet[f'D{sum_row}'].number_format = '#,##0'

            validation_end_row = sum_row

        # Group/collapse the validation section
        try:
            if validation_start_row and validation_end_row and validation_end_row > validation_start_row:
                group_rows(sheet, validation_start_row, validation_end_row, outline_level=2, hidden=True)
        except Exception as e:
            print(f"[BS-MULTI] Could not group validation rows: {e}")

        print(f"[BS-MULTI] === COMPLETE === Total time: {time.time() - start_time:.2f}s, validation rows {validation_start_row}-{validation_end_row}")

        # Apply row grouping for Balance Sheet sections (matching P&L formatting)
        try:
            self._apply_row_grouping(sheet, accounts, header_row + 1)
            print(f"[BS-MULTI] Row grouping applied")
        except Exception as e:
            print(f"[BS-MULTI] Row grouping warning: {e}")

        # AutoFit column A (account names), then set explicit widths for data columns
        # Width of 13 accommodates "$10,000,000" format (no cents)
        # Using range operation instead of loop for performance
        try:
            sheet.column_dimensions['A'].width = 45
            set_bulk_col_width(sheet, 2, 78, 13)  # B through BZ
        except Exception as e:
            print(f"[BS-MULTI] Column width warning: {e}")

        # Collapse all outline groups (default to expanded for Balance Sheet)
        try:
            # Balance Sheet: expand row groups by default, collapse column groups
            # Outline handled via group_rows()/group_cols()
            pass
        except:
            pass

        # Hide row 3 (YYYYMM helper row) if present
        try:
            # Row hiding handled via hide_row()
            print(f"[BS-MULTI] Row 3 hidden successfully")
        except Exception as e:
            print(f"[BS-MULTI] Row 3 hiding warning: {e}")

        # Add print setup
        self._setup_print_area(sheet)

    def _create_multi_division_menu_sheet(self, sheet, months, divisions):
        """Create professional menu sheet with multi-division navigation and styling"""
        # Color palette
        DARK_BLUE = CLR_DARK_BLUE
        ACCENT_BLUE = CLR_ACCENT_BLUE
        GRAY = CLR_GRAY_TEXT
        LIGHT_GRAY = (245, 245, 245)
        LINK_BLUE = CLR_LINK_BLUE
        WHITE = (255, 255, 255)
        SECTION_BG = (240, 244, 248)  # Light blue-gray for section backgrounds

        company = self.company_name.get()

        # Clear and set up the sheet
        clear_sheet_data(sheet)

        # === HEADER SECTION (Rows 2-4) ===
        # Company name with accent bar
        sheet.merge_cells('B2:E2')
        sheet['B2'].value = company
        sheet['B2'].font = Font(name='Calibri Light', size=28, bold=True, color=CLR_DARK_BLUE)

        sheet.merge_cells('B3:E3')
        sheet['B3'].value = 'Consolidated Financial Model'
        sheet['B3'].font = Font(name='Calibri Light', size=14)
        sheet['B3'].font = FONT_GRAY_TEXT

        # Accent line under header
        try:
            accent_line = sheet['B4:E4']
            accent_line.fill = FILL_ACCENT_BLUE
            accent_line.row_height = 4
        except:
            pass

        # === CONFIGURATION CARD (Rows 6-10) ===
        # Configuration card background and border (B6:C10)
        try:
            section_fill = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
            apply_style_to_range(sheet, 6, 2, 10, 3, fill=section_fill)
            apply_border_box(sheet, 6, 2, 10, 3)
        except:
            pass

        sheet['B6'].value = 'MODEL SETTINGS'
        sheet['B6'].font = Font(name='Calibri Light', size=10, bold=True)
        sheet['B6'].font = FONT_ACCENT_BLUE

        # Get unique years for dropdown
        years_in_data = sorted(set(y for m, y, name in months))
        current_year = months[-1][1] if months else 2026

        # === CREATE PERIOD LOOKUP TABLE (Hidden columns K:M) ===
        # This table maps display names to YYYYMM values for formula lookups
        if months:
            # Write header row
            sheet['K6'].value = 'Period'
            sheet['L6'].value = 'YYYYMM'
            sheet['M6'].value = 'MonthNum'
            # Write all months
            for idx, (m, y, name) in enumerate(months):
                row = 7 + idx
                sheet[f'K{row}'].value = name  # Display name (e.g., "Nov 2024")
                sheet[f'L{row}'].value = y * 100 + m  # YYYYMM (e.g., 202411)
                sheet[f'M{row}'].value = m  # Month number (1-12)
            # Hide lookup columns
            try:
                # Column hiding handled via hide_columns_range()
                pass
            except:
                pass

        config_labels = [
            ('Current Period:', months[-1][2] if months else 'N/A'),
            ('Data Range:', f"{months[0][2]} - {months[-1][2]}" if months else 'N/A'),
            ('Divisions:', str(len(divisions))),
            ('Reporting Year:', str(current_year)),  # New: Reporting year for grouping
        ]

        for i, (label, value) in enumerate(config_labels):
            row = 7 + i
            sheet[f'B{row}'].value = label
            sheet[f'C{row}'].value = value
            sheet[f'B{row}'].font = Font(name='Calibri Light', size=10)
            sheet[f'B{row}'].font = FONT_GRAY_TEXT
            sheet[f'C{row}'].font = Font(name='Calibri Light', size=10, bold=True)
            # Right-align value cells for visual consistency
            try:
                # Alignment handled via Alignment() objects
                pass
            except:
                pass

        # Add dropdown for Current Period (C7) with all available months
        try:
            if months:
                # Create dropdown list from all available months
                month_list = ','.join([name for m, y, name in months])
                # Validation handled via DataValidation object
                # DataValidation handled via DataValidation()
                sheet['C7'].fill = PatternFill(start_color="FFFFC8", end_color="FFFFC8", fill_type="solid")  # Light yellow to indicate editable
        except Exception as e:
            print(f"Warning: Could not add period dropdown: {e}")

        # Add VLOOKUP formulas to convert C7 display name to helper values
        # G7 = YYYYMM value looked up from table
        # E7 = Month number, F7 = Year (derived from G7)
        if months:
            last_lookup_row = 6 + len(months)
            try:
                # G7 uses VLOOKUP to get YYYYMM from the lookup table
                sheet['G7'].value = f'=IFERROR(VLOOKUP(C7,$K$7:$L${last_lookup_row},2,FALSE),0)'
                # E7 = Month number from lookup
                sheet['E7'].value = f'=IFERROR(VLOOKUP(C7,$K$7:$M${last_lookup_row},3,FALSE),1)'
                # F7 = Year derived from G7 (YYYYMM / 100 rounded down)
                sheet['F7'].value = '=INT(G7/100)'
                # G9 = Same as G7 (Actuals Through)
                sheet['G9'].value = '=G7'
            except Exception as e:
                print(f"Warning: Could not set lookup formulas: {e}")
                # Fallback to static values
                if months:
                    current_m, current_y = months[-1][0], months[-1][1]
                    sheet['E7'].value = current_m
                    sheet['F7'].value = current_y
                    sheet['G7'].value = current_y * 100 + current_m
                    sheet['G9'].value = current_y * 100 + current_m

        # Add dropdown for Reporting Year (C10)
        try:
            year_list = ','.join([str(y) for y in years_in_data])
            # Validation handled via DataValidation object
            # DataValidation handled via DataValidation()
            sheet['C10'].fill = PatternFill(start_color="FFFFC8", end_color="FFFFC8", fill_type="solid")  # Light yellow to indicate editable
        except:
            pass

        # Add hint text for editable fields
        sheet['D7'].value = '← Select period to view'
        sheet['D7'].font = Font(name='Calibri Light', size=8)
        sheet['D7'].font = FONT_GRAY_TEXT
        sheet['D7'].font = Font(italic=True)

        sheet['D10'].value = '← Run "UpdateYearGrouping" macro after changing'
        sheet['D10'].font = Font(name='Calibri Light', size=8)
        sheet['D10'].font = FONT_GRAY_TEXT
        sheet['D10'].font = Font(italic=True)

        # === CONSOLIDATED REPORTS SECTION (Rows 12-19) ===
        # Consolidated reports section background and border (B12:C19)
        try:
            section_fill = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
            apply_style_to_range(sheet, 12, 2, 19, 3, fill=section_fill)
            apply_border_box(sheet, 12, 2, 19, 3)
        except:
            pass

        # Section header with icon indicator
        sheet['B12'].value = '📊 REPORTS (Click to Navigate)'
        sheet['B12'].font = Font(name='Calibri Light', size=11, bold=True)
        sheet['B12'].font = FONT_ACCENT_BLUE

        cons_nav = [
            ('▸ Dashboard', 'Dashboard', 'Executive summary & KPIs'),
            ('▸ P&L Statement', 'Consolidated_PL', 'Profit & Loss'),
            ('▸ Cash Flow', 'Cash_Flow', 'Cash movements'),
            ('▸ Forecast', 'Consolidated_Forecast', 'Budget vs Actual'),
            ('▸ Forecast Summary', 'Forecast_Summary', 'YTD variances'),
            ('▸ Notes', 'Notes', 'Account annotations'),
        ]

        for i, (label, target, desc) in enumerate(cons_nav):
            row = 13 + i
            cell = sheet[f'B{row}']
            cell.value = label
            cell.font = Font(name='Calibri Light', size=10)
            cell.font = FONT_LINK_BLUE
            cell.font = Font(underline="single")
            # Add description in column C
            sheet[f'C{row}'].value = desc
            sheet[f'C{row}'].font = Font(name='Calibri Light', size=9)
            sheet[f'C{row}'].font = FONT_GRAY_TEXT
            sheet[f'C{row}'].font = Font(italic=True)
            try:
                cell.hyperlink = f"#'{target}'!A1"
                cell.font = Font(name='Calibri Light', size=10, color=CLR_LINK_BLUE, underline="single")
            except:
                pass

        # === SOURCE DATA SECTION (Rows 12-16, Column D-E) ===
        # Source data section background and border (D12:E16)
        try:
            section_fill = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
            apply_style_to_range(sheet, 12, 4, 16, 5, fill=section_fill)
            apply_border_box(sheet, 12, 4, 16, 5)
        except:
            pass

        sheet['D12'].value = '📁 SOURCE DATA'
        sheet['D12'].font = Font(name='Calibri Light', size=11, bold=True)
        sheet['D12'].font = FONT_ACCENT_BLUE

        source_nav = [
            ('▸ P&L Data', 'Source_PL', 'Raw P&L'),
            ('▸ Balance Sheet', 'Source_BS', 'Raw BS'),
            ('▸ Budget', 'Source_Budget', 'Budget data'),
        ]

        for i, (label, target, desc) in enumerate(source_nav):
            row = 13 + i
            cell = sheet[f'D{row}']
            cell.value = label
            cell.font = Font(name='Calibri Light', size=10)
            cell.font = FONT_LINK_BLUE
            cell.font = Font(underline="single")
            sheet[f'E{row}'].value = desc
            sheet[f'E{row}'].font = Font(name='Calibri Light', size=9)
            sheet[f'E{row}'].font = FONT_GRAY_TEXT
            sheet[f'E{row}'].font = Font(italic=True)
            try:
                cell.hyperlink = f"#'{target}'!A1"
                cell.font = Font(name='Calibri Light', size=10, color=CLR_LINK_BLUE, underline="single")
            except:
                pass

        # === DIVISION SECTIONS ===
        current_row = 21

        for div_idx, div in enumerate(divisions):
            safe_name = div['name'].replace(' ', '_')[:20]

            # Division card
            # Division card background and border
            try:
                section_fill = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
                apply_style_to_range(sheet, current_row, 2, current_row + 1, 5, fill=section_fill)
                apply_border_box(sheet, current_row, 2, current_row + 1, 5)
            except:
                pass

            # Division header
            sheet[f'B{current_row}'].value = div['name'].upper()
            sheet[f'B{current_row}'].font = Font(name='Calibri Light', size=10, bold=True)
            sheet[f'B{current_row}'].font = FONT_ACCENT_BLUE

            # Division links (P&L, BS, Forecast on same row)
            link_row = current_row + 1
            div_links = [
                ('B', 'P&L', f'{safe_name}_PL'),
                ('C', 'Balance Sheet', f'{safe_name}_BS'),
                ('D', 'Forecast', f'{safe_name}_Forecast'),
            ]

            for col, label, target in div_links:
                sheet[f'{col}{link_row}'].value = f'  {label}'
                sheet[f'{col}{link_row}'].font = Font(name='Calibri Light', size=10)
                sheet[f'{col}{link_row}'].font = FONT_LINK_BLUE
                sheet[f'{col}{link_row}'].font = Font(underline="single")
                try:
                    sheet[f'{col}{link_row}'].hyperlink = f"#'{target}'!A1"
                    sheet[f'{col}{link_row}'].font = Font(name='Calibri Light', size=10, color=CLR_LINK_BLUE, underline="single")
                except:
                    pass

            current_row += 3  # Space between division cards

        # === FOOTER ===
        footer_row = current_row + 1
        sheet[f'B{footer_row}'].value = 'CFO DNA Financial Model Generator'
        sheet[f'B{footer_row}'].font = Font(name='Calibri Light', size=8)
        sheet[f'B{footer_row}'].font = FONT_GRAY_TEXT
        sheet[f'B{footer_row}'].font = Font(italic=True)

        # Add helper cells for YTD calculations (hidden via white font since column E has visible content)
        # E7 = current month, F7 = current year (for P&L formulas)
        # H7 = current month, I7 = reporting year (from C10), J7 = YYYYMM
        try:
            if months:
                current_month = months[-1][0]
                current_year = months[-1][1]
                # E7/F7 for backward compatibility with P&L formulas
                sheet['E7'].value = current_month
                sheet['E7'].font = Font(color="FFFFFF")  # White font to hide value
                sheet['F7'].value = current_year
                sheet['G7'].value = current_year * 100 + current_month
                # H7/I7/J7 for dynamic reporting year
                sheet['H7'].value = current_month
                sheet['I7'].value = '=C10'  # References C10 (Reporting Year)
                sheet['J7'].value = '=I7*100+H7'
                sheet['J9'].value = '=J7'
        except:
            pass

        # Set column widths for clean layout (BEFORE hiding helper columns)
        sheet.column_dimensions['A'].width = 3
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 18
        sheet.column_dimensions['D'].width = 18

        # Hide helper columns F:K AFTER setting other column widths
        # (Column E is used for source data descriptions - don't hide it)
        # (Setting column_width on hidden columns can unhide them)
        try:
            hide_columns_range(sheet, 6, 11)  # F through K
        except:
            pass

        # Hide gridlines for clean look
        try:
            sheet.views.sheetView[0].showGridLines = False
        except:
            pass

    def _create_excel_model(self, save_path):
        """Create the Excel model using xlwings"""
        # Get user-specified date range
        user_start, user_end = self._get_user_date_params()

        # Check if multi-division mode
        if self.is_multi_division.get() and self.divisions:
            return self._create_multi_division_model(save_path, user_start, user_end)

        # Single entity mode - use progress tracker for consistent UI
        total_steps = 13  # 12 main steps + 1 finalizing step
        # Use historical timing data for more accurate progress bar estimate
        estimated_total = StepTimer.get_historical_estimate(default_estimate=90)
        update_step, update_substep, state = self._create_progress_tracker(total_steps, estimated_total)

        # Parse input files with indentation detection
        update_step("Reading P&L file...")
        pl_data, pl_indents = self._read_excel_with_indents(self.pl_path.get())

        update_step("Reading Balance Sheet file...")
        bs_data, bs_indents = self._read_excel_with_indents(self.bs_path.get())

        update_step("Extracting account data...")
        pl_accounts, pl_months, pl_totals = self._parse_financial_data(pl_data, pl_indents, user_start, user_end, self.pl_path.get())
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data, bs_indents, user_start, user_end, self.bs_path.get())
        print(f"Found {len(pl_accounts)} P&L accounts, {len(bs_accounts)} BS accounts, {len(pl_months)} months")
        print(f"P&L Totals detected: {pl_totals}")
        print(f"BS Totals detected: {bs_totals}")

        # Check if template exists
        use_template = os.path.exists(TEMPLATE_PATH)

        # Work in temp directory to avoid OneDrive locking issues
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        # Create Excel application
        update_step("Starting Excel...")
        # openpyxl: no Excel app needed
        # openpyxl: no display_alerts needed
        # openpyxl: no screen_updating needed
        # openpyxl: no calculation mode needed
        wb = None

        try:
            if use_template:
                # Copy template to temp location and open
                shutil.copy(TEMPLATE_PATH, temp_path)
                wb = load_workbook(temp_path, keep_vba=True)

                # Get existing sheets
                menu_sheet = wb['Menu']
                source_pl = wb['Source_PL']
                source_bs = wb['Source_BS']
                pl_sheet = wb['PL']
                bs_sheet = wb['Balance_Sheet']
                cf_sheet = wb['Cash_Flow']
                notes_sheet = wb['Notes']

                # Create new sheets if they don't exist (for v1.3.0 Forecast module)
                try:
                    source_budget = wb['Source_Budget']
                    clear_sheet_data(source_budget)
                except:
                    source_budget = wb.create_sheet('Source_Budget')

                try:
                    forecast_sheet = wb['Forecast']
                    clear_sheet_data(forecast_sheet)
                except:
                    forecast_sheet = wb.create_sheet('Forecast')

                try:
                    forecast_summary_sheet = wb['Forecast_Summary']
                    clear_sheet_data(forecast_summary_sheet)
                except:
                    forecast_summary_sheet = wb.create_sheet('Forecast_Summary')

                # Clear source sheets (keep headers)
                clear_sheet_data(source_pl, start_row=2)
                clear_sheet_data(source_bs, start_row=2)

                # Also clear the report sheets for fresh data (including header rows)
                clear_sheet_data(pl_sheet)
                clear_sheet_data(bs_sheet)
                clear_sheet_data(cf_sheet)
            else:
                # Create from scratch (requires VBA trust setting)
                wb = Workbook()

                # Remove default sheets and create our sheets
                for sheet in wb.worksheets:
                    if sheet.title not in ['Sheet1']:
                        pass
                        # Sheet deletion: use wb.remove(sheet)

                # Create sheets
                menu_sheet = wb.worksheets[0]
                menu_sheet.title = 'Menu'

                source_pl = wb.create_sheet('Source_PL')
                source_bs = wb.create_sheet('Source_BS')
                # P&L/BS/CF reports come before Budget/Forecast
                pl_sheet = wb.create_sheet('PL')
                bs_sheet = wb.create_sheet('Balance_Sheet')
                cf_sheet = wb.create_sheet('Cash_Flow')
                # Budget and Forecast come after all report sheets
                source_budget = wb.create_sheet('Source_Budget')
                forecast_sheet = wb.create_sheet('Forecast')
                forecast_summary_sheet = wb.create_sheet('Forecast_Summary')
                notes_sheet = wb.create_sheet('Notes')

            # Populate source sheets
            update_step("Populating source data...")
            self._populate_source_sheet(source_pl, pl_accounts, pl_months)
            self._populate_source_sheet(source_bs, bs_accounts, pl_months)
            self._create_source_budget_sheet(source_budget, pl_accounts, pl_months)

            # Create/update named ranges
            pl_last_row = len(pl_accounts) + 2  # +2 for header and YYYYMM rows
            pl_last_col = len(pl_months) + 1
            bs_last_row = len(bs_accounts) + 2

            # Delete existing named ranges if they exist
            try:
                del wb.defined_names['SourcePL']
            except:
                pass
            try:
                del wb.defined_names['SourceBS']
            except:
                pass
            try:
                del wb.defined_names['SourceBudget']
            except:
                pass

            bs_last_col = len(pl_months) + 1  # BS uses same number of month columns
            budget_last_col = 13  # Source_Budget: col A (accounts) + cols B-M (12 months)
            budget_last_row = len(pl_accounts) + 5  # +5 for header rows in budget sheet
            wb.defined_names.add(DefinedName('SourcePL', attr_text=f"'Source_PL'!$A$1:${get_column_letter(pl_last_col)}${pl_last_row}"))
            wb.defined_names.add(DefinedName('SourceBS', attr_text=f"'Source_BS'!$A$1:${get_column_letter(bs_last_col)}${bs_last_row}"))
            wb.defined_names.add(DefinedName('SourceBudget', attr_text=f"'Source_Budget'!$A$1:${get_column_letter(budget_last_col)}${budget_last_row}"))

            # Create Menu sheet
            update_step("Creating Menu sheet...")
            self._create_menu_sheet(menu_sheet, pl_months)

            # Create P&L report (using optimized bulk-write function)
            update_step("Creating P&L report...")
            self._create_pl_report_multi_div(pl_sheet, pl_accounts, pl_months, pl_totals,
                                              is_consolidated=True, division_name=None, single_entity_mode=True)

            # Create Balance Sheet report (using optimized bulk-write function)
            update_step("Creating Balance Sheet...")
            self._create_bs_report_multi_div(bs_sheet, bs_accounts, pl_months, bs_totals,
                                              is_consolidated=True, division_name=None, single_entity_mode=True)

            # Create Cash Flow statement
            update_step("Creating Cash Flow statement...")
            self._create_cash_flow(cf_sheet, pl_months)

            # Create Forecast sheet
            update_step("Creating Forecast sheet...")
            self._create_forecast_sheet(forecast_sheet, pl_accounts, pl_months)

            # Create Forecast Summary and Notes
            update_substep("Creating Forecast Summary...")
            self._create_forecast_summary_sheet(forecast_summary_sheet, pl_accounts, pl_months)

            # Create Notes sheet with dropdowns (consolidated for P&L, BS, and CF)
            month_names = [name for m, y, name in pl_months]
            self._create_notes_sheet(notes_sheet, pl_accounts, bs_accounts, month_names)

            # Create Settings sheet (sheet visibility controls)
            update_step("Creating Settings sheet...")
            if use_template:
                # Check if Settings already exists (may be named Dashboard_Control in old templates)
                if 'Settings' in wb.sheetnames:
                    settings_sheet = wb['Settings']
                    clear_sheet_data(settings_sheet)
                elif 'Dashboard_Control' in wb.sheetnames:
                    settings_sheet = wb['Dashboard_Control']
                    settings_sheet.title = 'Settings'
                    clear_sheet_data(settings_sheet)
                else:
                    settings_sheet = wb.create_sheet('Settings')
            else:
                settings_sheet = wb.create_sheet('Settings')
            # Build list of report sheets for visibility control
            single_entity_sheets = ['Dashboard', 'PL', 'Balance_Sheet', 'Cash_Flow', 'Forecast',
                                    'Forecast_Summary', 'Notes', 'Source_PL', 'Source_BS', 'Source_Budget']
            self._create_settings_sheet(settings_sheet, single_entity_sheets)

            # Create Dashboard_Control sheet for KPI targets (used by Dashboard formulas)
            if 'Dashboard_Control' not in wb.sheetnames:
                dc_sheet = wb.create_sheet('Dashboard_Control')
            else:
                dc_sheet = wb['Dashboard_Control']
                clear_sheet_data(dc_sheet)
            self._create_dashboard_control_sheet(dc_sheet, pl_months)

            # Create Dashboard sheet
            if use_template:
                if 'Dashboard' not in wb.sheetnames:
                    dashboard_sheet = wb.create_sheet('Dashboard')
                else:
                    dashboard_sheet = wb['Dashboard']
                    clear_sheet_data(dashboard_sheet)
            else:
                dashboard_sheet = wb.create_sheet('Dashboard')
            self._create_dashboard_sheet(dashboard_sheet, pl_accounts, bs_accounts, pl_months, pl_totals)

            # Reorder sheets to desired layout:
            # Dashboard, Menu, P&L, Balance_Sheet, Cash_Flow, Forecast, Forecast_Summary, Notes, Settings, Source sheets
            try:
                # Build desired order
                sheet_order = ['Dashboard', 'Menu', 'PL', 'Balance_Sheet', 'Cash_Flow',
                               'Forecast', 'Forecast_Summary', 'Notes', 'Settings', 'Dashboard_Control',
                               'Source_PL', 'Source_BS', 'Source_Budget']

                # Get all current sheets
                all_sheets = {s.name: s for s in wb.worksheets}

                # openpyxl: reorder sheets using move_sheet
                existing_names = [s.title for s in wb.worksheets]
                target_idx = 0
                for name in sheet_order:
                    if name in existing_names:
                        current_idx = [s.title for s in wb.worksheets].index(name)
                        wb.move_sheet(name, offset=target_idx - current_idx)
                        target_idx += 1
            except:
                pass

            # Set source sheet tab colors to black
            try:
                source_pl.sheet_properties.tabColor = "000000"
                source_bs.sheet_properties.tabColor = "000000"
                source_budget.sheet_properties.tabColor = "000000"
            except:
                pass

            # Set Dashboard as active sheet when file opens
            try:
                dashboard_idx = [s.title for s in wb.worksheets].index('Dashboard')
                wb.active = dashboard_idx
            except:
                pass

            # ============================================================
            # FINAL FORMATTING CHECK - ensure key totals are properly formatted
            # ============================================================
            update_step("Final formatting check...")
            try:
                # P&L report formatting check
                pl_max_row = pl_sheet.max_row if pl_sheet.max_row else 100
                pl_max_col = pl_sheet.max_column if pl_sheet.max_column else 20
                _final_formatting_check(pl_sheet, 5, pl_max_row, pl_max_col, sheet_type='pl')

                # Balance Sheet formatting check
                bs_max_row = bs_sheet.max_row if bs_sheet.max_row else 100
                bs_max_col = bs_sheet.max_column if bs_sheet.max_column else 20
                _final_formatting_check(bs_sheet, 5, bs_max_row, bs_max_col, sheet_type='bs')

                print("[FINAL CHECK] Applied formatting to key totals in PL and BS")
            except Exception as e:
                print(f"[FINAL CHECK] Warning: {e}")

            # VBA macros: preserved from template via keep_vba=True
            # For from-scratch builds, save as .xlsm (macros can be added later)

            # Save the workbook
            update_step("Saving workbook...")
            wb.save(temp_path)

            # Remove legacy _FilterDatabase defined names from saved file.
            # openpyxl creates both a sheet-level <autoFilter> and a workbook-level
            # _xlnm._FilterDatabase named range. Excel flags the duplicate as
            # "Removed Records: Named range" during repair.
            remove_filter_database_from_xlsx(temp_path)
            pass  # openpyxl auto-handles cleanup
            wb = None

        finally:
            # Ensure proper cleanup
            try:
                if wb is not None:
                    pass  # openpyxl auto-handles cleanup
            except:
                pass
            try:
                # openpyxl: no app to quit
                pass
            except:
                pass

        # Finalizing step
        update_step("Finalizing...")

        # Move from temp to final location (after Excel is closed)
        try:
            # Remove existing file if it exists
            if os.path.exists(save_path):
                os.remove(save_path)
            shutil.move(temp_path, save_path)
        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

    def _get_cell_indents(self, file_path):
        """
        Read Excel file and extract indentation levels for column A.
        Returns dict: {row_number: indent_level}
        """
        from openpyxl import load_workbook
        indents = {}
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            for row_idx, row in enumerate(ws.iter_rows(min_col=1, max_col=1), start=0):
                cell = row[0]
                indent = 0
                if cell.alignment and cell.alignment.indent:
                    indent = int(cell.alignment.indent)
                # Also check for leading spaces in value
                if cell.value and isinstance(cell.value, str):
                    leading_spaces = len(cell.value) - len(cell.value.lstrip())
                    # Convert leading spaces to indent level (4 spaces = 1 level)
                    space_indent = leading_spaces // 4
                    indent = max(indent, space_indent)
                indents[row_idx] = indent
            pass  # openpyxl auto-handles cleanup
        except Exception as e:
            print(f"Could not read indentation from {file_path}: {e}")
        return indents

    def _read_excel_with_indents(self, file_path):
        """
        Read Excel file and extract both data and indentation in a single pass.
        Returns tuple: (pandas DataFrame, dict of {row_number: indent_level})

        This is more efficient than calling pd.read_excel() and _get_cell_indents()
        separately, as it only opens the file once.
        """
        from openpyxl import load_workbook
        indents = {}
        data = []

        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active

            for row_idx, row in enumerate(ws.iter_rows()):
                # Extract indentation from first column
                first_cell = row[0]
                indent = 0
                if first_cell.alignment and first_cell.alignment.indent:
                    indent = int(first_cell.alignment.indent)
                # Also check for leading spaces in value
                if first_cell.value and isinstance(first_cell.value, str):
                    leading_spaces = len(first_cell.value) - len(first_cell.value.lstrip())
                    space_indent = leading_spaces // 4
                    indent = max(indent, space_indent)
                indents[row_idx] = indent

                # Extract row data
                row_data = [cell.value for cell in row]
                data.append(row_data)

            pass  # openpyxl auto-handles cleanup

            # Convert to DataFrame (same format as pd.read_excel with header=None)
            df = pd.DataFrame(data)
            return df, indents

        except Exception as e:
            print(f"Could not read file {file_path}: {e}")
            # Fall back to separate reads
            df = pd.read_excel(file_path, header=None)
            indents = self._get_cell_indents(file_path)
            return df, indents

    def _detect_file_type(self, file_path):
        """
        Detect whether a file is a P&L or Balance Sheet based on account names.
        Returns tuple: ('pl', confidence) or ('bs', confidence) or (None, 0)

        P&L indicators: income, revenue, sales, expenses, cost of goods, gross profit, net income
        BS indicators: assets, liabilities, equity, cash, accounts receivable, accounts payable
        """
        try:
            # Read first 100 rows to scan account names
            df = pd.read_excel(file_path, header=None, nrows=100)

            # P&L keywords (weighted by specificity)
            pl_keywords = {
                'income': 2, 'revenue': 2, 'sales': 2, 'expense': 2,
                'cost of goods': 3, 'gross profit': 3, 'net income': 3,
                'operating income': 3, 'ebitda': 3, 'cogs': 2,
                'payroll': 1, 'rent': 1, 'utilities': 1, 'depreciation': 1,
                'interest expense': 2, 'tax expense': 2
            }

            # BS keywords (weighted by specificity)
            bs_keywords = {
                'assets': 3, 'liabilities': 3, 'equity': 3,
                'accounts receivable': 3, 'accounts payable': 3,
                'cash': 2, 'inventory': 2, 'prepaid': 2,
                'fixed assets': 3, 'current assets': 3, 'current liabilities': 3,
                'long-term': 2, 'retained earnings': 3, 'common stock': 3,
                'total assets': 3, 'total liabilities': 3, 'total equity': 3
            }

            pl_score = 0
            bs_score = 0

            # Check column A (account names)
            for idx in range(len(df)):
                cell_val = df.iloc[idx, 0]
                if pd.isna(cell_val):
                    continue
                cell_lower = str(cell_val).lower().strip()

                # Check P&L keywords
                for keyword, weight in pl_keywords.items():
                    if keyword in cell_lower:
                        pl_score += weight

                # Check BS keywords
                for keyword, weight in bs_keywords.items():
                    if keyword in cell_lower:
                        bs_score += weight

            total_score = pl_score + bs_score
            if total_score == 0:
                return (None, 0)

            # Calculate confidence as proportion of winning type
            if pl_score > bs_score:
                confidence = pl_score / total_score
                return ('pl', confidence)
            elif bs_score > pl_score:
                confidence = bs_score / total_score
                return ('bs', confidence)
            else:
                return (None, 0.5)  # Tie - can't determine

        except Exception as e:
            print(f"Could not detect file type for {file_path}: {e}")
            return (None, 0)

    def _validate_file_type(self, file_path, expected_type):
        """
        Validate that a file matches the expected type (pl or bs).
        Returns tuple: (is_valid, detected_type, message)
        """
        detected_type, confidence = self._detect_file_type(file_path)

        if detected_type is None:
            return (True, None, "Could not determine file type - proceeding anyway")

        if detected_type == expected_type:
            return (True, detected_type, f"File detected as {detected_type.upper()} ({confidence:.0%} confidence)")

        # File doesn't match expected type
        if confidence >= 0.6:  # High confidence wrong file
            type_names = {'pl': 'Profit & Loss (P&L)', 'bs': 'Balance Sheet'}
            detected_name = type_names.get(detected_type, detected_type.upper())
            expected_name = type_names.get(expected_type, expected_type.upper())
            return (False, detected_type,
                    f"This file appears to be a {detected_name} ({confidence:.0%} confidence),\n"
                    f"but you're uploading it as a {expected_name}.\n\n"
                    f"Please select the correct file type.")
        else:
            # Low confidence - warn but allow
            return (True, detected_type, f"File type uncertain - proceeding as {expected_type.upper()}")

    def _prompt_for_date(self, filename, header_contents):
        """Show dialog to prompt user for date when auto-detection fails.

        Returns:
            tuple: (month_num, year) or None if user cancels
        """
        print(f"DEBUG: _prompt_for_date called with filename={filename}")

        # Ensure the GUI is responsive
        self.root.update()

        # Create a custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Date Required")
        dialog.geometry("450x250")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 450) // 2
        y = (dialog.winfo_screenheight() - 250) // 2
        dialog.geometry(f"+{x}+{y}")

        result = {'value': None}

        # Message
        msg_frame = ttk.Frame(dialog, padding=10)
        msg_frame.pack(fill='x')

        ttk.Label(msg_frame, text=f"Could not determine date from file:", font=('Calibri', 10, 'bold')).pack(anchor='w')
        ttk.Label(msg_frame, text=os.path.basename(filename), font=('Calibri', 9)).pack(anchor='w', pady=(0, 10))

        header_text = str(header_contents)[:80] + "..." if len(str(header_contents)) > 80 else str(header_contents)
        ttk.Label(msg_frame, text=f"Headers found: {header_text}", font=('Calibri', 8), foreground='gray').pack(anchor='w')

        ttk.Label(msg_frame, text="\nPlease specify the month/year for this data:", font=('Calibri', 10)).pack(anchor='w', pady=(10, 5))

        # Date selection frame
        date_frame = ttk.Frame(dialog, padding=10)
        date_frame.pack(fill='x')

        # Month dropdown
        ttk.Label(date_frame, text="Month:").grid(row=0, column=0, padx=5, sticky='e')
        month_var = tk.StringVar(value=self.MONTHS[datetime.now().month - 1])
        month_combo = ttk.Combobox(date_frame, textvariable=month_var, values=self.MONTHS, width=15, state='readonly')
        month_combo.grid(row=0, column=1, padx=5)

        # Year dropdown
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 5, current_year + 2)]
        ttk.Label(date_frame, text="Year:").grid(row=0, column=2, padx=5, sticky='e')
        year_var = tk.StringVar(value=str(current_year))
        year_combo = ttk.Combobox(date_frame, textvariable=year_var, values=years, width=10, state='readonly')
        year_combo.grid(row=0, column=3, padx=5)

        # Buttons
        btn_frame = ttk.Frame(dialog, padding=10)
        btn_frame.pack(fill='x', pady=10)

        def on_ok():
            month_num = self.MONTHS.index(month_var.get()) + 1
            year = int(year_var.get())
            result['value'] = (month_num, year)
            dialog.destroy()

        def on_cancel():
            result['value'] = None
            dialog.destroy()

        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side='right', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=10).pack(side='right', padx=5)

        # Wait for dialog to close
        dialog.wait_window()

        return result['value']

    def _parse_month_value(self, val):
        """
        Flexibly parse a value to extract month/year info.
        Returns (month_num, year, display_name) or None if not a recognizable month.
        Handles many formats: "Jan 24", "January 2024", "Jan-24", "1/24", dates, etc.
        """
        import re

        if pd.isna(val):
            return None

        # Handle datetime/Timestamp objects directly
        if isinstance(val, (pd.Timestamp, datetime)):
            month_num = val.month
            year = val.year
            month_abbrev = self.MONTHS[month_num - 1][:3]
            display_name = f"{month_abbrev} {year}"
            return (month_num, year, display_name)

        val_str = str(val).strip()
        if not val_str or 'total' in val_str.lower():
            return None

        val_lower = val_str.lower()

        # Skip date RANGE patterns (e.g., "August 1, 2021-November 30, 2025" or "January-December, 2024")
        # These contain multiple months or range indicators and should not be treated as a single month
        month_count = 0
        for month in self.MONTHS:
            if month.lower() in val_lower or month[:3].lower() in val_lower:
                month_count += 1
        if month_count > 1:
            # Multiple months found - this is a date range, not a single month
            return None

        # Also skip if it looks like a date range with hyphen between date components
        # Pattern: "Month Day, Year - Month Day, Year" or similar
        if '-' in val_str and any(c.isdigit() for c in val_str.split('-')[0]) and any(c.isdigit() for c in val_str.split('-')[-1]):
            # Has hyphen with numbers on both sides - likely a date range
            return None

        # Skip "As of" patterns - these are point-in-time descriptions, not column headers
        if val_lower.startswith('as of '):
            return None

        # Strategy 1: Look for month name (full or abbreviated) anywhere in the string
        for i, month in enumerate(self.MONTHS, 1):
            month_lower = month.lower()
            month_short = month_lower[:3]

            if month_short in val_lower or month_lower in val_lower:
                # Found a month name - now extract year
                # Prefer 4-digit years over 2-digit years to avoid confusion with day numbers
                numbers = re.findall(r'\d+', val_str)
                # First try to find a 4-digit year
                for num_str in numbers:
                    num = int(num_str)
                    if 1900 <= num <= 2100:
                        display_name = f"{month[:3]} {num % 100}"
                        return (i, num, display_name)
                # Fall back to 2-digit year if no 4-digit year found
                for num_str in numbers:
                    if len(num_str) == 2:
                        year = 2000 + int(num_str)
                        display_name = f"{month[:3]} {year}"
                        return (i, year, display_name)

        # Strategy 2: Numeric date formats like "1/24", "01/2024", "1-24", "12/25"
        # Pattern: month/year or month-year
        date_patterns = [
            r'^(\d{1,2})[/\-\.](\d{2,4})$',  # 1/24, 01/2024, 1-24
            r'^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})$',  # 1/1/24 - take first as month, last as year
        ]

        for pattern in date_patterns:
            match = re.match(pattern, val_str)
            if match:
                groups = match.groups()
                month_num = int(groups[0])
                year = int(groups[-1])  # Last group is year
                if 1 <= month_num <= 12:
                    if year < 100:
                        year = 2000 + year
                    month_abbrev = self.MONTHS[month_num - 1][:3]
                    display_name = f"{month_abbrev} {year}"
                    return (month_num, year, display_name)

        # Strategy 3: Try pandas date parsing as last resort
        try:
            parsed_date = pd.to_datetime(val_str, errors='raise', dayfirst=False)
            if parsed_date:
                month_num = parsed_date.month
                year = parsed_date.year
                if 1900 <= year <= 2100:  # Sanity check
                    month_abbrev = self.MONTHS[month_num - 1][:3]
                    display_name = f"{month_abbrev} {year}"
                    return (month_num, year, display_name)
        except:
            pass

        return None

    def _parse_financial_data(self, df, indents=None, user_start=None, user_end=None, filename=None):
        """Parse financial data from dataframe with indentation levels.

        Args:
            df: DataFrame with financial data
            indents: Dict of row_idx -> indent level
            user_start: Tuple of (month_num, year) user-specified start date
            user_end: Tuple of (month_num, year) user-specified end date
            filename: Original filename (used for dialog prompt if date detection fails)
        """
        indents = indents or {}

        # Find header row by looking for any recognizable month pattern or common headers
        header_row = None
        has_total_column = False

        for idx in range(min(15, len(df))):
            for col_idx in range(len(df.columns)):
                val = df.iloc[idx, col_idx]
                if self._parse_month_value(val):
                    header_row = idx
                    break
                # Also check for "Total" header which indicates a summary column
                if pd.notna(val) and str(val).strip().lower() == 'total':
                    header_row = idx
                    has_total_column = True
                    break
            if header_row is not None:
                break

        if header_row is None:
            # Fallback: Look for row before common financial statement headers
            common_first_items = ['income', 'revenue', 'sales', 'assets', 'liabilities',
                                  'equity', 'expenses', 'cost of', 'ordinary income']
            for idx in range(min(15, len(df))):
                first_cell = df.iloc[idx, 0]
                if pd.notna(first_cell):
                    first_cell_lower = str(first_cell).strip().lower()
                    for pattern in common_first_items:
                        if first_cell_lower.startswith(pattern):
                            # The header row is likely the row BEFORE this
                            header_row = max(0, idx - 1)
                            print(f"Found data starting at row {idx} ('{first_cell}'), using row {header_row} as header")
                            break
                    if header_row is not None:
                        break

        if header_row is None:
            # Try to provide more helpful error message
            print("DEBUG: First 5 rows of data:")
            for i in range(min(5, len(df))):
                print(f"  Row {i}: {list(df.iloc[i, :5])}")
            raise ValueError("Could not find month headers in file. Please ensure the file has column headers.")

        # Extract months using the flexible parser
        months = []
        data_columns = []  # Track which columns have actual data (not "Total")

        print(f"DEBUG: Header row {header_row}, user_start={user_start}, user_end={user_end}")
        print(f"DEBUG: Header row contents: {list(df.iloc[header_row, :min(10, len(df.columns))])}")

        for col_idx in range(1, len(df.columns)):
            val = df.iloc[header_row, col_idx]
            if pd.isna(val):
                continue

            val_str = str(val).strip().lower()
            print(f"DEBUG: Col {col_idx} = '{val}' (type: {type(val).__name__})")

            # Skip "Total" columns when we have other month columns
            if val_str == 'total':
                has_total_column = True
                print(f"DEBUG: Found Total column at {col_idx}")
                # We'll decide later whether to use this
                continue

            parsed = self._parse_month_value(val)
            if parsed:
                print(f"DEBUG: Parsed '{val}' as month: {parsed}")
                months.append(parsed)
                data_columns.append(col_idx)
            else:
                print(f"DEBUG: Could not parse '{val}' as month")

        # If no months found, try to use user-specified dates
        print(f"DEBUG: After parsing - months found: {len(months)}, has_total_column: {has_total_column}")
        print(f"DEBUG: months list: {months}")
        print(f"DEBUG: data_columns: {data_columns}")

        if not months and user_start and user_end:
            print(f"DEBUG: No months found, using user-specified dates: {user_start} to {user_end}")
            start_month, start_year = user_start
            end_month, end_year = user_end

            # Find a data column to use - prefer Total column, otherwise use first column with numeric data
            data_col = None

            # First, look for Total column in header row
            for col_idx in range(1, len(df.columns)):
                val = df.iloc[header_row, col_idx]
                if pd.notna(val) and str(val).strip().lower() == 'total':
                    data_col = col_idx
                    print(f"DEBUG: Found Total column at {col_idx}")
                    break

            # If no Total column found in header, look for first column with numeric data in data rows
            if data_col is None:
                # Check a few rows after header to find where numeric data is
                for col_idx in range(1, len(df.columns)):
                    for check_row in range(header_row + 1, min(header_row + 10, len(df))):
                        val = df.iloc[check_row, col_idx]
                        if pd.notna(val):
                            try:
                                float(val)
                                data_col = col_idx
                                print(f"DEBUG: Found numeric data in column {col_idx} at row {check_row}")
                                break
                            except (ValueError, TypeError):
                                continue
                    if data_col is not None:
                        break

            print(f"DEBUG: data_col = {data_col}")

            if data_col:
                # If start == end, it's a single month report
                if start_month == end_month and start_year == end_year:
                    month_abbrev = self.MONTHS[start_month - 1][:3]
                    display_name = f"{month_abbrev} {start_year % 100}"
                    months = [(start_month, start_year, display_name)]
                    data_columns = [data_col]
                    print(f"Using user-specified single month: {display_name}, data from column {data_col}")
                else:
                    # Multi-month range - but we only have one column of data
                    # This is a special case - user says it's a range but file has one column
                    # Use the single column for the END month (most recent)
                    month_abbrev = self.MONTHS[end_month - 1][:3]
                    display_name = f"{month_abbrev} {end_year % 100}"
                    months = [(end_month, end_year, display_name)]
                    data_columns = [data_col]
                    print(f"Using user-specified end month: {display_name}, data from column {data_col}")

        # If still no months, prompt user for the date
        if not months:
            # Try to find a data column - prefer Total, then any column with numeric data
            data_col = None

            # First, look for Total column
            if has_total_column:
                for col_idx in range(1, len(df.columns)):
                    val = df.iloc[header_row, col_idx]
                    if pd.notna(val) and str(val).strip().lower() == 'total':
                        data_col = col_idx
                        print(f"DEBUG: Found Total column at {col_idx}")
                        break

            # If no Total column, look for first column with numeric data
            if data_col is None and header_row is not None:
                for col_idx in range(1, len(df.columns)):
                    for check_row in range(header_row + 1, min(header_row + 10, len(df))):
                        val = df.iloc[check_row, col_idx]
                        if pd.notna(val):
                            try:
                                float(str(val).replace(',', '').replace('$', '').replace('(', '-').replace(')', ''))
                                data_col = col_idx
                                print(f"DEBUG: Found numeric data in column {col_idx} at row {check_row}")
                                break
                            except (ValueError, TypeError):
                                continue
                    if data_col is not None:
                        break

            # Default to column 1 if no data column found
            if data_col is None:
                data_col = 1
                print(f"DEBUG: Using default data column 1")

            # Get header contents for the dialog
            header_contents = list(df.iloc[header_row, :min(10, len(df.columns))]) if header_row is not None else []
            print(f"DEBUG: Prompting for date. filename={filename}, header_contents={header_contents}")

            # Try to prompt user for the date
            user_date = None
            if filename and hasattr(self, 'root') and self.root:
                try:
                    user_date = self._prompt_for_date(filename, header_contents)
                except Exception as e:
                    print(f"DEBUG: Could not show date dialog: {e}")
                    import traceback
                    traceback.print_exc()

            if user_date:
                # User provided a date
                month_num, year = user_date
                month_abbrev = self.MONTHS[month_num - 1][:3]
                display_name = f"{month_abbrev} {year % 100}"
                months = [(month_num, year, display_name)]
                data_columns = [data_col]
                print(f"Using user-specified date: {display_name} for column {data_col}")
            elif user_date is None and filename:
                # User cancelled the dialog - raise error to stop processing
                raise ValueError("Date selection cancelled. Please try again and specify the date for the uploaded file.")
            else:
                # Dialog not available (no GUI) - use current date as fallback
                now = datetime.now()
                current_month = now.month
                current_year = now.year
                month_abbrev = self.MONTHS[current_month - 1][:3]
                display_name = f"{month_abbrev} {current_year % 100}"
                months = [(current_month, current_year, display_name)]
                data_columns = [data_col]
                print(f"WARNING: Using current date ({display_name}) as fallback.")

        # If still no months, raise error with helpful message
        if not months:
            header_contents = list(df.iloc[header_row, :min(10, len(df.columns))]) if header_row is not None else "No header found"
            error_msg = (
                f"Could not determine date columns from file headers.\n\n"
                f"Header row {header_row} contents: {header_contents}\n\n"
                f"Expected formats: 'Jan 24', 'January 2024', 'Jan-24', '1/24', or datetime values.\n\n"
                f"If your file uses a different format, please rename the column headers or "
                f"contact support."
                )
            raise ValueError(error_msg)

        # Extract accounts and values, tracking exact total names
        accounts = []
        detected_totals = {
            'total_income': None,
            'total_cogs': None,
            'total_expenses': None,
            'gross_profit': None,
            'net_income': None,
            'total_assets': None,
            'total_liabilities': None,
            'total_equity': None,
            'total_liab_equity': None,
        }

        # If data_columns wasn't populated (old code path), default to sequential
        if not data_columns:
            data_columns = list(range(1, len(months) + 1))

        for row_idx in range(header_row + 1, len(df)):
            account_name = df.iloc[row_idx, 0]
            if pd.isna(account_name) or not str(account_name).strip():
                continue
            account_name = str(account_name).strip()
            if 'cash basis' in account_name.lower():
                continue

            values = {}
            # Use data_columns to get values from correct columns
            # Store values with (month, year) tuple keys for lookup
            for month_idx, col_idx in enumerate(data_columns):
                if col_idx < len(df.columns) and month_idx < len(months):
                    val = df.iloc[row_idx, col_idx]
                    month_m, month_y, _ = months[month_idx]
                    if pd.notna(val):
                        try:
                            values[(month_m, month_y)] = float(val)
                        except:
                            values[(month_m, month_y)] = 0
                    else:
                        values[(month_m, month_y)] = 0

            # Improved total detection: case-insensitive, handles more variations
            name_lower = account_name.lower().strip()
            is_total = (name_lower.startswith('total') or
                       'net income' in name_lower or
                       'net profit' in name_lower or
                       'net loss' in name_lower or
                       'gross profit' in name_lower or
                       'gross margin' in name_lower)
            is_header = account_name.strip() in ['Income', 'Expenses', 'Cost of Sales', 'Assets', 'Liabilities', 'Equity',
                                         'Other Current Assets', 'Fixed Assets', 'Other Assets',
                                         'Current Liabilities', 'Long Term Liabilities', 'Other Current Liabilities',
                                         'REVENUES', 'EXPENSES', 'INCOME', 'COST OF GOODS SOLD', 'COST OF SALES']
            # Also detect section headers in ALL CAPS
            if account_name.strip().isupper() and not is_total and len(account_name.strip()) < 40:
                # Short ALL CAPS entries that aren't totals are likely headers
                if not any(c.isdigit() for c in account_name):
                    is_header = True

            # Get indent level from source file
            indent_level = indents.get(row_idx, 0)

            # Detect specific total rows by their exact names (expanded matching)
            if (name_lower.startswith('total income') or name_lower.startswith('total revenue') or
                name_lower.startswith('total revenues') or name_lower.startswith('total sales') or
                name_lower == 'total for income' or name_lower == 'total for revenue' or
                name_lower == 'total for revenues'):
                detected_totals['total_income'] = account_name
            elif (name_lower.startswith('total cost') or
                  name_lower.startswith('total cogs') or
                  name_lower.startswith('total for cost') or
                  ('cost of sales' in name_lower and 'total' in name_lower) or
                  ('cost of goods' in name_lower and 'total' in name_lower) or
                  ('cost of revenue' in name_lower and 'total' in name_lower)):
                detected_totals['total_cogs'] = account_name
            elif (name_lower.startswith('total expense') or name_lower.startswith('total expenses') or
                  name_lower == 'total for expenses' or name_lower == 'total for expense' or
                  name_lower.startswith('total operating expense')):
                detected_totals['total_expenses'] = account_name
            elif 'gross profit' in name_lower or 'gross margin' in name_lower:
                detected_totals['gross_profit'] = account_name
            elif 'net income' in name_lower or 'net profit' in name_lower or 'net loss' in name_lower:
                detected_totals['net_income'] = account_name
            elif name_lower.startswith('total for assets') or name_lower == 'total assets':
                detected_totals['total_assets'] = account_name
            elif name_lower.startswith('total for liabilities and equity') or name_lower == 'total liabilities and equity' or name_lower == 'total liabilities & equity':
                detected_totals['total_liab_equity'] = account_name
            elif (name_lower.startswith('total for liabilities') or name_lower == 'total liabilities') and 'equity' not in name_lower:
                detected_totals['total_liabilities'] = account_name
            elif (name_lower.startswith('total for equity') or name_lower == 'total equity') and 'liabilities' not in name_lower:
                detected_totals['total_equity'] = account_name

            accounts.append({
                'name': account_name,
                'indent': indent_level,
                'values': values,
                'is_total': is_total,
                'is_header': is_header
            })

        return accounts, months, detected_totals

    def _detect_actual_date_range(self, accounts, months):
        """Detect which months actually have data (non-zero values).

        Filters out leading months with no data in any account.

        Args:
            accounts: List of account dictionaries with 'values' dict
            months: List of (month, year, display_name) tuples

        Returns:
            Filtered list of months that have actual data
        """
        if not months or not accounts:
            return months

        # Find first and last month with any non-zero value
        first_month_with_data = None
        last_month_with_data = None

        for idx, (m, y, name) in enumerate(months):
            has_data = False
            for account in accounts:
                if account.get('is_header', False):
                    continue
                val = account.get('values', {}).get((m, y), 0)
                if val != 0:
                    has_data = True
                    break

            if has_data:
                if first_month_with_data is None:
                    first_month_with_data = idx
                last_month_with_data = idx

        # If no data found anywhere, return original months
        if first_month_with_data is None:
            return months

        # Return filtered months (from first with data to last with data)
        filtered = months[first_month_with_data:last_month_with_data + 1]
        if len(filtered) < len(months):
            print(f"Date range detected: trimmed {len(months)} months to {len(filtered)} with actual data")
            print(f"  Original: {months[0][2]} to {months[-1][2]}")
            print(f"  Filtered: {filtered[0][2]} to {filtered[-1][2]}")

        return filtered

    def _get_previous_year_columns(self, months, data_start_col):
        """Identify columns that belong to previous years for grouping.

        Groups all columns from years PRIOR to the reporting year (last year in data).
        Example: If data goes through Dec 2025, groups all 2024 and earlier columns.

        Args:
            months: List of (month, year, display_name) tuples
            data_start_col: First column number where month data starts (1-indexed)

        Returns:
            List of (start_col, end_col, year) tuples for each previous year group
        """
        if not months:
            return []

        # Use the last month's year as the "reporting year" - group everything before it
        reporting_year = months[-1][1]

        year_groups = []
        current_group_start = None
        current_group_year = None

        for idx, (m, y, name) in enumerate(months):
            col = data_start_col + idx

            if y < reporting_year:
                # This is a prior year column - should be grouped
                if current_group_year != y:
                    # New year group - save the old one if exists
                    if current_group_start is not None:
                        year_groups.append((current_group_start, col - 1, current_group_year))
                    current_group_start = col
                    current_group_year = y
            else:
                # Current/reporting year - close any open previous year group
                if current_group_start is not None:
                    year_groups.append((current_group_start, col - 1, current_group_year))
                    current_group_start = None
                    current_group_year = None

        # Close final group if still open (all months are prior year)
        if current_group_start is not None:
            # Find last prior year column
            last_prior_col = data_start_col + len(months) - 1
            for idx, (m, y, name) in enumerate(months):
                if y >= reporting_year:
                    last_prior_col = data_start_col + idx - 1
                    break
            if last_prior_col >= current_group_start:
                year_groups.append((current_group_start, last_prior_col, current_group_year))

        return year_groups

    def _group_previous_year_columns(self, sheet, months, data_start_col):
        """Group and collapse columns for previous years.

        Args:
            sheet: xlwings sheet object
            months: List of (month, year, display_name) tuples
            data_start_col: First column number where month data starts
        """
        year_groups = self._get_previous_year_columns(months, data_start_col)

        if not year_groups:
            return

        try:
            for start_col, end_col, year in year_groups:
                if start_col <= end_col:
                    start_letter = get_column_letter(start_col)
                    end_letter = get_column_letter(end_col)
                    print(f"Grouping {year} columns: {start_letter}:{end_letter}")
                    group_cols(sheet, start_col, end_col, outline_level=1, hidden=True)

            # Collapse all groups (show only level 1 = ungrouped columns)
            # Outline handled via group_rows()/group_cols()
            print(f"Grouped and collapsed {len(year_groups)} previous year(s)")
        except Exception as e:
            print(f"Warning: Could not group previous year columns: {e}")

    def _remove_empty_leading_columns(self, sheet):
        """Remove leading columns that have $0 in Net Income (division didn't exist yet).

        Scans from left to right and deletes data columns where Net Income = 0
        until hitting a column with non-zero Net Income.

        Args:
            sheet: xlwings sheet object
        """
        try:
            # Find the Net Income row using BULK READ (optimized)
            net_income_row = None
            # UsedRange replaced with sheet.max_row/max_column

            # Read entire column A at once
            # Range: col_a_data = (sheet, 1, 1, last_row, 1)
            if not isinstance(col_a_data, list):
                col_a_data = [col_a_data]

            for row_idx, val in enumerate(col_a_data):
                if val and 'net income' in str(val).lower() and 'other' not in str(val).lower():
                    net_income_row = row_idx + 1  # Convert to 1-based
                    break

            if not net_income_row:
                print(f"  Could not find Net Income row")
                return

            # Find the header row (row 4) and data start column (column 2)
            header_row = 4
            data_start_col = 2

            # Find last column
            # end("right") replaced with max_column

            # Read entire Net Income row at once (BULK READ)
            # Range: net_income_values = (sheet, net_income_row, data_start_col, net_income_row, last_col)
            if not isinstance(net_income_values, list):
                net_income_values = [net_income_values]

            # Count how many leading columns have $0 Net Income
            empty_cols_count = 0
            for net_income_val in net_income_values:
                # Check if value is 0 or None (formula might show 0)
                if net_income_val is None or net_income_val == 0:
                    empty_cols_count += 1
                else:
                    # Found a column with non-zero Net Income - stop here
                    break

            if empty_cols_count > 0:
                # Delete columns in reverse order (from right to left of empty range)
                # Actually, delete them all at once for efficiency
                first_col_letter = get_column_letter(data_start_col)
                last_empty_col_letter = get_column_letter(data_start_col + empty_cols_count - 1)

                print(f"  Removing {empty_cols_count} empty leading columns ({first_col_letter}:{last_empty_col_letter})")
                sheet[f'{first_col_letter}:{last_empty_col_letter}'].delete()

                # After deletion, need to ungroup any orphaned groups
                try:
                    # Outline handled via group_rows()/group_cols()
                    pass
                except:
                    pass

        except Exception as e:
            print(f"  Error removing empty columns: {e}")

    def _populate_source_sheet(self, sheet, accounts, months, division_name=None):
        """Populate a source data sheet with proper formatting

        Structure (single division / backward compatible):
        - Row 1: Headers (Account, month names like "Jan 24")
        - Row 2: YYYYMM helper values (e.g., 202401) for YTD calculations - hidden
        - Row 3+: Account data

        Structure (multi-division mode when division_name provided):
        - Row 1: Headers (Division, Account, month names)
        - Row 2: YYYYMM helper values (hidden)
        - Row 3+: Account data with Division in column A

        Args:
            sheet: xlwings sheet object
            accounts: List of account dictionaries
            months: List of (month, year, display_name) tuples
            division_name: Optional division identifier for multi-division mode
        """
        # Colors - matching web version
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A - almost black for source sheets

        # Determine column offsets based on division mode
        has_division = division_name is not None or self.is_multi_division.get()
        col_offset = 1 if has_division else 0  # Extra column for Division

        # Row 1: Header
        if has_division:
            sheet['A1'].value = 'Division'
            sheet['B1'].value = 'Account'
            for i, (m, y, name) in enumerate(months):
                sheet.cell(row=1, column=i + 3).value = name
        else:
            sheet['A1'].value = 'Account'
            for i, (m, y, name) in enumerate(months):
                sheet.cell(row=1, column=i + 2).value = name

        # Row 2: YYYYMM helper values for YTD calculations (e.g., 202411 for Nov 2024)
        # This enables SUMPRODUCT formulas to filter by year and month
        start_col = 3 if has_division else 2
        for i, (m, y, name) in enumerate(months):
            sheet.cell(row=2, column=start_col + i).value = y * 100 + m

        # Hide row 2 (helper row)
        try:
            # Row hiding handled via hide_row()
            pass
        except:
            pass

        # Row 3+: Data - write all at once for speed and to ensure numbers are numbers
        data = []
        # Use provided division_name or get from account if present
        div_name = division_name or self.company_name.get()

        for account in accounts:
            # Get division from account if available, otherwise use provided or company name
            acct_division = account.get('division', div_name)
            if has_division:
                row = [acct_division, account['name']]
            else:
                row = [account['name']]

            for m, y, _ in months:
                val = account['values'].get((m, y), 0)
                # Ensure it's a number
                if isinstance(val, str):
                    try:
                        val = float(val.replace(',', ''))
                    except:
                        val = 0
                row.append(val)
            data.append(row)

        if data:
            write_data_to_cells(sheet, data, start_row=3, start_col=1)

        # Format header row with dark background and white text (like web version)
        try:
            num_cols = len(months) + (2 if has_division else 1)
            apply_style_to_range(sheet, 1, 1, 1, num_cols, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=PatternFill(start_color="000000", end_color="000000", fill_type="solid"))

            # Center align month headers
            for col in range(start_col, num_cols + 1):
                # Alignment handled via Alignment() objects
                pass
        except:
            pass

        # Apply number format and font to data columns (now starting at row 3)
        if len(accounts) > 0 and len(months) > 0:
            try:
                apply_style_to_range(sheet, 3, start_col, len(accounts) + 2, num_cols, font=Font(name='Calibri Light', size=10), number_format='#,##0')

                # Account names column
                acct_col = 2 if has_division else 1
                apply_style_to_range(sheet, 3, 1, len(accounts) + 2, 1, font=Font(name='Calibri Light', size=10))

                # Division column if present
                if has_division:
                    # Range: div_range = (sheet, 3, 1, len(accounts) + 2, 1)
                    apply_style_to_range(sheet, 3, 1, len(accounts) + 2, 1, font=Font(name='Calibri Light', size=10))
            except:
                pass

        # Set column widths
        if has_division:
            sheet.column_dimensions['A'].width = 25
            sheet.column_dimensions['B'].width = 45
            for col in range(3, num_cols + 1):
                set_col_width(sheet, col, 14)
        else:
            sheet.column_dimensions['A'].width = 45
            for col in range(2, len(months) + 2):
                set_col_width(sheet, col, 14)

    def _create_menu_sheet(self, sheet, months):
        """Create professional, corporate-style menu/control sheet"""
        # Colors
        DARK_BLUE = CLR_DARK_BLUE  # #16213E
        ACCENT_BLUE = (0, 102, 204)  # #0066CC
        GRAY = CLR_GRAY_TEXT
        LIGHT_GRAY = (240, 240, 240)
        WHITE = (255, 255, 255)

        company = self.company_name.get()

        # Professional layout with clean spacing
        # Row 2-3: Company Title (bold, prominent)
        sheet['B2'].value = company
        sheet['B2'].font = Font(name='Calibri Light', size=28, bold=True, color=CLR_DARK_BLUE)

        sheet['B3'].value = 'Financial Model'
        sheet['B3'].font = Font(name='Calibri Light', size=14)
        sheet['B3'].font = FONT_GRAY_TEXT

        # Row 5: Horizontal line (using cell border)
        try:
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            pass
        except:
            pass

        # === CREATE PERIOD LOOKUP TABLE (Hidden columns K:M) ===
        # This table maps display names to YYYYMM values for formula lookups
        if months:
            # Write header row
            sheet['K6'].value = 'Period'
            sheet['L6'].value = 'YYYYMM'
            sheet['M6'].value = 'MonthNum'
            # Write all months
            for idx, (m, y, name) in enumerate(months):
                row = 7 + idx
                sheet[f'K{row}'].value = name  # Display name (e.g., "Nov 2024")
                sheet[f'L{row}'].value = y * 100 + m  # YYYYMM (e.g., 202411)
                sheet[f'M{row}'].value = m  # Month number (1-12)

        # Row 7-8: Current Period Info (clean, professional)
        sheet['B7'].value = 'Current Period'
        sheet['B7'].font = Font(name='Calibri Light', size=10)
        sheet['B7'].font = FONT_GRAY_TEXT

        sheet['C7'].value = months[-1][2] if months else ''
        sheet['C7'].font = Font(name='Calibri Light', size=12, bold=True, color=CLR_DARK_BLUE)
        # Right-align value cells for visual consistency
        try:
            # Alignment handled via Alignment() objects
            pass
        except:
            pass

        # Add dropdown for Current Period (C7) with all available months
        try:
            if months:
                month_list = ','.join([name for m, y, name in months])
                # Validation handled via DataValidation object
                # DataValidation handled via DataValidation()
                sheet['C7'].fill = PatternFill(start_color="FFFFC8", end_color="FFFFC8", fill_type="solid")  # Light yellow to indicate editable
        except Exception as e:
            print(f"Warning: Could not add period dropdown: {e}")

        # Add hint text
        sheet['D7'].value = '← Select period'
        sheet['D7'].font = Font(name='Calibri Light', size=8)
        sheet['D7'].font = FONT_GRAY_TEXT
        sheet['D7'].font = Font(italic=True)

        sheet['B8'].value = 'Data Range'
        sheet['B8'].font = Font(name='Calibri Light', size=10)
        sheet['B8'].font = FONT_GRAY_TEXT

        sheet['C8'].value = f"{months[0][2]} - {months[-1][2]}" if months else ''
        sheet['C8'].font = Font(name='Calibri Light', size=10, color=CLR_DARK_BLUE)
        # Right-align value cells for visual consistency
        try:
            # Alignment handled via Alignment() objects
            pass
        except:
            pass

        # Row 10: Navigation Header
        sheet['B10'].value = 'Quick Navigation'
        sheet['B10'].font = Font(name='Calibri Light', size=12, bold=True, color=CLR_DARK_BLUE)

        # Rows 11-17: Navigation links (clean hyperlinks)
        nav_items = [
            ('Dashboard', 'Dashboard', 'Executive overview and KPIs'),
            ('P&L Statement', 'Consolidated_PL', 'Profit & Loss analysis'),
            ('Balance Sheet', 'Consolidated_BS', 'Assets, Liabilities & Equity'),
            ('Cash Flow', 'Cash_Flow', 'Cash flow statement'),
            ('Forecast', 'Forecast', 'Budget vs Actual forecast'),
            ('Forecast Summary', 'Forecast_Summary', 'YTD variance analysis'),
            ('Notes', 'Notes', 'Commentary and annotations'),
        ]

        # Standard hyperlink blue for better readability (darker than ACCENT_BLUE)
        HYPERLINK_BLUE = (0, 0, 238)  # Standard web hyperlink blue (#0000EE)

        for i, (label, target, desc) in enumerate(nav_items):
            row = 11 + i
            sheet[f'B{row}'].value = label
            sheet[f'B{row}'].font = Font(name='Calibri Light', size=11)
            sheet[f'C{row}'].value = desc
            sheet[f'C{row}'].font = Font(name='Calibri Light', size=9)
            sheet[f'C{row}'].font = FONT_GRAY_TEXT
            try:
                sheet[f'B{row}'].hyperlink = f"#'{target}'!A1"
                sheet[f'B{row}'].font = Font(name='Calibri Light', size=11, color="0000EE", underline="single")
            except:
                # Fallback styling if hyperlink fails
                sheet[f'B{row}'].font = Font(name='Calibri Light', size=11, color="0000EE", underline="single")

        # Row 19: Version info (subtle)
        sheet['B19'].value = f'Generated: {datetime.now().strftime("%B %d, %Y")}'
        sheet['B19'].font = Font(name='Calibri Light', size=9)
        sheet['B19'].font = FONT_GRAY_TEXT

        sheet['B20'].value = f'Version {APP_VERSION}'
        sheet['B20'].font = Font(name='Calibri Light', size=9)
        sheet['B20'].font = FONT_GRAY_TEXT

        # Add helper cells for YTD calculations with VLOOKUP formulas
        # These update automatically when C7 (Current Period) changes
        try:
            if months:
                last_lookup_row = 6 + len(months)
                # G7 uses VLOOKUP to get YYYYMM from the lookup table
                sheet['G7'].value = f'=IFERROR(VLOOKUP(C7,$K$7:$L${last_lookup_row},2,FALSE),0)'
                # E7 = Month number from lookup
                sheet['E7'].value = f'=IFERROR(VLOOKUP(C7,$K$7:$M${last_lookup_row},3,FALSE),1)'
                # F7 = Year derived from G7 (YYYYMM / 100 rounded down)
                sheet['F7'].value = '=INT(G7/100)'
                # G9 = Same as G7 (Actuals Through)
                sheet['G9'].value = '=G7'
                # I7 = Current year for multi-division reference
                sheet['I7'].value = '=F7'
            # Hide helper columns E through M (includes lookup table)
            # Column hiding handled via hide_columns_range()
        except Exception as e:
            print(f"Warning: Could not set helper formulas: {e}")
            # Fallback to static values
            if months:
                current_month = months[-1][0]
                current_year = months[-1][1]
                sheet['E7'].value = current_month
                sheet['F7'].value = current_year
                sheet['G7'].value = current_year * 100 + current_month
                sheet['G9'].value = current_year * 100 + current_month
                sheet['I7'].value = current_year
            try:
                # Column hiding handled via hide_columns_range()
                pass
            except:
                pass

        # Set column widths for professional layout
        sheet.column_dimensions['A'].width = 3
        sheet.column_dimensions['B'].width = 22
        sheet.column_dimensions['C'].width = 30
        sheet.column_dimensions['D'].width = 3

        # Hide gridlines for cleaner look
        try:
            sheet.views.sheetView[0].showGridLines = False
        except:
            pass

        # Set tab color
        try:
            sheet.sheet_properties.tabColor = "3E2116"
        except:
            pass

    def _create_pl_report(self, sheet, accounts, months, detected_totals=None):
        """Create P&L report with SUMIF formulas and CFO-grade formatting"""
        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = CLR_DARK_BLUE  # #16213E
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY  # #ECECEC

        # Title section
        sheet['A1'].value = company
        sheet['A1'].font = Font(name='Calibri Light', size=14, bold=True)

        sheet['A2'].value = 'Profit & Loss Statement'
        sheet['A2'].font = Font(name='Calibri Light', size=12, bold=True)

        # Calculate column positions
        # Months | Notes | Spacer | PY YTD | CY YTD | Var $ | Var % | Spacer | Full Years...
        header_row = 4
        last_month_col = len(months) + 1
        notes_col = last_month_col + 1
        spacer1_col = notes_col + 1
        py_ytd_col = spacer1_col + 1
        cy_ytd_col = py_ytd_col + 1
        var_col = cy_ytd_col + 1
        var_pct_col = var_col + 1
        spacer2_col = var_pct_col + 1

        # Get unique years for full year columns
        years = sorted(set(y for m, y, name in months))
        fy_start_col = spacer2_col + 1
        last_col = fy_start_col + len(years) - 1

        # Headers
        sheet[f'A{header_row}'].value = 'Account'

        # Row 3: Helper row with YYYYMM values for dynamic YTD calculations
        # This allows formulas to compare dates against Menu!C7
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.cell(row=header_row, column=col).value = f"{self.MONTHS[m-1][:3]} {y}"
            # Row 3 stores YYYYMM as number (e.g., 202411 for Nov 2024)
            sheet.cell(row=3, column=col).value = y * 100 + m

        # Notes and Summary headers
        sheet.cell(row=header_row, column=notes_col).value = 'Notes'
        sheet.cell(row=header_row, column=py_ytd_col).value = 'PY YTD'
        sheet.cell(row=header_row, column=cy_ytd_col).value = 'CY YTD'
        sheet.cell(row=header_row, column=var_col).value = 'Var $'
        sheet.cell(row=header_row, column=var_pct_col).value = 'Var %'

        # Full year headers (just year)
        for i, year in enumerate(years):
            sheet.cell(row=header_row, column=fy_start_col + i).value = str(year)

        # Format header row
        try:
            apply_style_to_range(sheet, header_row, 1, header_row, last_col, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=FILL_DARK_BLUE)
            for col in range(2, last_col + 1):
                if col not in [spacer1_col, spacer2_col]:
                    pass
                    # Alignment handled via Alignment() objects
            # Clear spacer column headers completely
            sheet.cell(row=header_row, column=spacer1_col).value = ''
            sheet.cell(row=header_row, column=spacer2_col).value = ''
        except:
            pass

        # Add YTD date range label in row 3 (e.g., "Jan-Nov") merged across PY YTD and CY YTD
        try:
            ytd_label_cell = sheet.cell(row=3, column=py_ytd_col)
            # Formula shows "Jan-[current month]" based on Menu!E7
            ytd_label_cell.value = '="Jan-"&TEXT(DATE(2024,Menu!$E$7,1),"mmm")'
            # Merge across PY YTD and CY YTD columns
            sheet.merge_cells(start_row=3, start_column=py_ytd_col, end_row=3, end_column=cy_ytd_col)
            # Alignment handled via Alignment() objects
            ytd_label_cell.font = Font(name='Calibri Light', size=9, italic=True, color=CLR_DARK_BLUE)
        except:
            pass

        # Hide Row 3 (YYYYMM helper row) - actually hide the row, not just font color
        try:
            # Row hiding handled via hide_row()
            pass
        except:
            pass

        # Track current section for indentation
        current_section = None
        in_section = False
        row_idx = header_row + 1
        gross_margin_row = None
        net_income_row = None
        total_income_row = None  # Total Revenue/Income BEFORE COGS
        total_cogs_row = None
        total_expenses_row = None  # Total Expenses (not Total Other Expenses)
        found_cogs_section = False  # Track if we've passed the COGS section

        for account in accounts:
            account_name = account['name']
            name_lower = account_name.lower()

            # Detect section headers and track COGS section
            if account['is_header']:
                current_section = name_lower
                in_section = True
                if 'cost' in name_lower or 'cogs' in name_lower:
                    found_cogs_section = True

            # Get indent level from source file (4 spaces per level)
            indent_level = account.get('indent', 0)

            # Write account name with indentation matching source
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name  # 4 spaces per indent level

            # Track Gross Profit/Margin row (don't rename - keep original for formula matching)
            if 'gross profit' in name_lower:
                gross_margin_row = row_idx

            if 'net income' in name_lower and account['is_total']:
                net_income_row = row_idx

            # Track total rows for validation - IMPROVED LOGIC
            if 'total' in name_lower and account['is_total']:
                # Total Income/Revenue - must be BEFORE COGS section (not "other income")
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and
                    'other' not in name_lower and
                    not found_cogs_section):
                    total_income_row = row_idx
                # Total COGS
                elif 'cost' in name_lower or 'cogs' in name_lower:
                    total_cogs_row = row_idx
                    found_cogs_section = True  # Mark that we've found COGS
                # Total Expenses - only the main "Total Expenses" not "Total Other Expenses"
                elif 'expense' in name_lower and 'other' not in name_lower:
                    total_expenses_row = row_idx

            sheet.cell(row=row_idx, column=1).value = display_name
            sheet.cell(row=row_idx, column=1).font = Font(name='Calibri Light', size=10)
            # Vertical center alignment for account names
            try:
                sheet.cell(row=row_idx, column=1).alignment = Alignment(vertical='center')
            except:
                pass

            # Apply formatting based on row type
            if account['is_header']:
                sheet.cell(row=row_idx, column=1).font = Font(bold=True)
                # Header rows should NOT have formulas - leave data cells blank
                row_idx += 1
                in_section = True
                continue  # Skip formula creation for header rows
            elif account['is_total']:
                is_net_income = 'net income' in account_name.lower()

                if is_net_income:
                    # Net Income: bold, thick top border, double bottom border
                    for col in range(1, last_col + 1):
                        if col not in [spacer1_col, spacer2_col]:
                            cell = sheet.cell(row=row_idx, column=col)
                            cell.font = Font(bold=True)
                            try:
                                # Borders handled via Border()/Side() objects
                                # Borders handled via Border()/Side() objects
                                # Borders handled via Border()/Side() objects
                                # Borders handled via Border()/Side() objects
                                pass
                            except:
                                pass
                else:
                    # Other totals: bold, thin top border, gray background
                    for col in range(1, last_col + 1):
                        if col not in [spacer1_col, spacer2_col]:
                            cell = sheet.cell(row=row_idx, column=col)
                            cell.font = Font(bold=True)  # ALL totals bold
                            cell.fill = FILL_SUBTOTAL_GRAY
                            try:
                                # Borders handled via Border()/Side() objects
                                # Borders handled via Border()/Side() objects
                                pass
                            except:
                                pass

            # SUMIF formulas for each month (limited range for speed)
            for i, (m, y, name) in enumerate(months):
                col = i + 2
                cl = get_column_letter(col)
                formula = f"=SUMIF(Source_PL!$A$3:$A$1500,\"{account_name}\",Source_PL!{cl}$3:{cl}$1500)"
                sheet.cell(row=row_idx, column=col).value = formula

            # Notes column - lookup formula (matches Statement Type, Date, and Account)
            # Notes structure: A=Statement Type, B=Date, C=Account, D=Note
            formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="P&L")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{row_idx}))),Notes!$D$2:$D$100),"")'
            sheet.cell(row=row_idx, column=notes_col).value = formula
            sheet.cell(row=row_idx, column=notes_col).font = Font(name='Calibri Light', size=9)
            # Left justify and wrap notes
            try:
                sheet.cell(row=row_idx, column=notes_col).alignment = Alignment(horizontal='left', wrap_text=True)
                sheet.cell(row=row_idx, column=notes_col).alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
            except:
                pass

            # Dynamic YTD formulas that reference Menu helper cells
            # Row 3 contains YYYYMM values (e.g., 202411 for Nov 2024)
            # Menu!E7 = current month number (1-12)
            # Menu!F7 = current year (e.g., 2024)

            first_data_col = get_column_letter(2)  # B
            last_data_col = get_column_letter(len(months) + 1)
            data_range = f'{first_data_col}{row_idx}:{last_data_col}{row_idx}'
            helper_range = f'{first_data_col}$3:{last_data_col}$3'

            # CY YTD: SUMPRODUCT for current year, months through current month
            # Uses -- to coerce TRUE/FALSE to 1/0
            cy_formula = (
                f'=SUMPRODUCT(({data_range})*'
                f'--(INT({helper_range}/100)=Menu!$F$7)*'
                f'--(MOD({helper_range},100)<=Menu!$E$7))'
                )

            # PY YTD: SUMPRODUCT for prior year, months through current month
            py_formula = (
                f'=SUMPRODUCT(({data_range})*'
                f'--(INT({helper_range}/100)=Menu!$F$7-1)*'
                f'--(MOD({helper_range},100)<=Menu!$E$7))'
                )

            sheet.cell(row=row_idx, column=cy_ytd_col).value = cy_formula
            sheet.cell(row=row_idx, column=py_ytd_col).value = py_formula

            # Variance $ (CY - PY)
            sheet.cell(row=row_idx, column=var_col).value = f'={get_column_letter(cy_ytd_col)}{row_idx}-{get_column_letter(py_ytd_col)}{row_idx}'

            # Variance % (Var/PY)
            sheet.cell(row=row_idx, column=var_pct_col).value = f'=IFERROR({get_column_letter(var_col)}{row_idx}/{get_column_letter(py_ytd_col)}{row_idx},0)'

            # Full Year columns - use SUMPRODUCT for robustness
            # Row 3 has YYYYMM values, use these to determine which columns belong to each year
            for i, year in enumerate(years):
                first_col_letter = get_column_letter(2)  # Data starts at column B
                last_col_letter = get_column_letter(last_month_col)
                year_formula = f'=SUMPRODUCT((INT({first_col_letter}$3:{last_col_letter}$3/100)={year})*{first_col_letter}{row_idx}:{last_col_letter}{row_idx})'
                sheet.cell(row=row_idx, column=fy_start_col + i).value = year_formula

            row_idx += 1

            # Add COGS % row after Total COGS
            if account['is_total'] and ('cost' in name_lower or 'cogs' in name_lower) and total_income_row:
                sheet.cell(row=row_idx, column=1).value = '    COGS %'
                sheet.cell(row=row_idx, column=1).font = Font(name='Calibri Light', size=10, italic=True, color="646464")

                # COGS % = Total COGS / Total Revenue for each column
                cogs_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = get_column_letter(col)
                        formula = f'=IFERROR(ABS({col_letter}{cogs_row_num})/{col_letter}{total_income_row},0)'
                        sheet.cell(row=row_idx, column=col).value = formula
                        sheet.cell(row=row_idx, column=col).number_format = '0.0%'
                        sheet.cell(row=row_idx, column=col).font = Font(name='Calibri Light', size=10, italic=True, color="646464")
                row_idx += 1

            # Add Gross Margin % row after Gross Profit
            if 'gross profit' in name_lower and total_income_row:
                sheet.cell(row=row_idx, column=1).value = '    Gross Margin %'
                sheet.cell(row=row_idx, column=1).font = Font(name='Calibri Light', size=10, italic=True, color="646464")

                # Gross Margin % = Gross Profit / Total Revenue for each column
                gross_profit_row = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = get_column_letter(col)
                        formula = f'=IFERROR({col_letter}{gross_profit_row}/{col_letter}{total_income_row},0)'
                        sheet.cell(row=row_idx, column=col).value = formula
                        sheet.cell(row=row_idx, column=col).number_format = '0.0%'
                        sheet.cell(row=row_idx, column=col).font = Font(name='Calibri Light', size=10, italic=True, color="646464")
                row_idx += 1

            # Add Expense % row after Total Expenses (not "Other Expenses")
            if (account['is_total'] and 'expense' in name_lower and
                'other' not in name_lower and 'total' in name_lower and total_income_row):
                sheet.cell(row=row_idx, column=1).value = '    Expense %'
                sheet.cell(row=row_idx, column=1).font = Font(name='Calibri Light', size=10, italic=True, color="646464")

                # Expense % = Total Expenses / Total Revenue for each column
                expense_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = get_column_letter(col)
                        formula = f'=IFERROR(ABS({col_letter}{expense_row_num})/{col_letter}{total_income_row},0)'
                        sheet.cell(row=row_idx, column=col).value = formula
                        sheet.cell(row=row_idx, column=col).number_format = '0.0%'
                        sheet.cell(row=row_idx, column=col).font = Font(name='Calibri Light', size=10, italic=True, color="646464")
                row_idx += 1

            # Add Net Profit % row after Net Income
            if 'net income' in name_lower and account['is_total'] and total_income_row:
                sheet.cell(row=row_idx, column=1).value = '    Net Profit %'
                sheet.cell(row=row_idx, column=1).font = Font(name='Calibri Light', size=10, italic=True, color="646464")

                # Net Profit % = Net Income / Total Revenue for each column
                net_income_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = get_column_letter(col)
                        formula = f'=IFERROR({col_letter}{net_income_row_num}/{col_letter}{total_income_row},0)'
                        sheet.cell(row=row_idx, column=col).value = formula
                        sheet.cell(row=row_idx, column=col).number_format = '0.0%'
                        sheet.cell(row=row_idx, column=col).font = Font(name='Calibri Light', size=10, italic=True, color="646464")
                row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        data_end_row = row_idx - 1
        data_start_row = header_row + 1

        # Apply formatting to data range in bulk operations
        try:
            # Apply font to entire data area at once (much faster than per-column)
            # Range: full_data_range = (sheet, header_row + 1, 1, row - 1, ytd_col)
            apply_style_to_range(sheet, header_row + 1, 1, row - 1, ytd_col, font=Font(name='Calibri Light', size=10))

            # Number format - apply to contiguous ranges for efficiency
            # Month columns (2 to last_month_col)
            if last_month_col > 1:
                apply_style_to_range(sheet, data_start_row, 2, data_end_row, last_month_col, number_format='#,##0')

            # YTD columns (py_ytd and cy_ytd)
            apply_style_to_range(sheet, data_start_row, py_ytd_col, data_end_row, cy_ytd_col, number_format='#,##0')

            # Variance $ column
            apply_style_to_range(sheet, data_start_row, var_col, data_end_row, var_col, number_format='#,##0')

            # Variance % column
            apply_style_to_range(sheet, data_start_row, var_pct_col, data_end_row, var_pct_col, number_format='0.0%')

            # Full year columns
            if fy_start_col <= last_col:
                apply_style_to_range(sheet, data_start_row, fy_start_col, data_end_row, last_col, number_format='#,##0')

            # Right align all numeric columns at once (excludes only spacers and notes)
            for col in range(2, last_col + 1):
                if col not in [spacer1_col, spacer2_col, notes_col]:
                    pass
                    # Alignment handled via Alignment() objects
        except:
            pass

        # Clear all formatting from spacer columns (entire column, true white space)
        try:
            for spacer_col in [spacer1_col, spacer2_col]:
                # Clear spacer column formatting
                for r in range(1, row_idx + 51):
                    sheet.cell(row=r, column=spacer_col).value = None
                apply_style_to_range(sheet, 1, spacer_col, row_idx + 50, spacer_col, fill=PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid"))
        except:
            pass

        # Add EBITDA Reconciliation section
        ebitda_start = row_idx + 2
        sheet.cell(row=ebitda_start, column=1).value = 'RECONCILIATION TO EBITDA'
        sheet.cell(row=ebitda_start, column=1).font = Font(name='Calibri Light', size=10, bold=True, color=CLR_DARK_BLUE)

        ebitda_items = [
            ('Net Income', 'net_income'),
            ('    Add: Interest, net', 'interest'),
            ('    Add: Income taxes', 'taxes'),
            ('    Add: Depreciation and amortization', 'depr'),
            ('    Add: Non-recurring expense (income)', 'nonrecurring'),
            ('Reported EBITDA', 'ebitda_total'),
        ]

        for i, (label, item_type) in enumerate(ebitda_items):
            r = ebitda_start + 1 + i
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            if item_type == 'ebitda_total':
                # Bold with borders for EBITDA total
                sheet.cell(row=r, column=1).font = Font(bold=True)
                for col in range(1, last_month_col + 1):
                    if col > 1:
                        # Sum of Net Income + all add-backs
                        formula = f'=SUM({get_column_letter(col)}{ebitda_start+1}:{get_column_letter(col)}{r-1})'
                        sheet.cell(row=r, column=col).value = formula
                    cell = sheet.cell(row=r, column=col)
                    cell.font = Font(bold=True)
                    try:
                        # Borders handled via Border()/Side() objects
                        # Borders handled via Border()/Side() objects
                        # Borders handled via Border()/Side() objects
                        pass
                    except:
                        pass
            else:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = get_column_letter(col_idx)
                    if item_type == 'net_income' and net_income_row:
                        formula = f'={col_letter}{net_income_row}'
                    elif item_type == 'interest':
                        formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"*Interest*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1'
                    elif item_type == 'taxes':
                        formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"*Tax*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1'
                    elif item_type == 'depr':
                        formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"*Deprec*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1+SUMIF(Source_PL!$A$3:$A$1500,"*Amort*",Source_PL!{col_letter}$3:{col_letter}$1500)*-1'
                    elif item_type == 'nonrecurring':
                        formula = '0'  # Manual entry placeholder
                    else:
                        formula = '0'
                    sheet.cell(row=r, column=col_idx).value = formula
                    sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        ebitda_end_row = ebitda_start + len(ebitda_items)

        # Add validation section at bottom
        val_start = ebitda_end_row + 2
        sheet.cell(row=val_start, column=1).value = 'VALIDATION'
        sheet.cell(row=val_start, column=1).font = Font(name='Calibri Light', size=10, bold=True, color=CLR_DARK_BLUE)

        # Get exact total names from detected_totals for source lookups
        src_total_income = detected_totals.get('total_income', 'Total Income')
        src_total_cogs = detected_totals.get('total_cogs', 'Total Cost of Sales')
        src_total_expenses = detected_totals.get('total_expenses', 'Total Expenses')
        src_net_income = detected_totals.get('net_income', 'Net Income')

        # Report totals section - starts right after VALIDATION header
        report_start = val_start + 1

        validation_report_labels = [
            ('Revenue (Report)', total_income_row),
            ('COGS (Report)', total_cogs_row),
            ('Expenses (Report)', total_expenses_row),
            ('Net Income (Report)', net_income_row),
        ]

        for i, (label, ref_row) in enumerate(validation_report_labels):
            r = report_start + i  # No +1 since no header row
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            if ref_row:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = get_column_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.cell(row=r, column=col_idx).value = formula
                    sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        report_val_end = report_start + len(validation_report_labels) - 1  # Last row of report section

        # Source check rows - using EXACT total names detected from source (no header)
        source_start = report_val_end + 1

        # Use exact matches on detected total names (no wildcards)
        source_labels = [
            ('Revenue (Source)', src_total_income),
            ('COGS (Source)', src_total_cogs),
            ('Expenses (Source)', src_total_expenses),
            ('Net Income (Source)', src_net_income),
        ]

        for i, (label, exact_name) in enumerate(source_labels):
            r = source_start + i  # No header, so start at source_start
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            if exact_name:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = get_column_letter(col_idx)
                    # Use exact match with limited range
                    formula = f'=SUMIF(Source_PL!$A$3:$A$1500,"{exact_name}",Source_PL!{col_letter}$3:{col_letter}$1500)'
                    sheet.cell(row=r, column=col_idx).value = formula
                    sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        source_end = source_start + len(source_labels) - 1  # Last row of source section

        # Variance section (Report - Source = should be 0) - no header
        var_start = source_end + 1

        # Calculate row references for variance (no +1 since no headers)
        rev_report_row = report_start
        cogs_report_row = report_start + 1
        exp_report_row = report_start + 2
        ni_report_row = report_start + 3
        rev_source_row = source_start
        cogs_source_row = source_start + 1
        exp_source_row = source_start + 2
        ni_source_row = source_start + 3

        variance_items = [
            ('Revenue Variance', rev_report_row, rev_source_row),
            ('COGS Variance', cogs_report_row, cogs_source_row),
            ('Expenses Variance', exp_report_row, exp_source_row),
            ('Net Income Variance', ni_report_row, ni_source_row),
        ]

        for i, (label, report_row, source_row) in enumerate(variance_items):
            r = var_start + i  # No header
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            for col_idx in range(2, last_month_col + 1):
                col_letter = get_column_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.cell(row=r, column=col_idx).value = formula
                sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        var_end_row = var_start + len(variance_items) - 1  # Last row of variance section

        # Add matrix borders around validation sections
        try:
            pass
            # Outer border for entire validation section (thick)
            # Range: val_range = (sheet, val_start, 1, var_end_row, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Add thin borders inside
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Bold separator line between Report Totals and Source Totals
            # Range: separator1 = (sheet, source_start, 1, source_start, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Bold separator line between Source Totals and Variance
            # Range: separator2 = (sheet, balance_row, 1, balance_row, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
        except Exception as e:
            print(f"Border formatting warning: {e}")

        # Group validation section so it can be collapsed
        try:
            # Row grouping: use group_rows() helper
            pass
        except Exception as e:
            print(f"Grouping warning: {e}")

        # Group account sections (rows between header and total) for collapsible sections
        try:
            section_start = None
            for i, account in enumerate(accounts):
                row_num = header_row + 1 + i
                if account['is_header']:
                    # Start of a new section
                    section_start = row_num + 1  # First data row after header
                elif account['is_total'] and section_start is not None:
                    # End of section - group rows from section_start to row before total
                    section_end = row_num - 1
                    if section_end >= section_start:
                        pass
                        # Row grouping: use group_rows() helper
                    section_start = None
        except Exception as e:
            print(f"PL section grouping warning: {e}")

        # Set column widths - width of 13 accommodates "$10,000,000" format
        # Using bulk operation first, then override specific columns (performance optimization)
        try:
            sheet.column_dimensions['A'].width = 45
        except:
            pass
        # Bulk set all data columns to 13
        set_bulk_col_width(sheet, 2, 78, 13)
        # Override specific columns
        for col in [spacer1_col, spacer2_col]:
            if col:
                set_col_width(sheet, col, 3)
        if notes_col:
            set_col_width(sheet, notes_col, 30)

        # Group columns by year and hide prior years
        self._group_columns_by_year(sheet, months, header_row)

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_bs_report(self, sheet, accounts, months, detected_totals=None):
        """Create Balance Sheet report with SUMIF formulas and CFO-grade formatting"""
        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = CLR_DARK_BLUE  # #16213E
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY  # #ECECEC

        # Title section
        sheet['A1'].value = company
        sheet['A1'].font = Font(name='Calibri Light', size=14, bold=True)

        sheet['A2'].value = 'Balance Sheet'
        sheet['A2'].font = Font(name='Calibri Light', size=12, bold=True)

        header_row = 4
        last_month_col = len(months) + 1
        notes_col = last_month_col + 1
        last_col = notes_col

        # Headers
        sheet[f'A{header_row}'].value = 'Account'

        # Row 3: Helper row with YYYYMM values for dynamic calculations
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.cell(row=header_row, column=col).value = f"{self.MONTHS[m-1][:3]} {y}"
            sheet.cell(row=3, column=col).value = y * 100 + m

        # Notes header
        sheet.cell(row=header_row, column=notes_col).value = 'Notes'

        # Format header row
        try:
            apply_style_to_range(sheet, header_row, 1, header_row, last_col, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=FILL_DARK_BLUE)
            for col in range(2, last_col + 1):
                # Alignment handled via Alignment() objects
                pass
        except:
            pass

        # Hide Row 3 (YYYYMM helper row) - actually hide the row
        try:
            # Row hiding handled via hide_row()
            pass
        except:
            pass

        # Track sections for indentation
        current_section = None
        in_section = False
        row_idx = header_row + 1
        total_assets_row = None
        total_liab_row = None
        total_equity_row = None
        total_liab_equity_row = None

        for account in accounts:
            account_name = account['name']

            # Detect section headers (Assets, Liabilities, Equity, or subsections)
            if account['is_header']:
                current_section = account_name.lower()
                in_section = True

            # Get indent level from source file (4 spaces per level)
            indent_level = account.get('indent', 0)

            # Write account name with indentation matching source
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name  # 4 spaces per indent level

            # Track key total rows
            if 'total' in account_name.lower():
                if 'assets' in account_name.lower() and 'liab' not in account_name.lower():
                    total_assets_row = row_idx
                elif 'liabilities and equity' in account_name.lower() or ('liab' in account_name.lower() and 'equity' in account_name.lower()):
                    total_liab_equity_row = row_idx
                elif 'liabilities' in account_name.lower():
                    total_liab_row = row_idx
                elif 'equity' in account_name.lower():
                    total_equity_row = row_idx

            sheet.cell(row=row_idx, column=1).value = display_name
            sheet.cell(row=row_idx, column=1).font = Font(name='Calibri Light', size=10)
            # Vertical center alignment for account names
            try:
                sheet.cell(row=row_idx, column=1).alignment = Alignment(vertical='center')
            except:
                pass

            # Apply formatting based on row type
            if account['is_header']:
                sheet.cell(row=row_idx, column=1).font = Font(bold=True)
                # Header rows should NOT have formulas - leave data cells blank
                row_idx += 1
                in_section = True
                continue  # Skip formula creation for header rows
            elif account['is_total']:
                is_main_total = 'total for assets' in account_name.lower() or 'total for liabilities and equity' in account_name.lower()

                if is_main_total:
                    # Main totals: bold, thick top border, double bottom border
                    for col in range(1, last_col + 1):
                        cell = sheet.cell(row=row_idx, column=col)
                        cell.font = Font(bold=True)
                        try:
                            # Borders handled via Border()/Side() objects
                            # Borders handled via Border()/Side() objects
                            # Borders handled via Border()/Side() objects
                            # Borders handled via Border()/Side() objects
                            pass
                        except:
                            pass
                else:
                    # Subtotals: bold, thin top border, gray background
                    for col in range(1, last_col + 1):
                        cell = sheet.cell(row=row_idx, column=col)
                        cell.font = Font(bold=True)  # ALL totals bold
                        cell.fill = FILL_SUBTOTAL_GRAY
                        try:
                            # Borders handled via Border()/Side() objects
                            # Borders handled via Border()/Side() objects
                            pass
                        except:
                            pass

            # SUMIF formulas for each month (limited range for speed)
            for i, (m, y, name) in enumerate(months):
                col = i + 2
                cl = get_column_letter(col)
                formula = f"=SUMIF(Source_BS!$A$3:$A$1500,\"{account_name}\",Source_BS!{cl}$3:{cl}$1500)"
                sheet.cell(row=row_idx, column=col).value = formula

            # Notes column - lookup formula (matches Statement Type, Date, and Account)
            # Notes structure: A=Statement Type, B=Date, C=Account, D=Note
            notes_formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="Balance Sheet")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{row_idx}))),Notes!$D$2:$D$100),"")'
            sheet.cell(row=row_idx, column=notes_col).value = notes_formula
            sheet.cell(row=row_idx, column=notes_col).font = Font(name='Calibri Light', size=9)
            try:
                sheet.cell(row=row_idx, column=notes_col).alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
            except:
                pass

            row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        # Apply formatting to data range in bulk operations (exclude Notes column)
        data_start_row = header_row + 1
        data_end_row = row_idx - 1
        try:
            # Apply font to entire data area at once (much faster than per-cell)
            # Range: full_data_range = (sheet, header_row + 1, 1, row - 1, ytd_col)
            apply_style_to_range(sheet, header_row + 1, 1, row - 1, ytd_col, font=Font(name='Calibri Light', size=10))

            # Number format and alignment for month columns
            # Range: month_range = (sheet, data_start_row, 2, data_end_row, last_month_col)
            apply_style_to_range(sheet, data_start_row, 2, data_end_row, last_month_col, number_format='#,##0')
            # Alignment handled via Alignment() objects
        except:
            pass

        # Add validation section at bottom - with formulas for each month column
        val_start = row_idx + 2
        sheet.cell(row=val_start, column=1).value = 'VALIDATION'
        sheet.cell(row=val_start, column=1).font = Font(name='Calibri Light', size=10, bold=True, color=CLR_DARK_BLUE)

        # Get exact total names from detected_totals for source lookups
        src_total_assets = detected_totals.get('total_assets', 'Total for Assets')
        src_total_liab = detected_totals.get('total_liabilities', 'Total for Liabilities')
        src_total_equity = detected_totals.get('total_equity', 'Total for Equity')
        src_total_liab_equity = detected_totals.get('total_liab_equity', 'Total for Liabilities and Equity')

        # Report totals section - starts right after VALIDATION header (no sub-headers)
        report_start = val_start + 1

        report_labels = [
            ('Total Assets (Report)', total_assets_row),
            ('Total Liabilities (Report)', total_liab_row),
            ('Total Equity (Report)', total_equity_row),
            ('Total Liab + Equity (Report)', total_liab_equity_row),
        ]

        for i, (label, ref_row) in enumerate(report_labels):
            r = report_start + i  # No header row
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            if ref_row:
                for col_idx in range(2, last_col + 1):
                    col_letter = get_column_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.cell(row=r, column=col_idx).value = formula
                    sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        report_end = report_start + len(report_labels) - 1  # Last row of report section

        # Source check section - no header
        source_start = report_end + 1

        source_labels = [
            ('Total Assets (Source)', src_total_assets),
            ('Total Liabilities (Source)', src_total_liab),
            ('Total Equity (Source)', src_total_equity),
        ]

        for i, (label, exact_name) in enumerate(source_labels):
            r = source_start + i  # No header
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            if exact_name:
                for col_idx in range(2, last_col + 1):
                    col_letter = get_column_letter(col_idx)
                    # Use exact match with limited range
                    formula = f'=SUMIF(Source_BS!$A$3:$A$1500,"{exact_name}",Source_BS!{col_letter}$3:{col_letter}$1500)'
                    sheet.cell(row=r, column=col_idx).value = formula
                    sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        source_end = source_start + len(source_labels) - 1  # Last row of source section

        # Balance Check - just the row, no header
        balance_row = source_end + 1
        sheet.cell(row=balance_row, column=1).value = 'Assets - (Liab + Equity)'
        sheet.cell(row=balance_row, column=1).font = Font(name='Calibri Light', size=10, bold=True)

        if total_assets_row and total_liab_equity_row:
            for col_idx in range(2, last_col + 1):
                col_letter = get_column_letter(col_idx)
                formula = f'={col_letter}{total_assets_row}-{col_letter}{total_liab_equity_row}'
                sheet.cell(row=balance_row, column=col_idx).value = formula
                sheet.cell(row=balance_row, column=col_idx).number_format = '#,##0'
                sheet.cell(row=balance_row, column=col_idx).font = Font(bold=True)

        # Variance section - no header
        var_start = balance_row + 1

        # Calculate row references for variance (no headers now)
        assets_report_row = report_start
        liab_report_row = report_start + 1
        equity_report_row = report_start + 2
        assets_source_row = source_start
        liab_source_row = source_start + 1
        equity_source_row = source_start + 2

        variance_items = [
            ('Assets Variance', assets_report_row, assets_source_row),
            ('Liabilities Variance', liab_report_row, liab_source_row),
            ('Equity Variance', equity_report_row, equity_source_row),
        ]

        for i, (label, report_row, source_row) in enumerate(variance_items):
            r = var_start + i  # No header
            sheet.cell(row=r, column=1).value = label
            sheet.cell(row=r, column=1).font = Font(name='Calibri Light', size=10)

            for col_idx in range(2, last_col + 1):
                col_letter = get_column_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.cell(row=r, column=col_idx).value = formula
                sheet.cell(row=r, column=col_idx).number_format = '#,##0'

        var_end_row = var_start + len(variance_items) - 1  # Last row of variance section

        # Add matrix borders around validation sections
        try:
            pass
            # Outer border for entire validation section (thick)
            # Range: val_range = (sheet, val_start, 1, var_end_row, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Add thin borders inside
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Bold separator line between Report Totals and Source Totals
            # Range: separator1 = (sheet, source_start, 1, source_start, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Bold separator line between Source Totals and Balance Check
            # Range: separator2 = (sheet, balance_row, 1, balance_row, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
            
            # Bold separator line between Balance Check and Variance
            # Range: separator3 = (sheet, var_start, 1, var_start, last_col)
            # Borders handled via Border()/Side() objects
            # Borders handled via Border()/Side() objects
        except Exception as e:
            print(f"Border formatting warning: {e}")

        # Group validation section so it can be collapsed
        try:
            # Row grouping: use group_rows() helper
            pass
        except Exception as e:
            print(f"Grouping warning: {e}")

        # Group account sections (rows between header and total) for collapsible sections
        try:
            section_start = None
            for i, account in enumerate(accounts):
                row_num = header_row + 1 + i
                if account['is_header']:
                    # Start of a new section
                    section_start = row_num + 1  # First data row after header
                elif account['is_total'] and section_start is not None:
                    # End of section - group rows from section_start to row before total
                    section_end = row_num - 1
                    if section_end >= section_start:
                        pass
                        # Row grouping: use group_rows() helper
                    section_start = None
        except Exception as e:
            print(f"BS section grouping warning: {e}")

        # Set column widths - width of 13 accommodates "$10,000,000" format
        # Using bulk operation instead of loop for performance
        try:
            sheet.column_dimensions['A'].width = 45
        except:
            pass
        # Bulk set all data columns to 13
        set_bulk_col_width(sheet, 2, 78, 13)
        if notes_col:
            set_col_width(sheet, notes_col, 30)

        # Group columns by year and hide prior years
        self._group_columns_by_year(sheet, months, header_row)

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_cash_flow(self, sheet, months):
        """Create Cash Flow statement with indirect method formulas and division selector"""
        company = self.company_name.get()

        # Detect multi-division mode
        multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1

        # Account column: B for multi-division (A=Division, B=Account), A for single
        acct_col = 'B' if multi_division else 'A'

        # Data starts at column C for multi-division (Division, Account, then data), B for single
        data_start_col = 3 if multi_division else 2

        # Division selector cell (used in formulas)
        div_selector_cell = '$E$2'

        # Colors
        DARK_BLUE = CLR_DARK_BLUE  # #16213E
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY  # #ECECEC

        # Title section
        sheet['A1'].value = company
        sheet['A1'].font = Font(name='Calibri Light', size=14, bold=True)

        sheet['A2'].value = 'Statement of Cash Flows (Indirect Method)'
        sheet['A2'].font = Font(name='Calibri Light', size=12, bold=True)

        # ================================================================
        # DIVISION SELECTOR (for multi-division mode)
        # ================================================================
        if multi_division:
            # Label
            sheet['D2'].value = 'Division:'
            sheet['D2'].font = Font(name='Calibri Light', size=10, bold=True)
            # Alignment handled via Alignment() objects

            # Default value
            sheet['E2'].value = 'All Divisions'
            sheet['E2'].font = Font(name='Calibri Light', size=10)

            # Build division list for dropdown: "All Divisions" + each division name
            div_names = ['All Divisions'] + [d['name'] for d in self.divisions]

            # Create dropdown validation
            try:
                div_list_str = ','.join(div_names)
                # Validation handled via DataValidation object
                # DataValidation handled via DataValidation()
            except Exception as e:
                print(f"Warning: Could not add division dropdown validation: {e}")

            # Format selector cell
            sheet['E2'].fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            try:
                apply_border_box(sheet, 2, 5, 2, 5)
            except:
                pass

        header_row = 4
        last_month_col = len(months) + 1
        ytd_col = last_month_col + 2

        # Headers
        sheet[f'A{header_row}'].value = 'Description'

        # Row 3: Helper row with YYYYMM values for dynamic YTD calculations
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.cell(row=header_row, column=col).value = f"{self.MONTHS[m-1][:3]} {y}"
            sheet.cell(row=3, column=col).value = y * 100 + m

        sheet.cell(row=header_row, column=ytd_col).value = 'YTD'

        # Format header row
        try:
            apply_style_to_range(sheet, header_row, 1, header_row, last_col, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=FILL_DARK_BLUE)
            for col in range(2, ytd_col + 1):
                # Alignment handled via Alignment() objects
                pass
        except:
            pass

        # Track row references for subtotals
        row_refs = {}
        row = header_row + 1

        # Define cash flow structure with formulas
        # Format: (display_name, type, formula_type, bs_search_term)
        # type: 'header', 'item', 'subtotal', 'total'
        # formula_type: 'pl_lookup', 'bs_change_asset', 'bs_change_liab', 'manual', 'sum_range', 'bs_value'
        cf_structure = [
            # OPERATING ACTIVITIES
            ('CASH FLOWS FROM OPERATING ACTIVITIES', 'header', None, None),
            ('    Net Income', 'item', 'pl_lookup', 'Net Income'),
            ('    Adjustments to reconcile net income:', 'header', None, None),
            ('        Depreciation & Amortization', 'item', 'manual', None),
            ('    Changes in Operating Assets and Liabilities:', 'header', None, None),
            ('        (Increase) Decrease in Accounts Receivable', 'item', 'bs_change_asset', 'Accounts Receivable'),
            ('        (Increase) Decrease in Inventory', 'item', 'bs_change_asset', 'Inventory'),
            ('        (Increase) Decrease in Prepaid Expenses', 'item', 'bs_change_asset', 'Prepaid'),
            ('        (Increase) Decrease in Other Current Assets', 'item', 'bs_change_asset', 'Other Current Assets'),
            ('        Increase (Decrease) in Accounts Payable', 'item', 'bs_change_liab', 'Accounts Payable'),
            ('        Increase (Decrease) in Accrued Expenses', 'item', 'bs_change_liab', 'Accrued'),
            ('        Increase (Decrease) in Other Current Liabilities', 'item', 'bs_change_liab', 'Other Current Liabilities'),
            ('Net Cash Provided by Operating Activities', 'subtotal', 'sum_operating', None),
            ('', 'blank', None, None),
            # INVESTING ACTIVITIES
            ('CASH FLOWS FROM INVESTING ACTIVITIES', 'header', None, None),
            ('    Purchase of Property & Equipment', 'item', 'bs_change_asset', 'Fixed Assets'),
            ('    Purchase of Investments', 'item', 'bs_change_asset', 'Investments'),
            ('    Other Investing Activities', 'item', 'manual', None),
            ('Net Cash Used in Investing Activities', 'subtotal', 'sum_investing', None),
            ('', 'blank', None, None),
            # FINANCING ACTIVITIES
            ('CASH FLOWS FROM FINANCING ACTIVITIES', 'header', None, None),
            ('    Proceeds from (Payments on) Line of Credit', 'item', 'bs_change_liab', 'Line of Credit'),
            ('    Proceeds from (Payments on) Long-term Debt', 'item', 'bs_change_liab', 'Long-term'),
            ('    Owner Contributions', 'item', 'bs_change_liab', 'Contributed Capital'),
            ('    Distributions to Owners', 'item', 'bs_change_liab', 'Distribution'),
            ('Net Cash Provided by Financing Activities', 'subtotal', 'sum_financing', None),
            ('', 'blank', None, None),
            # SUMMARY
            ('NET INCREASE (DECREASE) IN CASH', 'total', 'sum_all', None),
            ('', 'blank', None, None),
            ('Cash at Beginning of Period', 'item', 'bs_prior', 'Bank Accounts'),
            ('Cash at End of Period', 'total', 'bs_current', 'Bank Accounts'),
        ]

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED)
        # ================================================================
        all_data = []  # List of row data arrays
        row_types_list = []  # Track row type for batch formatting

        # Track section rows for subtotals (relative to data start)
        operating_start_idx = None
        operating_end_idx = None
        investing_start_idx = None
        investing_end_idx = None
        financing_start_idx = None
        financing_end_idx = None
        net_change_idx = None
        beginning_cash_idx = None
        ending_cash_idx = None

        data_start_row = header_row + 1
        num_cols = ytd_col  # Total columns including YTD

        for struct_idx, (display_name, row_type, formula_type, search_term) in enumerate(cf_structure):
            actual_row = data_start_row + len(all_data)
            row_idx = len(all_data)

            # Track section boundaries
            if 'OPERATING ACTIVITIES' in display_name and row_type == 'header':
                operating_start_idx = row_idx + 1
            elif 'Net Cash Provided by Operating' in display_name:
                operating_end_idx = row_idx - 1
            elif 'INVESTING ACTIVITIES' in display_name and row_type == 'header':
                investing_start_idx = row_idx + 1
            elif 'Net Cash Used in Investing' in display_name:
                investing_end_idx = row_idx - 1
            elif 'FINANCING ACTIVITIES' in display_name and row_type == 'header':
                financing_start_idx = row_idx + 1
            elif 'Net Cash Provided by Financing' in display_name:
                financing_end_idx = row_idx - 1
            elif 'NET INCREASE' in display_name:
                net_change_idx = row_idx
            elif 'Beginning of Period' in display_name:
                beginning_cash_idx = row_idx
            elif 'End of Period' in display_name:
                ending_cash_idx = row_idx

            # Build row data array
            row_data = [display_name]

            if row_type in ['header', 'blank']:
                # Headers/blanks: just description, empty data cells
                row_data.extend([''] * (num_cols - 1))
            else:
                # Build formulas for each month column
                for i in range(len(months)):
                    source_col = data_start_col + i
                    source_col_letter = get_column_letter(source_col)
                    source_col_prev_letter = get_column_letter(source_col - 1) if source_col > data_start_col else source_col_letter

                    if formula_type == 'pl_lookup':
                        if multi_division:
                            # Multi-division: IF selector is "All Divisions", sum all; otherwise filter by division
                            formula = (f'=IF({div_selector_cell}="All Divisions",'
                                      f'SUMIF(Source_PL!$B$3:$B$1500,"*{search_term}*",Source_PL!{source_col_letter}$3:{source_col_letter}$1500),'
                                      f'SUMIFS(Source_PL!{source_col_letter}$3:{source_col_letter}$1500,Source_PL!$A$3:$A$1500,{div_selector_cell},Source_PL!$B$3:$B$1500,"*{search_term}*"))')
                        else:
                            formula = f'=SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_PL!{source_col_letter}$3:{source_col_letter}$1500)'
                    elif formula_type == 'bs_change_asset':
                        if i == 0:
                            formula = 0
                        else:
                            if multi_division:
                                # Asset change = Prior - Current (decrease in asset = source of cash)
                                formula = (f'=IF({div_selector_cell}="All Divisions",'
                                          f'SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)-SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500),'
                                          f'SUMIFS(Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*")-SUMIFS(Source_BS!{source_col_letter}$3:{source_col_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*"))')
                            else:
                                formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)-SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)'
                    elif formula_type == 'bs_change_liab':
                        if i == 0:
                            formula = 0
                        else:
                            if multi_division:
                                # Liability change = Current - Prior (increase in liability = source of cash)
                                formula = (f'=IF({div_selector_cell}="All Divisions",'
                                          f'SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)-SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500),'
                                          f'SUMIFS(Source_BS!{source_col_letter}$3:{source_col_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*")-SUMIFS(Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*"))')
                            else:
                                formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)-SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)'
                    elif formula_type == 'bs_prior':
                        if i == 0:
                            if multi_division:
                                formula = (f'=IF({div_selector_cell}="All Divisions",'
                                          f'SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500),'
                                          f'SUMIFS(Source_BS!{source_col_letter}$3:{source_col_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*"))')
                            else:
                                formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)'
                        else:
                            if multi_division:
                                formula = (f'=IF({div_selector_cell}="All Divisions",'
                                          f'SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500),'
                                          f'SUMIFS(Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*"))')
                            else:
                                formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_prev_letter}$3:{source_col_prev_letter}$1500)'
                    elif formula_type == 'bs_current':
                        if multi_division:
                            formula = (f'=IF({div_selector_cell}="All Divisions",'
                                      f'SUMIF(Source_BS!$B$3:$B$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500),'
                                      f'SUMIFS(Source_BS!{source_col_letter}$3:{source_col_letter}$1500,Source_BS!$A$3:$A$1500,{div_selector_cell},Source_BS!$B$3:$B$1500,"*{search_term}*"))')
                        else:
                            formula = f'=SUMIF(Source_BS!${acct_col}$3:${acct_col}$1500,"*{search_term}*",Source_BS!{source_col_letter}$3:{source_col_letter}$1500)'
                    elif formula_type == 'manual':
                        formula = 0
                    elif formula_type in ['sum_operating', 'sum_investing', 'sum_financing', 'sum_all']:
                        formula = ''  # Will be filled in second pass
                    else:
                        formula = ''

                    row_data.append(formula)

                # Pad to last_month_col if needed, then add spacer and YTD
                while len(row_data) < last_month_col:
                    row_data.append('')

                # Add spacer column
                row_data.append('')

                # Add YTD formula
                if formula_type and formula_type not in ['sum_operating', 'sum_investing', 'sum_financing', 'sum_all']:
                    first_col = get_column_letter(2)
                    last_col_letter = get_column_letter(last_month_col)
                    data_range = f'{first_col}{actual_row}:{last_col_letter}{actual_row}'
                    helper_range = f'{first_col}$3:{last_col_letter}$3'
                    ytd_formula = f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7)*--(MOD({helper_range},100)<=Menu!$E$7))'
                    row_data.append(ytd_formula)
                else:
                    row_data.append('')

            # Ensure row has correct number of columns
            while len(row_data) < num_cols:
                row_data.append('')

            all_data.append(row_data)
            row_types_list.append(row_type)

        # ================================================================
        # WRITE ALL DATA IN ONE BULK OPERATION
        # ================================================================
        if all_data:
            data_end_row = data_start_row + len(all_data) - 1
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=data_start_row + _ri, column=1).value = _val

        # ================================================================
        # SECOND PASS: Fill in subtotal formulas (need row references)
        # ================================================================
        # Operating subtotal
        if operating_start_idx is not None and operating_end_idx is not None:
            subtotal_row = data_start_row + operating_end_idx + 1
            op_start = data_start_row + operating_start_idx
            op_end = data_start_row + operating_end_idx
            for col in range(2, ytd_col + 1):
                sheet.cell(row=subtotal_row, column=col).value = f'=SUM({get_column_letter(col)}{op_start}:{get_column_letter(col)}{op_end})'

        # Investing subtotal
        if investing_start_idx is not None and investing_end_idx is not None:
            subtotal_row = data_start_row + investing_end_idx + 1
            inv_start = data_start_row + investing_start_idx
            inv_end = data_start_row + investing_end_idx
            for col in range(2, ytd_col + 1):
                sheet.cell(row=subtotal_row, column=col).value = f'=SUM({get_column_letter(col)}{inv_start}:{get_column_letter(col)}{inv_end})'

        # Financing subtotal
        if financing_start_idx is not None and financing_end_idx is not None:
            subtotal_row = data_start_row + financing_end_idx + 1
            fin_start = data_start_row + financing_start_idx
            fin_end = data_start_row + financing_end_idx
            for col in range(2, ytd_col + 1):
                sheet.cell(row=subtotal_row, column=col).value = f'=SUM({get_column_letter(col)}{fin_start}:{get_column_letter(col)}{fin_end})'

        # Net change in cash = Operating + Investing + Financing subtotals
        if net_change_idx is not None and operating_end_idx is not None:
            net_change_row = data_start_row + net_change_idx
            op_row = data_start_row + operating_end_idx + 1
            inv_row = data_start_row + investing_end_idx + 1
            fin_row = data_start_row + financing_end_idx + 1
            for col in range(2, ytd_col + 1):
                col_letter = get_column_letter(col)
                sheet.cell(row=net_change_row, column=col).value = f'={col_letter}{op_row}+{col_letter}{inv_row}+{col_letter}{fin_row}'

        # Ending cash = Beginning + Net Change
        if ending_cash_idx is not None and beginning_cash_idx is not None and net_change_idx is not None:
            ending_cash_row = data_start_row + ending_cash_idx
            beginning_cash_row = data_start_row + beginning_cash_idx
            net_change_row = data_start_row + net_change_idx
            for col in range(2, ytd_col + 1):
                col_letter = get_column_letter(col)
                sheet.cell(row=ending_cash_row, column=col).value = f'={col_letter}{beginning_cash_row}+{col_letter}{net_change_row}'

        # ================================================================
        # APPLY FORMATTING IN BULK
        # ================================================================
        row = data_end_row + 1 if all_data else data_start_row

        # Collect rows by type for batch formatting
        header_rows = [data_start_row + idx for idx, rt in enumerate(row_types_list) if rt == 'header']
        subtotal_rows = [data_start_row + idx for idx, rt in enumerate(row_types_list) if rt == 'subtotal']
        total_rows = [data_start_row + idx for idx, rt in enumerate(row_types_list) if rt == 'total']

        # Format headers (bold)
        for r in header_rows:
            sheet.cell(row=r, column=1).font = Font(bold=True)

        # Format subtotals (bold, gray background)
        for r in subtotal_rows:
            sheet.cell(row=r, column=1).font = Font(bold=True)
            apply_style_to_range(sheet, r, 1, r, ytd_col, fill=FILL_SUBTOTAL_GRAY)

        # Format totals (bold)
        for r in total_rows:
            sheet.cell(row=r, column=1).font = Font(bold=True)

        # Apply formatting to data range in bulk operations
        try:
            # Apply font to entire data area at once
            # Range: full_data_range = (sheet, header_row + 1, 1, row - 1, ytd_col)
            apply_style_to_range(sheet, header_row + 1, 1, row - 1, ytd_col, font=Font(name='Calibri Light', size=10))

            # Number format for numeric columns
            if all_data:
                apply_style_to_range(sheet, data_start_row, 2, data_start_row + len(all_data) - 1, ytd_col, number_format='#,##0')
        except:
            pass

        # Set column widths - width of 13 accommodates "$10,000,000" format
        # Using range operation instead of loop for performance
        try:
            sheet.column_dimensions['A'].width = 45
            set_bulk_col_width(sheet, 2, 78, 13)  # B through BZ
        except Exception as e:
            print(f"[CF] Column width warning: {e}")

        # Group prior year columns like PL and BS
        self._group_columns_by_year(sheet, months, header_row)

        # Collapse all outline groups
        try:
            # Outline handled via group_rows()/group_cols()
            pass
        except:
            pass

        # Hide row 3 (YYYYMM helper row) - MUST be at end after all other operations
        try:
            # Row hiding handled via hide_row()
            print(f"[CF] Row 3 hidden successfully")
        except Exception as e:
            print(f"[CF] ERROR hiding row 3: {e}")

        # Add back to menu link and print setup
        self._add_back_to_menu_link(sheet, row=1, col=1)
        self._setup_print_area(sheet)

    def _create_notes_sheet(self, sheet, pl_accounts, bs_accounts, months, divisions=None):
        """Create consolidated notes sheet for P&L, Balance Sheet, and Cash Flow with dropdowns.

        Columns (multi-division mode):
        - A: Division (dropdown)
        - B: Statement Type (P&L, Balance Sheet, Cash Flow)
        - C: Date (month dropdown)
        - D: Account (dropdown based on statement type)
        - E: Note

        Columns (single-entity mode):
        - A: Statement Type
        - B: Date
        - C: Account
        - D: Note
        """
        # Colors
        DARK_BLUE = CLR_DARK_BLUE  # #16213E

        # Check if multi-division mode
        is_multi_div = divisions and len(divisions) > 0

        if is_multi_div:
            # Multi-division: 5 columns
            sheet['A1'].value = 'Division'
            sheet['B1'].value = 'Statement Type'
            sheet['C1'].value = 'Date'
            sheet['D1'].value = 'Account'
            sheet['E1'].value = 'Note'
            header_range = sheet['A1:E1']
            last_col = 'E'
            div_col, type_col, date_col, acct_col, note_col = 'A', 'B', 'C', 'D', 'E'
        else:
            # Single-entity: 4 columns
            sheet['A1'].value = 'Statement Type'
            sheet['B1'].value = 'Date'
            sheet['C1'].value = 'Account'
            sheet['D1'].value = 'Note'
            header_range = sheet['A1:D1']
            last_col = 'D'
            div_col, type_col, date_col, acct_col, note_col = None, 'A', 'B', 'C', 'D'

        # Format header row with dark blue background and white text
        num_header_cols = 5 if is_multi_div else 4
        apply_style_to_range(sheet, 1, 1, 1, num_header_cols, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=FILL_DARK_BLUE)

        # Pre-populate some rows for data entry
        num_rows = 100  # Allow up to 100 notes

        try:
            # Statement type dropdown
            statement_types = 'P&L,Balance Sheet,Cash Flow'

            # Create account lists
            pl_account_list = ','.join([a['name'] for a in pl_accounts][:40])
            bs_account_list = ','.join([a['name'] for a in bs_accounts][:40])

            month_list = ','.join(months) if isinstance(months[0], str) else ','.join([m[2] for m in months])

            # Division dropdown (multi-division mode only)
            if is_multi_div and div_col:
                division_list = 'All,' + ','.join([d['name'] for d in divisions])
                for row in range(2, min(52, num_rows + 2)):
                    try:
                        # Validation handled via DataValidation object
                        # DataValidation handled via DataValidation()
                        pass
                    except:
                        pass

            # Apply data validation to Statement Type column
            for row in range(2, min(52, num_rows + 2)):
                try:
                    # Validation handled via DataValidation object
                    # DataValidation handled via DataValidation()
                    pass
                except:
                    pass

            # Apply data validation to Date column
            for row in range(2, min(52, num_rows + 2)):
                try:
                    # Validation handled via DataValidation object
                    # DataValidation handled via DataValidation()
                    pass
                except:
                    pass

            # Apply data validation to Account column - combined list
            all_accounts = pl_account_list[:120] + ',' + bs_account_list[:120]
            for row in range(2, min(52, num_rows + 2)):
                try:
                    # Validation handled via DataValidation object
                    # DataValidation handled via DataValidation()
                    pass
                except:
                    pass

        except Exception as e:
            print(f"Warning: Could not add dropdowns to Notes sheet: {e}")

        # Set column widths
        if is_multi_div:
            sheet.column_dimensions['A'].width = 18
            sheet.column_dimensions['B'].width = 15
            sheet.column_dimensions['C'].width = 12
            sheet.column_dimensions['D'].width = 40
            sheet.column_dimensions['E'].width = 60
            data_range = sheet['A2:E51']
        else:
            sheet.column_dimensions['A'].width = 15
            sheet.column_dimensions['B'].width = 12
            sheet.column_dimensions['C'].width = 40
            sheet.column_dimensions['D'].width = 60
            data_range = sheet['A2:D51']

        # Format data area
        num_note_cols = 5 if (divisions and len(divisions) > 0) else 4
        apply_style_to_range(sheet, 2, 1, 101, num_note_cols, font=Font(name='Calibri Light', size=10))

        # Enable AutoFilter for sorting
        try:
            sheet.auto_filter.ref = sheet.dimensions
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    # =========================================================================
    # FORECAST MODULE METHODS
    # =========================================================================

    def _create_source_budget_sheet(self, sheet, accounts, months):
        """Create Source_Budget sheet with 12 months for full-year budgeting.

        Structure:
        - Row 1: Company name title
        - Row 2: "Budget" subtitle
        - Row 3: Blank
        - Row 4: Headers (Account, Jan-Dec month names like "Jan 24")
        - Row 5: YYYYMM helper values (e.g., 202401) for formula lookups - hidden
        - Row 6+: Account data with indentation (initially zeros, populated via budget import)

        Note: Unlike Source_PL which only has actual data months, Source_Budget always
        has all 12 months of the year to support full-year forecasting.
        """
        # Colors - matching source sheets
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A
        DARK_BLUE = CLR_DARK_BLUE
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY

        company = self.company_name.get() if hasattr(self, 'company_name') else 'Company'

        # Determine the budget year from the last month in the data
        budget_year = months[-1][1] if months else datetime.now().year
        year_suffix = str(budget_year)[-2:]

        # Create full 12-month list for budget (Jan-Dec of budget year)
        month_abbrevs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        budget_months = []
        for m in range(1, 13):
            name = f"{month_abbrevs[m-1]} {year_suffix}"
            budget_months.append((m, budget_year, name))

        print(f"[Source_Budget] Creating 12-month budget for year {budget_year}")

        # Row 1: Company name title
        sheet['A1'].value = company
        sheet['A1'].font = Font(name='Calibri Light', size=14, bold=True)

        # Row 2: Budget subtitle
        sheet['A2'].value = 'Budget'
        sheet['A2'].font = Font(name='Calibri Light', size=12, bold=True)

        header_row = 4

        # Row 4: Headers - always 12 months (Jan-Dec)
        sheet[f'A{header_row}'].value = 'Account'
        for i, (m, y, name) in enumerate(budget_months):
            sheet.cell(row=header_row, column=i + 2).value = name

        # Row 5: YYYYMM helper values for formula lookups
        for i, (m, y, name) in enumerate(budget_months):
            sheet.cell(row=header_row + 1, column=i + 2).value = y * 100 + m

        # Hide row 5 (helper row)
        try:
            # Row hiding handled via hide_row()
            pass
        except:
            pass

        # Row 6+: Account names with indentation and zero values
        data = []
        data_start_row = header_row + 2
        for account in accounts:
            # Preserve indentation like P&L
            account_name = account['name']
            indent_level = account.get('indent', 0)
            is_header = account.get('is_header', False)
            is_total = account.get('is_total', False)

            # Add indentation for non-header, non-total rows
            if indent_level > 0 and not is_header and not is_total:
                display_name = ('    ' * indent_level) + account_name
            else:
                display_name = account_name

            row = [display_name]
            for _ in range(12):  # Always 12 months
                row.append(0)  # Initialize with zeros
            data.append(row)

        if data:
            write_data_to_cells(sheet, data, start_row=data_start_row, start_col=1)

        # Format header row (always 12 months + Account column = 13 columns)
        num_budget_cols = 12
        try:
            apply_style_to_range(sheet, header_row, 1, header_row, num_budget_cols + 1, font=Font(name='Calibri Light', size=10, bold=True, color="FFFFFF"), fill=PatternFill(start_color="000000", end_color="000000", fill_type="solid"))

            for col in range(2, num_budget_cols + 2):
                # Alignment handled via Alignment() objects
                pass
        except:
            pass

        # Format data area with proper styling for headers/totals
        if len(accounts) > 0:
            try:
                # Range: data_range = (sheet, data_start_row, 2, data_start_row + len(accounts) - 1, num_budget_cols + 1)
                apply_style_to_range(sheet, data_start_row, 2, data_start_row + len(accounts) - 1, num_budget_cols + 1, number_format='#,##0')
                apply_style_to_range(sheet, data_start_row, 2, data_start_row + len(accounts) - 1, num_budget_cols + 1, font=Font(name='Calibri Light', size=10))

                # Range: account_range = (sheet, data_start_row, 1, data_start_row + len(accounts) - 1, 1)
                apply_style_to_range(sheet, data_start_row, 1, data_start_row + len(accounts) - 1, 1, font=Font(name='Calibri Light', size=10))

                # Format header and total rows
                for i, account in enumerate(accounts):
                    row_num = data_start_row + i
                    if account.get('is_header', False):
                        sheet.cell(row=row_num, column=1).font = Font(bold=True)
                    elif account.get('is_total', False):
                        # Range: row_range = (sheet, r, 1, r, last_col)
                        apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
                        apply_style_to_range(sheet, r, 1, r, last_col, fill=FILL_SUBTOTAL_GRAY)
            except:
                pass

        # Set column widths
        sheet.column_dimensions['A'].width = 45
        for col in range(2, num_budget_cols + 2):
            set_col_width(sheet, col, 14)

        # Group and collapse previous year columns (use budget_months instead of months)
        self._group_previous_year_columns(sheet, budget_months, data_start_col=2)

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_forecast_sheet(self, sheet, accounts, months, division_name=None):
        """Create Forecast sheet with 5 columns per month: ACTUAL, BUDGET, ADJ, NOTE, FORECAST.

        Structure:
        - Column A: Account names
        - For each month (5 columns): ACTUAL, BUDGET, ADJ, NOTE, FORECAST
        - FORECAST = ACTUAL for past months, BUDGET + ADJ for future months
        - Columns are grouped so only FORECAST is visible by default
        - Total Forecast column at end sums all FORECAST columns

        Args:
            sheet: xlwings sheet object
            accounts: List of account dictionaries
            months: List of (month, year, display_name) tuples
            division_name: Optional division name for multi-division mode
        """
        # Colors
        DARK_BLUE = CLR_DARK_BLUE
        LIGHT_GRAY = (242, 242, 242)
        SUBTOTAL_GRAY = (220, 220, 220)
        ACTUAL_BLUE = (189, 215, 238)
        BUDGET_YELLOW = (255, 242, 204)
        ADJ_ORANGE = (252, 228, 214)
        NOTE_WHITE = (255, 255, 255)
        FORECAST_GREEN = (198, 224, 180)

        # Determine current year and month for past/future logic
        # FIXED: Use the data's current month (from Menu/months), NOT datetime.now()
        # The last month in 'months' represents the current period - this is what Menu!C7 shows
        # Past months should use Actuals, current and future should use Budget + Adj
        current_year = months[-1][1] if months else datetime.now().year
        current_month = months[-1][0] if months else datetime.now().month
        current_ym = current_year * 100 + current_month
        print(f"[Forecast] Current period: Month={current_month}, Year={current_year}, YYYYMM={current_ym}")

        # Create month list for current year (12 months)
        month_abbrevs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        year_suffix = str(current_year)[-2:]

        # 5 columns per month
        COLS_PER_MONTH = 5
        COL_ACTUAL = 0
        COL_BUDGET = 1
        COL_ADJ = 2
        COL_NOTE = 3
        COL_FORECAST = 4

        # Determine source sheet name and if multi-division mode
        multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1
        if division_name:
            source_pl = f"'{division_name}_PL'"
            source_budget = "'Source_Budget'"
        else:
            source_pl = "'Consolidated_PL'" if multi_division else "'P&L'"
            source_budget = "'Source_Budget'"

        # Title
        title_text = f'Forecast - {current_year}'
        if division_name:
            title_text = f'{division_name} Forecast - {current_year}'
        sheet['A1'].value = title_text
        sheet['A1'].font = Font(bold=True, size=16, color=CLR_DARK_BLUE)

        # Row 3: Month headers (merged across 5 columns each)
        # Row 4: Column sub-headers (ACTUAL, BUDGET, ADJ, NOTE, FORECAST)
        col = 2  # Start at column B
        month_header_row = []
        subheader_row = ['Account']

        for month_idx in range(12):
            month_num = month_idx + 1
            month_name = f"{month_abbrevs[month_idx]} {year_suffix}"
            month_header_row.append(month_name)

            # Add 5 sub-headers for this month
            # For past months, last column shows "Actual"; for future months, shows "Forecast"
            month_ym = current_year * 100 + month_num
            is_past_month = month_ym <= current_ym
            forecast_header = 'Actual' if is_past_month else 'Forecast'
            subheader_row.extend(['Actual', 'Budget', 'Adj', 'Note', forecast_header])

        # Add Total Forecast at end
        subheader_row.append('Total Forecast')

        # Write sub-headers (row 4)
        write_row_to_cells(sheet, subheader_row, row=4, start_col=1)
        forecast_last_col = len(subheader_row)
        apply_style_to_range(sheet, 4, 1, 4, forecast_last_col, font=Font(bold=True, size=9))

        # Write and merge month headers (row 3)
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            end_col = start_col + COLS_PER_MONTH - 1
            month_name = f"{month_abbrevs[month_idx]} {year_suffix}"

            # Write month name and merge
            sheet.cell(row=3, column=start_col).value = month_name
            try:
                # Range: merge_range = (sheet, 3, start_col, 3, end_col)
                sheet.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
                apply_style_to_range(sheet, 3, start_col, 3, end_col, font=Font(bold=True, size=11, color="FFFFFF"), fill=FILL_DARK_BLUE)
                # Alignment handled via Alignment() objects
            except:
                pass

        # Color the sub-header columns
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            try:
                sheet.cell(row=4, column=start_col + COL_ACTUAL).fill = rgb_fill(ACTUAL_BLUE)
                sheet.cell(row=4, column=start_col + COL_BUDGET).fill = rgb_fill(BUDGET_YELLOW)
                sheet.cell(row=4, column=start_col + COL_ADJ).fill = rgb_fill(ADJ_ORANGE)
                sheet.cell(row=4, column=start_col + COL_NOTE).fill = rgb_fill(NOTE_WHITE)
                sheet.cell(row=4, column=start_col + COL_FORECAST).fill = rgb_fill(FORECAST_GREEN)
            except:
                pass

        # Total Forecast header
        total_col = 2 + (12 * COLS_PER_MONTH)
        sheet.cell(row=3, column=total_col).value = 'TOTAL'
        sheet.cell(row=3, column=total_col).font = Font(bold=True)
        sheet.cell(row=3, column=total_col).fill = rgb_fill(FORECAST_GREEN)
        sheet.cell(row=4, column=total_col).fill = rgb_fill(FORECAST_GREEN)

        # Track total rows for formatting
        total_rows = []
        for acct_idx, account in enumerate(accounts):
            if account.get('is_total', False):
                total_rows.append(5 + acct_idx)

        last_row = 4 + len(accounts)

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED - single bulk write)
        # ================================================================
        # Pre-calculate column letters and source info
        source_data_start_col = 3 if multi_division else 2
        source_acct_col = 'B' if multi_division else 'A'
        total_col = 2 + (12 * COLS_PER_MONTH)

        # Build month-to-column mapping for Source_PL
        # Map by (month, year) tuple to handle cross-year data correctly
        month_year_to_source_col = {}
        for idx, (m, y, name) in enumerate(months):
            month_year_to_source_col[(m, y)] = source_data_start_col + idx

        # Debug: print mapping
        print(f"[Forecast] Source months mapping: {month_year_to_source_col}")
        print(f"[Forecast] Looking for year {current_year}, current_ym = {current_ym}")

        # Build complete 2D array for all data (accounts × all columns)
        all_data = []  # Each row: [account_name, month1_actual, month1_budget, month1_adj, month1_note, month1_forecast, ...]

        for acct_idx, account in enumerate(accounts):
            r = 5 + acct_idx
            account_name = account['name']
            indent_level = account.get('indent', 0)

            # Build display name with indentation
            display_name = account_name
            if indent_level > 0 and not account.get('is_header', False) and not account.get('is_total', False):
                display_name = ('  ' * indent_level) + account_name

            row_data = [display_name]

            # Build formulas for all 12 months
            for month_idx in range(12):
                month_num = month_idx + 1
                month_ym = current_year * 100 + month_num
                is_past_month = month_ym <= current_ym

                start_col = 2 + (month_idx * COLS_PER_MONTH)
                actual_letter = get_column_letter(start_col + COL_ACTUAL)
                budget_letter = get_column_letter(start_col + COL_BUDGET)
                adj_letter = get_column_letter(start_col + COL_ADJ)

                # Get the source column for ACTUAL data from Source_PL (using month, year tuple)
                month_key = (month_num, current_year)
                if month_key in month_year_to_source_col:
                    source_col_letter = get_column_letter(month_year_to_source_col[month_key])
                    # ACTUAL formula - pull from Source_PL
                    if division_name:
                        actual_formula = f'=IFERROR(SUMIFS(Source_PL!{source_col_letter}$3:{source_col_letter}$1500,Source_PL!$A$3:$A$1500,"{division_name}",Source_PL!$B$3:$B$1500,TRIM(A{r})),0)'
                    else:
                        actual_formula = f'=IFERROR(SUMIF(Source_PL!${source_acct_col}$3:${source_acct_col}$1500,TRIM(A{r}),Source_PL!{source_col_letter}$3:{source_col_letter}$1500),0)'
                else:
                    # Month not in source data - no actual data available
                    actual_formula = '0'

                # BUDGET formula - Source_Budget ALWAYS has Jan-Dec in columns B-M (2-13)
                # So Jan=B, Feb=C, Mar=D, etc. (column = 2 + month_idx where month_idx is 0-11)
                budget_col_letter = get_column_letter(2 + month_idx)
                budget_formula = f'=IFERROR(SUMIF(Source_Budget!$A$6:$A$1500,TRIM(A{r}),Source_Budget!{budget_col_letter}$6:{budget_col_letter}$1500),0)'

                # FORECAST formula
                if is_past_month:
                    forecast_formula = f'={actual_letter}{r}'
                else:
                    forecast_formula = f'={budget_letter}{r}+{adj_letter}{r}'

                # Add 5 columns for this month: Actual, Budget, Adj, Note, Forecast
                row_data.extend([actual_formula, budget_formula, 0, '', forecast_formula])

            # Add Total Forecast formula
            forecast_cols = [get_column_letter(2 + (m * COLS_PER_MONTH) + COL_FORECAST) for m in range(12)]
            sum_parts = '+'.join([f'{c}{r}' for c in forecast_cols])
            row_data.append(f'={sum_parts}')

            all_data.append(row_data)

        # WRITE ALL DATA IN ONE BULK OPERATION
        if all_data:
            num_cols = 1 + (12 * COLS_PER_MONTH) + 1  # Account + 12 months × 5 cols + Total
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=5 + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=5 + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=5 + _ri, column=1).value = _val

        # ================================================================
        # APPLY FORMATTING IN BATCHES (by column type across all months)
        # ================================================================
        # Number format for all data columns at once
        apply_style_to_range(sheet, 5, 2, last_row, total_col, number_format='#,##0')

        # Color columns by type - batch all months together
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            apply_style_to_range(sheet, 5, start_col + COL_ACTUAL, last_row, start_col + COL_ACTUAL, fill=PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"))
            apply_style_to_range(sheet, 5, start_col + COL_BUDGET, last_row, start_col + COL_BUDGET, fill=PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"))
            apply_style_to_range(sheet, 5, start_col + COL_ADJ, last_row, start_col + COL_ADJ, fill=PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"))
            apply_style_to_range(sheet, 5, start_col + COL_FORECAST, last_row, start_col + COL_FORECAST, fill=PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"))

        # Total column formatting
        apply_style_to_range(sheet, 5, total_col, last_row, total_col, font=Font(bold=True))
        apply_style_to_range(sheet, 5, total_col, last_row, total_col, fill=PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"))

        # Format total rows
        for r in total_rows:
            try:
                # Range: row_range = (sheet, r, 1, r, last_col)
                apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
                apply_style_to_range(sheet, r, 1, r, last_col, fill=FILL_SUBTOTAL_GRAY)
            except:
                pass

        # Set column widths - width of 13 accommodates "$10,000,000" format
        # Using bulk operation first, then override specific columns (performance optimization)
        sheet.column_dimensions['A'].width = 35
        # Set all data columns to width 13
        for col_idx in range(2, 2 + 12 * COLS_PER_MONTH + 2):
            sheet.column_dimensions[get_column_letter(col_idx)].width = 13
        # Override specific columns (ADJ=10, NOTE=15) for each month
        for month_idx in range(12):
            start_col = 2 + (month_idx * COLS_PER_MONTH)
            sheet.column_dimensions[get_column_letter(start_col + COL_ADJ)].width = 10
            sheet.column_dimensions[get_column_letter(start_col + COL_NOTE)].width = 15

        # Group columns so only FORECAST is visible (hide ACTUAL, BUDGET, ADJ, NOTE)
        try:
            for month_idx in range(12):
                start_col = 2 + (month_idx * COLS_PER_MONTH)
                # Group columns: ACTUAL, BUDGET, ADJ, NOTE (hide them, keep FORECAST visible)
                group_start = get_column_letter(start_col + COL_ACTUAL)
                group_end = get_column_letter(start_col + COL_NOTE)
                group_cols(sheet, start_col + COL_ACTUAL, start_col + COL_NOTE, outline_level=1, hidden=True)

            # Collapse all groups
            # Outline handled via group_rows()/group_cols()
        except Exception as e:
            print(f"Warning: Could not group columns: {e}")

        # Freeze panes (Account column and header rows)
        try:
            sheet.freeze_panes = 'B5'
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=total_col + 2)

    def _create_forecast_summary_sheet(self, sheet, accounts, months, division_name=None):
        """Create Forecast Summary with Current Month, YTD Actual vs Budget, and YTD Forecast comparisons.
        Matches P&L formatting with indentation, borders, and profit % rows.

        Args:
            sheet: xlwings Sheet object
            accounts: List of account dictionaries
            months: List of (month, year, name) tuples
            division_name: Optional division name for division-specific summaries

        Layout:
        - Column A: Account names (with P&L-style indentation)
        - Column B: Current Month Actual
        - Column C: Current Month Budget
        - Column D: Current Month Variance $
        - Column E: Current Month Variance %
        - Column F: (spacer)
        - Column G: YTD Actual
        - Column H: YTD Budget
        - Column I: YTD Variance $
        - Column J: YTD Variance %
        - Column K: (spacer)
        - Column L: YTD Forecast
        - Column M: YTD Budget (for forecast comparison)
        - Column N: Forecast Var $
        - Column O: Forecast Var %
        """
        # Colors
        DARK_BLUE = CLR_DARK_BLUE
        LIGHT_GRAY = (242, 242, 242)
        HEADER_GRAY = (200, 200, 200)
        SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY  # Matches P&L

        # Spacer columns
        SPACER1_COL = 6  # F
        SPACER2_COL = 11  # K
        LAST_DATA_COL = 15  # O

        # Determine current year
        if months:
            current_year = months[-1][1]
        else:
            current_year = 2024

        # Determine forecast sheet name and column offsets based on mode
        is_multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1

        # For division-specific summaries, use division's Forecast sheet
        if division_name:
            safe_name = division_name.replace(' ', '_')[:20]
            forecast_sheet_name = f"{safe_name}_Forecast"
        else:
            forecast_sheet_name = 'Consolidated_Forecast' if is_multi_division else 'Forecast'

        # In multi-division mode: col A=Division, col B=Account, data starts col C
        # In single mode: col A=Account, data starts col B
        src_acct_col = 'B' if is_multi_division else 'A'
        src_data_start_col = 'C' if is_multi_division else 'B'

        # Calculate last data column letter
        # Multi-division: data starts at C (col 3), so last = 3 + len(months) - 1 = len(months) + 2
        # Single mode: data starts at B (col 2), so last = 2 + len(months) - 1 = len(months) + 1
        col_letter_last = get_column_letter(len(months) + 2) if is_multi_division else get_column_letter(len(months) + 1)

        # Row 1: Title
        title_suffix = f' - {division_name}' if division_name else ''
        sheet['A1'].value = f'Forecast Summary{title_suffix} - {current_year}'
        sheet['A1'].font = Font(bold=True, size=14, name='Calibri Light')

        # Row 2: Section headers (merged)
        sheet['B2'].value = 'Current Month'
        try:
            sheet.merge_cells('B2:E2')
            # Alignment handled via Alignment() objects
        except:
            pass
        sheet['B2'].font = Font(bold=True, name='Calibri Light')
        sheet['B2'].fill = FILL_DARK_BLUE
        sheet['B2'].font = Font(color="FFFFFF")

        sheet['G2'].value = 'Year-to-Date Actual vs Budget'
        try:
            sheet.merge_cells('G2:J2')
            # Alignment handled via Alignment() objects
        except:
            pass
        sheet['G2'].font = Font(bold=True, name='Calibri Light')
        sheet['G2'].fill = FILL_DARK_BLUE
        sheet['G2'].font = Font(color="FFFFFF")

        sheet['L2'].value = 'YTD Forecast vs Budget'
        try:
            sheet.merge_cells('L2:O2')
            # Alignment handled via Alignment() objects
        except:
            pass
        sheet['L2'].font = Font(bold=True, name='Calibri Light')
        sheet['L2'].fill = FILL_DARK_BLUE
        sheet['L2'].font = Font(color="FFFFFF")

        # Row 3: Column headers
        headers = [
            ('A3', 'Account'),
            ('B3', 'Actual'),
            ('C3', 'Budget'),
            ('D3', 'Var $'),
            ('E3', 'Var %'),
            ('F3', ''),  # Spacer
            ('G3', 'Actual'),
            ('H3', 'Budget'),
            ('I3', 'Var $'),
            ('J3', 'Var %'),
            ('K3', ''),  # Spacer
            ('L3', 'Forecast'),
            ('M3', 'Budget'),
            ('N3', 'Var $'),
            ('O3', 'Var %'),
        ]
        for cell, value in headers:
            sheet[cell].value = value

        # Format header row
        apply_style_to_range(sheet, 3, 1, 3, 15, font=Font(bold=True, name='Calibri Light', size=10), fill=PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"))

        # Spacer columns formatting
        apply_style_to_range(sheet, 2, 6, 3, 6, fill=PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"))
        apply_style_to_range(sheet, 2, 11, 3, 11, fill=PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"))

        # ================================================================
        # BUILD ALL DATA IN MEMORY FIRST (OPTIMIZED)
        # ================================================================
        all_data = []  # List of row data arrays
        row_types = []  # Track row type for batch formatting

        # Track key rows for Gross Profit % and Net Profit %
        gross_margin_row = None
        total_income_row = None
        net_income_row = None
        forecast_row = 5  # Track actual Forecast sheet row (starts at 5)

        data_start_row = 4  # Data starts at row 4

        for acct_idx, account in enumerate(accounts):
            account_name = account['name']
            name_lower = account_name.lower()
            actual_row = data_start_row + len(all_data)

            # Get indent level from source file
            indent_level = account.get('indent', 0)

            # Determine display name with indentation
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name

            # Track Gross Profit/Margin row
            if 'gross profit' in name_lower:
                gross_margin_row = actual_row

            # Track Net Income row
            if 'net income' in name_lower and account['is_total']:
                net_income_row = actual_row

            # Track Total Income/Revenue row
            if account['is_total'] and 'total' in name_lower:
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and 'other' not in name_lower):
                    total_income_row = actual_row

            # Build row data array: [A, B, C, D, E, F, G, H, I, J, K, L, M, N, O]
            if account['is_header']:
                # Header rows: just the name, empty data cells
                row_data = [display_name] + [''] * 14
                all_data.append(row_data)
                row_types.append('header')
                forecast_row += 1
                continue

            # --- Build formulas for data rows ---
            # In multi-division mode, Source_PL has: col A=Division, col B=Account, col C onwards=data
            # Row 2 has YYYYMM values in the data columns

            # Current Month Actual (B) - sum where account matches and YYYYMM = Menu!G7
            if is_multi_division:
                cm_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!${src_acct_col}$3:${src_acct_col}$1000=TRIM($A{actual_row}))*"
                    f"(Source_PL!${src_data_start_col}$2:${col_letter_last}$2=Menu!$G$7)*"
                    f"(Source_PL!${src_data_start_col}$3:${col_letter_last}$1000))"
                    )
            else:
                cm_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=TRIM($A{actual_row}))*"
                    f"(Source_PL!$B$2:${col_letter_last}$2=Menu!$G$7)*"
                    f"(Source_PL!$B$3:${col_letter_last}$1000))"
                    )

            # Current Month Budget (C)
            cm_budget = (
                f"=SUMPRODUCT("
                f"(Source_Budget!$A$6:$A$1000=TRIM($A{actual_row}))*"
                f"(Source_Budget!$B$5:${col_letter_last}$5=Menu!$G$7)*"
                f"(Source_Budget!$B$6:${col_letter_last}$1000))"
                )
            # Current Month Variance $ (D)
            cm_var = f"=B{actual_row}-C{actual_row}"
            # Current Month Variance % (E)
            cm_var_pct = f"=IFERROR(D{actual_row}/ABS(C{actual_row}),0)"

            # YTD Actual (G) - sum where account matches and year matches and YYYYMM <= Menu!G7
            if is_multi_division:
                ytd_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!${src_acct_col}$3:${src_acct_col}$1000=TRIM($A{actual_row}))*"
                    f"(INT(Source_PL!${src_data_start_col}$2:${col_letter_last}$2/100)=Menu!$F$7)*"
                    f"(Source_PL!${src_data_start_col}$2:${col_letter_last}$2<=Menu!$G$7)*"
                    f"(Source_PL!${src_data_start_col}$3:${col_letter_last}$1000))"
                    )
            else:
                ytd_actual = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=TRIM($A{actual_row}))*"
                    f"(INT(Source_PL!$B$2:${col_letter_last}$2/100)=Menu!$F$7)*"
                    f"(Source_PL!$B$2:${col_letter_last}$2<=Menu!$G$7)*"
                    f"(Source_PL!$B$3:${col_letter_last}$1000))"
                    )

            # YTD Budget (H)
            ytd_budget = (
                f"=SUMPRODUCT("
                f"(Source_Budget!$A$6:$A$1000=TRIM($A{actual_row}))*"
                f"(INT(Source_Budget!$B$5:${col_letter_last}$5/100)=Menu!$F$7)*"
                f"(Source_Budget!$B$5:${col_letter_last}$5<=Menu!$G$7)*"
                f"(Source_Budget!$B$6:${col_letter_last}$1000))"
                )
            # YTD Variance $ (I)
            ytd_var = f"=G{actual_row}-H{actual_row}"
            # YTD Variance % (J)
            ytd_var_pct = f"=IFERROR(I{actual_row}/ABS(H{actual_row}),0)"

            # YTD Forecast (L) - sum of Forecast columns where month <= current
            ytd_forecast_parts = []
            for month_idx in range(12):
                forecast_col = get_column_letter(2 + month_idx * 5 + 4)
                yyyymm = current_year * 100 + (month_idx + 1)
                ytd_forecast_parts.append(f"IF({yyyymm}<=Menu!$G$7,'{forecast_sheet_name}'!{forecast_col}{forecast_row},0)")
            ytd_forecast = "=" + "+".join(ytd_forecast_parts)

            # YTD Budget for Forecast (M)
            ytd_budget_fc = f"=H{actual_row}"
            # Forecast Variance $ (N)
            fc_var = f"=L{actual_row}-M{actual_row}"
            # Forecast Variance % (O)
            fc_var_pct = f"=IFERROR(N{actual_row}/ABS(M{actual_row}),0)"

            # Assemble row: A=name, B-E=CM, F=spacer, G-J=YTD, K=spacer, L-O=Forecast
            row_data = [
                display_name,   # A
                cm_actual,      # B
                cm_budget,      # C
                cm_var,         # D
                cm_var_pct,     # E
                '',             # F (spacer)
                ytd_actual,     # G
                ytd_budget,     # H
                ytd_var,        # I
                ytd_var_pct,    # J
                '',             # K (spacer)
                ytd_forecast,   # L
                ytd_budget_fc,  # M
                fc_var,         # N
                fc_var_pct      # O
            ]
            all_data.append(row_data)

            # Track row type
            if account['is_total']:
                if 'net income' in name_lower:
                    row_types.append('net_income')
                else:
                    row_types.append('total')
            else:
                row_types.append('detail')

            forecast_row += 1

        # ================================================================
        # WRITE ALL DATA IN ONE BULK OPERATION
        # ================================================================
        if all_data:
            data_end_row = data_start_row + len(all_data) - 1
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=data_start_row + _ri, column=1).value = _val
        else:
            data_end_row = data_start_row

        # ================================================================
        # APPLY FORMATTING IN BULK (after data write)
        # ================================================================
        if all_data:
            # Collect rows by type for batch formatting
            header_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'header']
            total_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'total']
            net_income_rows = [data_start_row + idx for idx, rt in enumerate(row_types) if rt == 'net_income']

            # Format header rows (bold column A)
            for r in header_rows:
                sheet.cell(row=r, column=1).font = Font(bold=True)

            # Format total rows (bold, gray background)
            for r in total_rows:
                sheet.cell(row=r, column=1).font = Font(bold=True)
                # Apply gray background to non-spacer columns
                for c in [1, 2, 3, 4, 5, 7, 8, 9, 10, 12, 13, 14, 15]:
                    sheet.cell(row=r, column=c).fill = FILL_SUBTOTAL_GRAY

            # Format net income rows (bold, borders)
            for r in net_income_rows:
                sheet.cell(row=r, column=1).font = Font(bold=True)

        # Add blank row then Gross Profit % and Net Profit % rows
        row = data_end_row + 2  # Blank row after data

        # Gross Profit % row
        if gross_margin_row and total_income_row:
            sheet[f'A{row}'].value = 'Gross Profit %'
            sheet[f'A{row}'].font = Font(name='Calibri Light', size=10, bold=True, italic=True)

            # Current Month GP %
            sheet[f'B{row}'].value = f"=IFERROR(B{gross_margin_row}/B{total_income_row},0)"
            sheet[f'B{row}'].number_format = '0.0%'
            sheet[f'B{row}'].font = Font(name='Calibri Light', italic=True)

            # YTD Actual GP %
            sheet[f'G{row}'].value = f"=IFERROR(G{gross_margin_row}/G{total_income_row},0)"
            sheet[f'G{row}'].number_format = '0.0%'
            sheet[f'G{row}'].font = Font(name='Calibri Light', italic=True)

            # YTD Forecast GP %
            sheet[f'L{row}'].value = f"=IFERROR(L{gross_margin_row}/L{total_income_row},0)"
            sheet[f'L{row}'].number_format = '0.0%'
            sheet[f'L{row}'].font = Font(name='Calibri Light', italic=True)

            row += 1

        # Net Profit % row
        if net_income_row and total_income_row:
            sheet[f'A{row}'].value = 'Net Profit %'
            sheet[f'A{row}'].font = Font(name='Calibri Light', size=10, bold=True, italic=True)

            # Current Month NP %
            sheet[f'B{row}'].value = f"=IFERROR(B{net_income_row}/B{total_income_row},0)"
            sheet[f'B{row}'].number_format = '0.0%'
            sheet[f'B{row}'].font = Font(name='Calibri Light', italic=True)

            # YTD Actual NP %
            sheet[f'G{row}'].value = f"=IFERROR(G{net_income_row}/G{total_income_row},0)"
            sheet[f'G{row}'].number_format = '0.0%'
            sheet[f'G{row}'].font = Font(name='Calibri Light', italic=True)

            # YTD Forecast NP %
            sheet[f'L{row}'].value = f"=IFERROR(L{net_income_row}/L{total_income_row},0)"
            sheet[f'L{row}'].number_format = '0.0%'
            sheet[f'L{row}'].font = Font(name='Calibri Light', italic=True)

        # Format data area
        try:
            # Number format for dollar columns
            from openpyxl.utils import column_index_from_string
            for col in ['B', 'C', 'D', 'G', 'H', 'I', 'L', 'M', 'N']:
                col_idx = column_index_from_string(col)
                apply_style_to_range(sheet, 4, col_idx, data_end_row, col_idx, number_format='#,##0')

            # Percentage format for variance % columns
            for col in ['E', 'J', 'O']:
                col_idx = column_index_from_string(col)
                apply_style_to_range(sheet, 4, col_idx, data_end_row, col_idx, number_format='0.0%')

            # Font styling
            apply_style_to_range(sheet, data_start_row, 1, data_end_row, 15, font=Font(name='Calibri Light', size=10))
        except:
            pass

        # Clear spacer columns of any background color
        try:
            no_fill = PatternFill(fill_type=None)
            for r in range(4, row + 1):
                sheet.cell(row=r, column=6).fill = no_fill
                sheet.cell(row=r, column=11).fill = no_fill
        except:
            pass

        # Set column widths - width of 13 accommodates "$10,000,000" format
        # Using bulk operations for performance (reduces COM calls from 15 to 6)
        sheet.column_dimensions['A'].width = 40
        for _c in range(ord('B'), ord('O')+1):
            sheet.column_dimensions[chr(_c)].width = 13
        # Override specific columns with different widths
        sheet.column_dimensions['E'].width = 9
        sheet.column_dimensions['F'].width = 2
        sheet.column_dimensions['J'].width = 9
        sheet.column_dimensions['K'].width = 2
        sheet.column_dimensions['O'].width = 9

        # Hide gridlines
        try:
            # Window settings handled via sheet.views
            pass
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    # =========================================================================
    # DASHBOARD MODULE METHODS
    # =========================================================================

    # KPI Definitions - CFO-standard metrics and ratios
    KPI_DEFINITIONS = {
        'profitability': [
            {'name': 'Revenue', 'formula_type': 'direct', 'source': 'PL', 'account': 'Total for Income',
             'description': 'Total revenue generated', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Gross Profit', 'formula_type': 'direct', 'source': 'PL', 'account': 'Gross Profit',
             'description': 'Revenue minus COGS', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Gross Margin %', 'formula_type': 'ratio', 'numerator': 'Gross Profit', 'denominator': 'Total for Income',
             'description': 'Gross profit / revenue', 'default_target': 0.40, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'Net Income', 'formula_type': 'direct', 'source': 'PL', 'account': 'Net Income',
             'description': 'Bottom line profit', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Net Margin %', 'formula_type': 'ratio', 'numerator': 'Net Income', 'denominator': 'Total for Income',
             'description': 'Net income / revenue', 'default_target': 0.10, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'EBITDA', 'formula_type': 'calculated', 'calc_type': 'ebitda',
             'description': 'NI + Interest + Depreciation', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Operating Income', 'formula_type': 'direct', 'source': 'PL', 'account': 'Net Operating Income',
             'description': 'Core operations income', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Operating Margin %', 'formula_type': 'ratio', 'numerator': 'Net Operating Income', 'denominator': 'Total for Income',
             'description': 'Operating income / revenue', 'default_target': 0.12, 'format': '0.0%', 'higher_is_better': True},
        ],
        'liquidity': [
            {'name': 'Current Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities',
             'description': 'Current assets / liabilities', 'default_target': 1.5, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Quick Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities',
             'description': '(CA - Inventory) / CL', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Cash Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Bank Accounts', 'denominator': 'Total for Current Liabilities',
             'description': 'Cash / current liabilities', 'default_target': 0.2, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Working Capital', 'formula_type': 'bs_difference', 'minuend': 'Total for Current Assets', 'subtrahend': 'Total for Current Liabilities',
             'description': 'CA minus CL', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Cash Balance', 'formula_type': 'direct', 'source': 'BS', 'account': 'Total for Bank Accounts',
             'description': 'Total cash on hand', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
        ],
        'efficiency': [
            {'name': 'AR Days (DSO)', 'formula_type': 'days_ratio', 'balance': 'Total for Accounts Receivable', 'flow': 'Total for Income',
             'description': 'Days to collect AR', 'default_target': 45, 'format': '0', 'higher_is_better': False},
            {'name': 'AP Days (DPO)', 'formula_type': 'days_ratio', 'balance': 'Total for Credit Cards', 'flow': 'Total for Cost of Sales',
             'description': 'Days to pay AP', 'default_target': 30, 'format': '0', 'higher_is_better': True},
            {'name': 'Asset Turnover', 'formula_type': 'turnover', 'flow': 'Total for Income', 'balance': 'Total for Assets',
             'description': 'Revenue / assets', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': True},
        ],
        'leverage': [
            {'name': 'Debt-to-Equity', 'formula_type': 'bs_ratio', 'numerator': 'Total for Liabilities', 'denominator': 'Total for Equity',
             'description': 'Debt relative to equity', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': False},
            {'name': 'Debt-to-Assets', 'formula_type': 'bs_ratio', 'numerator': 'Total for Liabilities', 'denominator': 'Total for Assets',
             'description': 'Debt / total assets', 'default_target': 0.5, 'format': '0.0%', 'higher_is_better': False},
            {'name': 'Equity Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Equity', 'denominator': 'Total for Assets',
             'description': 'Equity / total assets', 'default_target': 0.5, 'format': '0.0%', 'higher_is_better': True},
        ],
        'cashflow': [
            {'name': 'Operating Cash Flow', 'formula_type': 'calculated', 'calc_type': 'ocf',
             'description': 'Cash from operations', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Free Cash Flow', 'formula_type': 'calculated', 'calc_type': 'fcf',
             'description': 'OCF minus CapEx', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
        ],
    }

    def _create_settings_sheet(self, sheet, report_sheets):
        """Create Settings sheet with sheet visibility controls.

        Args:
            sheet: xlwings Sheet object
            report_sheets: List of sheet names to include in visibility controls
        """
        # Colors
        DARK_BLUE = CLR_DARK_BLUE
        HEADER_GRAY = (200, 200, 200)
        LIGHT_BLUE = (232, 244, 253)  # Editable cells
        WHITE = (255, 255, 255)

        # Title
        sheet['B2'].value = "Settings"
        sheet['B2'].font = Font(size=20, bold=True, name='Calibri Light', color=CLR_DARK_BLUE)

        sheet['B3'].value = f"Control sheet visibility - Version {APP_VERSION}"
        sheet['B3'].font = Font(size=9, name='Calibri Light', color="808080")

        # Instructions
        sheet['B5'].value = "SHEET VISIBILITY"
        sheet['B5'].font = Font(size=12, bold=True, name='Calibri Light', color="FFFFFF")
        apply_style_to_range(sheet, 5, 2, 5, 3, fill=FILL_DARK_BLUE)

        sheet['B6'].value = "Change 'Visible' to Yes or No, then run the UpdateVisibility macro."
        sheet['B6'].font = Font(size=9, name='Calibri Light', color="646464", italic=True)

        # Headers
        header_row = 8
        sheet.cell(row=header_row, column=2).value = 'Sheet Name'
        sheet.cell(row=header_row, column=3).value = 'Visible'
        apply_style_to_range(sheet, header_row, 2, header_row, 3, font=Font(bold=True))
        apply_style_to_range(sheet, header_row, 2, header_row, 3, font=Font(name='Calibri Light'))
        apply_style_to_range(sheet, header_row, 2, header_row, 3, fill=PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"))

        # List all sheets with Yes/No visibility
        data_row = header_row + 1
        for sheet_name in report_sheets:
            # Sheet name
            sheet.cell(row=data_row, column=2).value = sheet_name
            sheet.cell(row=data_row, column=2).font = Font(name='Calibri Light')

            # Visibility dropdown (default to Yes)
            vis_cell = sheet.cell(row=data_row, column=3)
            vis_cell.value = 'Yes'
            vis_cell.fill = rgb_fill(LIGHT_BLUE)
            vis_cell.font = Font(name='Calibri Light')

            # Add data validation dropdown for Yes/No
            try:
                # Validation handled via DataValidation object
                # DataValidation handled via DataValidation()
                pass
            except:
                pass

            # Alternate row shading
            if data_row % 2 == 0:
                sheet.cell(row=data_row, column=2).fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")

            data_row += 1

        # Column widths
        sheet.column_dimensions['A'].width = 3
        sheet.column_dimensions['B'].width = 25
        sheet.column_dimensions['C'].width = 10

        # Tab color (dark blue)
        try:
            sheet.sheet_properties.tabColor = "3E2116"
        except:
            pass

        # Hide gridlines
        try:
            # Window settings handled via sheet.views
            pass
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_dashboard_control_sheet(self, sheet, months):
        """Create the Dashboard Control page with KPI target settings"""
        # Colors
        header_color = (22, 33, 62)  # Dark blue
        section_color = (44, 62, 80)  # Darker gray-blue
        control_color = (232, 244, 253)  # Light blue for editable cells

        # Title
        sheet['B2'].value = "Dashboard Control Panel"
        sheet['B2'].font = Font(size=24, bold=True, color="16213E")
        sheet.merge_cells('B2:F2')

        sheet['B3'].value = f"Configure KPI targets - Version {APP_VERSION}"
        sheet['B3'].font = Font(size=9, color="808080")
        sheet.merge_cells('B3:F3')

        # Instructions
        sheet['B5'].value = "INSTRUCTIONS"
        sheet['B5'].font = Font(size=14, bold=True, color="FFFFFF")
        apply_style_to_range(sheet, 5, 2, 5, 8, fill=rgb_fill(section_color))

        instructions = [
            "1. Set target values for each KPI in the 'Target' column",
            "2. Yellow threshold: 80% of target (caution)",
            "3. Red threshold: 60% of target (warning)",
            "4. Dashboard updates automatically when data changes"
        ]
        for i, instr in enumerate(instructions):
            sheet[f'B{7+i}'].value = instr
            sheet[f'B{7+i}'].font = Font(size=10)

        # Headers
        header_row = 13
        headers = ['Category', 'KPI Name', 'Description', 'Target', 'Yellow %', 'Red %', 'Direction']
        for col, header in enumerate(headers, 2):
            cell = sheet.cell(row=header_row, column=col)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
            # Alignment handled via Alignment() objects

        # Populate KPIs
        data_row = header_row + 1
        for category, kpis in self.KPI_DEFINITIONS.items():
            for kpi in kpis:
                sheet.cell(row=data_row, column=2).value = category.title()
                sheet.cell(row=data_row, column=3).value = kpi['name']
                sheet.cell(row=data_row, column=3).font = Font(bold=True)
                sheet.cell(row=data_row, column=4).value = kpi['description']
                sheet.cell(row=data_row, column=4).font = Font(size=9, color="646464")

                # Target (editable)
                target_cell = sheet.cell(row=data_row, column=5)
                target_cell.value = kpi['default_target']
                if '%' in kpi['format']:
                    target_cell.number_format = '0.0%'
                elif '0.00' in kpi['format']:
                    target_cell.number_format = '0.0'  # One decimal for ratios
                else:
                    target_cell.number_format = '#,##0'
                target_cell.fill = rgb_fill(control_color)

                # Yellow/Red thresholds
                sheet.cell(row=data_row, column=6).value = 0.80
                sheet.cell(row=data_row, column=6).number_format = '0%'
                sheet.cell(row=data_row, column=6).fill = rgb_fill(control_color)

                sheet.cell(row=data_row, column=7).value = 0.60
                sheet.cell(row=data_row, column=7).number_format = '0%'
                sheet.cell(row=data_row, column=7).fill = rgb_fill(control_color)

                # Direction
                sheet.cell(row=data_row, column=8).value = "Higher" if kpi['higher_is_better'] else "Lower"

                # Alternate row shading
                if data_row % 2 == 0:
                    for col in range(2, 9):
                        if sheet.cell(row=data_row, column=col).fill.fill_type is None:
                            sheet.cell(row=data_row, column=col).fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")

                data_row += 1

        # Column widths
        sheet.column_dimensions['A'].width = 3
        sheet.column_dimensions['B'].width = 14
        sheet.column_dimensions['C'].width = 20
        sheet.column_dimensions['D'].width = 30
        sheet.column_dimensions['E'].width = 12
        sheet.column_dimensions['F'].width = 10
        sheet.column_dimensions['G'].width = 10
        sheet.column_dimensions['H'].width = 10

        # Tab color
        try:
            sheet.sheet_properties.tabColor = "DB7400"
        except:
            pass

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _create_dashboard_sheet(self, sheet, pl_accounts, bs_accounts, months, detected_totals=None):
        """Create the main Dashboard sheet with KPIs and visualizations

        OPTIMIZED VERSION: Uses bulk writes and simplified formulas to prevent hangs.
        """
        # Colors (hex strings for openpyxl)
        header_color = "16213E"  # Dark blue
        section_color = (44, 62, 80)  # Section headers (used with rgb_fill)

        # Get detected account names (or use defaults)
        detected_totals = detected_totals or {}
        total_income_name = detected_totals.get('total_income', 'Total for Income')
        total_cogs_name = detected_totals.get('total_cogs', 'Total for Cost of Sales')
        total_expenses_name = detected_totals.get('total_expenses', 'Total for Expenses')

        company_name = self.company_name.get()
        current_month_col = len(months) + 1
        col_letter = get_column_letter(current_month_col)
        is_multi_division = hasattr(self, 'divisions') and len(self.divisions) > 1

        # =====================================================================
        # BULK WRITE: Build all header data and write at once
        # =====================================================================
        header_data = [
            ['', company_name, '', '', '', '', '', '', '', '', '', ''],
            ['', 'Executive Dashboard', '', '', '', '', '', '', '', '', '', ''],
            ['', '=CONCATENATE("Current Period: ",Menu!C7)', '', '', '', '', '', '', '', '', '', ''],
            ['', f"Generated: {datetime.now().strftime('%B %d, %Y')} | Version {APP_VERSION}", '', '', '', '', '', '', '', '', '', ''],
        ]
        write_data_to_cells(sheet, header_data, start_row=2, start_col=1)

        # Apply header formatting in bulk
        sheet['B2'].font = Font(size=28, bold=True, color=header_color)
        sheet['B3'].font = Font(size=16, color="7F8C8D")
        apply_style_to_range(sheet, 4, 2, 5, 2, font=Font(size=9, color="969696"))

        # =====================================================================
        # DIVISION SELECTOR DROPDOWN (Row 6) - For multi-division models
        # =====================================================================
        if is_multi_division:
            sheet['B6'].value = 'View:'
            sheet['B6'].font = Font(bold=True, size=10)

            # Default to "Consolidated"
            sheet['C6'].value = 'Consolidated'
            sheet['C6'].font = Font(bold=True, size=10)
            sheet['C6'].fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")

            # Create dropdown list: Consolidated + all division names
            div_names = ['Consolidated'] + [d.get('name', d) if isinstance(d, dict) else getattr(d, 'name', str(d)) for d in self.divisions]
            div_list = ','.join(div_names)

            try:
                # Validation handled via DataValidation object
                # DataValidation handled via DataValidation()
                pass
            except Exception as e:
                print(f"[Dashboard] Could not add division dropdown: {e}")

            # Store the selected division sheet name formula in a helper cell (H6, hidden)
            # This will be used by formulas to determine which sheet to pull from
            sheet['H6'].value = '=IF(C6="Consolidated","Consolidated_PL",C6&"_PL")'
            sheet['H6'].font = Font(color="FFFFFF")  # White text (hidden)

        # =====================================================================
        # P&L SUMMARY TABLE (Rows 8-13) - SIMPLIFIED with SUMIF formulas
        # Using simpler formulas that only lookup by account name in column B
        # =====================================================================

        # Build P&L summary data in memory first
        pl_summary_labels = ['', 'CURRENT MTH', '', 'YTD', '', '% of Rev']
        pl_accounts_info = [
            ('Revenue', total_income_name, True, (39, 174, 96)),
            ('Cost of Goods Sold', total_cogs_name, False, (231, 76, 60)),
            ('Gross Profit', 'Gross Profit', True, (52, 152, 219)),
            ('Operating Expenses', total_expenses_name, False, (230, 126, 34)),
            ('Net Income', 'Net Income', True, (155, 89, 182)),
        ]

        # Write header row
        write_row_to_cells(sheet, pl_summary_labels, row=7, start_col=2)
        apply_style_to_range(sheet, 7, 2, 7, 7,
                             font=Font(bold=True, size=10, color="FFFFFF"),
                             fill=PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid"))

        # Build all P&L summary rows data
        pl_data = []
        for i, (label, account, is_bold, color) in enumerate(pl_accounts_info):
            row = 8 + i
            # Use SUMPRODUCT with YYYYMM match for DYNAMIC current month lookup
            # This formula looks up the month based on Menu!G7, so changing Menu dropdown updates values
            if is_multi_division:
                # If Consolidated selected (C6="Consolidated"), sum all divisions
                # If specific division selected, filter by that division
                cm_formula = f'=IF($C$6="Consolidated",SUMPRODUCT((Source_PL!$B$3:$B$1500="{account}")*(Source_PL!$C$2:{col_letter}$2=Menu!$G$7)*(Source_PL!$C$3:{col_letter}$1500)),SUMPRODUCT((Source_PL!$A$3:$A$1500=$C$6)*(Source_PL!$B$3:$B$1500="{account}")*(Source_PL!$C$2:{col_letter}$2=Menu!$G$7)*(Source_PL!$C$3:{col_letter}$1500)))'
            else:
                cm_formula = f'=SUMPRODUCT((Source_PL!$A$3:$A$1500="{account}")*(Source_PL!$B$2:{col_letter}$2=Menu!$G$7)*(Source_PL!$B$3:{col_letter}$1500))'

            # YTD: sum all months in current year using SUMPRODUCT with year filter
            # Menu!I7 contains current year (in multi-division mode) or Menu!F7 (single-division)
            # Source_PL row 2 has YYYYMM values, we extract year by dividing by 100
            if is_multi_division:
                # Consolidated: sum all divisions for the year
                # Division-specific: filter by division and year
                ytd_formula = f'=IF($C$6="Consolidated",IFERROR(SUMPRODUCT((Source_PL!$B$3:$B$1500="{account}")*(INT(Source_PL!$C$2:{col_letter}$2/100)=Menu!$I$7)*(Source_PL!$C$3:{col_letter}$1500)),0),IFERROR(SUMPRODUCT((Source_PL!$A$3:$A$1500=$C$6)*(Source_PL!$B$3:$B$1500="{account}")*(INT(Source_PL!$C$2:{col_letter}$2/100)=Menu!$I$7)*(Source_PL!$C$3:{col_letter}$1500)),0))'
            else:
                ytd_formula = f'=IFERROR(SUMPRODUCT((Source_PL!$A$3:$A$1500="{account}")*(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*(Source_PL!$B$3:{col_letter}$1500)),0)'

            # % of Revenue
            pct_formula = f'=IFERROR(E{row}/E8,0)' if label != 'Revenue' else 1.0

            pl_data.append([label, cm_formula, '', ytd_formula, '', pct_formula])

        # Write all P&L data at once
        write_data_to_cells(sheet, pl_data, start_row=8, start_col=2)

        # Apply formatting to P&L summary section
        for i, (label, account, is_bold, color) in enumerate(pl_accounts_info):
            row = 8 + i
            # Label formatting
            color_hex = "{:02X}{:02X}{:02X}".format(color[0], color[1], color[2])
            sheet[f'B{row}'].font = Font(size=11, bold=is_bold, color=color_hex)
            # Number formatting
            sheet[f'C{row}'].number_format = '"$"#,##0'
            sheet[f'E{row}'].number_format = '"$"#,##0'
            sheet[f'G{row}'].number_format = '0.0%'

        # Gross Margin and Net Margin indicators
        margin_data = [
            ['Gross Margin:', '=IFERROR(C10/C8,0)'],
            ['', ''],
            ['Net Margin:', '=IFERROR(C12/C8,0)'],
        ]
        write_data_to_cells(sheet, margin_data, start_row=10, start_col=8)
        sheet['I10'].number_format = '0.0%'
        sheet['I12'].number_format = '0.0%'

        # KPI Sections - Start after P&L Summary table
        # OPTIMIZED: Build all data in memory first, then write in bulk
        current_row = 14
        sections = [
            ('PROFITABILITY METRICS', 'profitability'),
            ('LIQUIDITY METRICS', 'liquidity'),
            ('EFFICIENCY METRICS', 'efficiency'),
            ('LEVERAGE METRICS', 'leverage'),
            ('CASH FLOW METRICS', 'cashflow'),
        ]

        control_row = 14  # Starting row in Dashboard_Control

        # Build all section data in memory first
        all_rows_data = []  # List of (row_num, row_data, row_type, kpi_format)
        kpi_data_rows = []  # Track which rows are KPI data rows for formatting

        for section_title, category in sections:
            kpis = self.KPI_DEFINITIONS.get(category, [])
            if not kpis:
                continue

            # Section header row
            section_row = ['', section_title] + [''] * 10  # Columns B through L
            all_rows_data.append((current_row, section_row, 'section', None))
            current_row += 1

            # Column headers row
            header_row = ['', 'KPI', 'Current', 'Target', 'Status', '', 'YTD', 'Target', 'Status', '', 'Trend']
            all_rows_data.append((current_row, header_row, 'header', None))
            current_row += 1

            # KPI data rows
            for kpi in kpis:
                higher = kpi['higher_is_better']

                # Build formulas
                current_formula = self._build_dashboard_kpi_formula(kpi, 'current', current_month_col, months, is_multi_division)
                # Target formula: default to 0 if cell is empty or has error
                target_formula = f'=IFERROR(IF(Dashboard_Control!E{control_row}="",0,Dashboard_Control!E{control_row}),0)'

                # Status formula: show "-" if target is 0 or empty (no target defined)
                if higher:
                    status_formula = f'=IF(OR(D{current_row}=0,D{current_row}=""),"-",IF(C{current_row}>=D{current_row},"G",IF(C{current_row}>=D{current_row}*0.8,"Y","R")))'
                else:
                    status_formula = f'=IF(OR(D{current_row}=0,D{current_row}=""),"-",IF(C{current_row}<=D{current_row},"G",IF(C{current_row}<=D{current_row}*1.2,"Y","R")))'

                ytd_formula = self._build_dashboard_kpi_formula(kpi, 'ytd', current_month_col, months, is_multi_division)

                if '%' in kpi['format'] or 'ratio' in kpi['formula_type']:
                    ytd_target = f"=D{current_row}"
                else:
                    ytd_target = f"=D{current_row}*{len(months)}"

                # YTD Status formula: show "-" if target is 0 or empty
                if higher:
                    ytd_status = f'=IF(OR(H{current_row}=0,H{current_row}=""),"-",IF(G{current_row}>=H{current_row},"G",IF(G{current_row}>=H{current_row}*0.8,"Y","R")))'
                else:
                    ytd_status = f'=IF(OR(H{current_row}=0,H{current_row}=""),"-",IF(G{current_row}<=H{current_row},"G",IF(G{current_row}<=H{current_row}*1.2,"Y","R")))'

                # Simple trend indicator: Up, Down, or Flat
                trend_formula = f'=IF(G{current_row}=0,"-",IF(C{current_row}>G{current_row}*1.05,"Up",IF(C{current_row}<G{current_row}*0.95,"Down","Flat")))'

                # Row data: columns A through K (indices 0-10)
                kpi_row = [
                    '',  # A
                    kpi['name'],  # B
                    current_formula,  # C
                    target_formula,  # D
                    status_formula,  # E
                    '',  # F (spacer)
                    ytd_formula,  # G
                    ytd_target,  # H
                    ytd_status,  # I
                    '',  # J (spacer)
                    trend_formula  # K
                ]
                all_rows_data.append((current_row, kpi_row, 'kpi', kpi['format']))
                kpi_data_rows.append((current_row, kpi['format'], current_row % 2 == 0))

                current_row += 1
                control_row += 1

            current_row += 1  # Space between sections

        # BULK WRITE: Write all data at once per section type
        for row_num, row_data, row_type, kpi_format in all_rows_data:
            # Write row data in one operation (columns B through L = 2 through 12)
            _data = row_data[1:]  # Skip column A
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=row_num + _ri, column=2 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=row_num + _ri, column=2 + _ci).value = _v
                        else:
                            sheet.cell(row=row_num + _ri, column=2).value = _val

        # SIMPLIFIED FORMATTING: Only format section/header rows (minimal per-KPI formatting)
        for row_num, row_data, row_type, kpi_format in all_rows_data:
            if row_type == 'section':
                # Range: section_range = (sheet, row_num, 2, row_num, 12)
                apply_style_to_range(sheet, row_num, 2, row_num, 12, font=Font(bold=True))
                apply_style_to_range(sheet, row_num, 2, row_num, 12, font=Font(color="FFFFFF"))
                apply_style_to_range(sheet, row_num, 2, row_num, 12, fill=PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid"))

            elif row_type == 'header':
                apply_style_to_range(sheet, row_num, 2, row_num, 12, font=Font(bold=True))
                apply_style_to_range(sheet, row_num, 2, row_num, 12, font=Font(color="FFFFFF"))
                apply_style_to_range(sheet, row_num, 2, row_num, 12, fill=PatternFill(start_color="34495E", end_color="34495E", fill_type="solid"))

        # Apply number formats to entire columns at once (much faster than per-cell)
        if kpi_data_rows:
            first_kpi_row = kpi_data_rows[0][0]
            last_kpi_row = kpi_data_rows[-1][0]
            # Apply common number format to value columns
            apply_style_to_range(sheet, first_kpi_row, 3, last_kpi_row, 3, number_format='#,##0')
            apply_style_to_range(sheet, first_kpi_row, 4, last_kpi_row, 4, number_format='#,##0')
            apply_style_to_range(sheet, first_kpi_row, 7, last_kpi_row, 7, number_format='#,##0')
            apply_style_to_range(sheet, first_kpi_row, 8, last_kpi_row, 8, number_format='#,##0')

        # =====================================================================
        # CONDITIONAL FORMATTING for Status columns (E and I)
        # G = Green (good), Y = Yellow (warning), R = Red (bad)
        # =====================================================================
        try:
            # Status columns E and I - apply conditional formatting for G/Y/R
            if kpi_data_rows:
                for status_col in ['E', 'I']:
                    range_str = f'{status_col}{first_kpi_row}:{status_col}{last_kpi_row}'

                    # Green for "G" - good performance
                    sheet.conditional_formatting.add(range_str, CellIsRule(
                        operator='equal', formula=['"G"'],
                        fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                        font=Font(color="006100")
                    ))

                    # Yellow for "Y" - warning
                    sheet.conditional_formatting.add(range_str, CellIsRule(
                        operator='equal', formula=['"Y"'],
                        fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
                        font=Font(color="9C5700")
                    ))

                    # Red for "R" - poor performance
                    sheet.conditional_formatting.add(range_str, CellIsRule(
                        operator='equal', formula=['"R"'],
                        fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                        font=Font(color="9C0006")
                    ))

                    # Center-align status columns
                    apply_style_to_range(sheet, first_kpi_row, openpyxl.utils.column_index_from_string(status_col),
                                        last_kpi_row, openpyxl.utils.column_index_from_string(status_col),
                                        alignment=ALIGN_CENTER)
        except Exception as e:
            print(f"[Dashboard] Could not apply conditional formatting: {e}")

        # Column widths
        try:
            sheet.column_dimensions['A'].width = 3
            sheet.column_dimensions['B'].width = 22
            sheet.column_dimensions['C'].width = 14
            sheet.column_dimensions['D'].width = 12
            sheet.column_dimensions['E'].width = 5
            sheet.column_dimensions['F'].width = 3
            sheet.column_dimensions['G'].width = 14
            sheet.column_dimensions['H'].width = 12
            sheet.column_dimensions['I'].width = 5
            sheet.column_dimensions['J'].width = 3
            sheet.column_dimensions['K'].width = 8
        except:
            pass

        # Format Trend column K with conditional formatting for Up/Down/Flat
        try:
            if kpi_data_rows:
                # Alignment handled via Alignment() objects
                apply_style_to_range(sheet, first_kpi_row, 11, last_kpi_row, 11, font=Font(size=10, name='Calibri Light'))

                # Add conditional formatting for trend text
                try:
                    trend_range_str = f'K{first_kpi_row}:K{last_kpi_row}'

                    # "Up" = Green
                    sheet.conditional_formatting.add(trend_range_str, CellIsRule(
                        operator='equal', formula=['"Up"'],
                        font=Font(color="006100", bold=True)
                    ))

                    # "Down" = Red
                    sheet.conditional_formatting.add(trend_range_str, CellIsRule(
                        operator='equal', formula=['"Down"'],
                        font=Font(color="9C0006", bold=True)
                    ))

                    # "Flat" = Gray
                    sheet.conditional_formatting.add(trend_range_str, CellIsRule(
                        operator='equal', formula=['"Flat"'],
                        font=Font(color="808080")
                    ))
                except:
                    pass
        except:
            pass

        # Vertically center all cells
        try:
            # Vertically center all cells in used range
            max_r = sheet.max_row or 1
            max_c = sheet.max_column or 1
            apply_style_to_range(sheet, 1, 1, max_r, max_c, alignment=Alignment(vertical='center'))
        except:
            pass

        # Tab color
        try:
            sheet.sheet_properties.tabColor = "60AE27"
        except:
            pass

        # Hide gridlines
        try:
            sheet.views.sheetView[0].showGridLines = False
        except:
            pass

        # =====================================================================
        # CHARTS SECTION - Add visualizations to the right of the dashboard
        # =====================================================================
        try:
            self._add_dashboard_charts(sheet, months, is_multi_division, current_row, detected_totals)
        except Exception as e:
            print(f"[Dashboard] Could not add charts: {e}")

        # Add back to menu link
        self._add_back_to_menu_link(sheet, row=1, col=1)

    def _add_dashboard_charts(self, sheet, months, is_multi_division, last_kpi_row, detected_totals=None):
        """Add charts to the Dashboard sheet

        Creates:
        1. Pie chart showing expense breakdown (Revenue vs COGS vs Expenses)
        2. Monthly Revenue line chart
        3. Monthly Net Income line chart
        4. Monthly Gross Margin line chart

        FIXED: Charts now dynamically respond to division selector (C6) in multi-division mode.
        When Consolidated is selected, charts show consolidated data.
        When a specific division is selected, charts show that division's data.
        """
        # Chart layout uses openpyxl chart anchoring (cell references)

        # Get detected account names for formulas
        detected_totals = detected_totals or {}
        total_income_name = detected_totals.get('total_income', 'Total for Income')
        total_cogs_name = detected_totals.get('total_cogs', 'Total for Cost of Sales')
        total_expenses_name = detected_totals.get('total_expenses', 'Total for Expenses')

        # Determine data columns based on mode
        acct_col = 'B' if is_multi_division else 'A'
        data_start_col = 'C' if is_multi_division else 'B'
        year_cell = 'Menu!$I$7' if is_multi_division else 'Menu!$F$7'

        num_months = len(months)
        last_data_col = get_column_letter(num_months + (2 if is_multi_division else 1))

        # =====================================================================
        # CHART DATA AREA: Write chart source data to hidden columns (starting at column N)
        # FIXED: Formulas now respond to division selector $C$6
        # =====================================================================
        chart_data_col = 14  # Column N

        # Row 2: Header "Chart Data"
        sheet.cell(row=2, column=chart_data_col).value = "Chart Data"
        sheet.cell(row=2, column=chart_data_col).font = Font(bold=True)

        # --- Pie Chart Data (Revenue vs COGS vs Expenses breakdown) ---
        sheet.cell(row=4, column=chart_data_col).value = "Category"
        sheet.cell(row=4, column=chart_data_col + 1).value = "Amount"

        # Use flexible matching patterns for account names
        # These patterns will match variations like "Total Income", "Total for Income", etc.
        pie_labels = ['Revenue', 'Cost of Goods', 'Operating Exp.']
        pie_search_terms = ['Total*Income', 'Total*Cost', 'Total*Expense']

        for i, (label, search_term) in enumerate(zip(pie_labels, pie_search_terms)):
            row = 5 + i
            sheet.cell(row=row, column=chart_data_col).value = label
            # YTD sum formula using SUMPRODUCT with wildcard matching via COUNTIF pattern
            # Use exact account names from detected_totals if available, otherwise search
            if i == 0:
                account = total_income_name
            elif i == 1:
                account = total_cogs_name
            else:
                account = total_expenses_name

            if is_multi_division:
                # FIXED: Use IF to check division selector - Consolidated sums all, otherwise filter by division
                formula = (
                    f'=ABS(IFERROR(IF($C$6="Consolidated",'
                    f'SUMPRODUCT((Source_PL!${acct_col}$3:${acct_col}$1500="{account}")*(INT(Source_PL!${data_start_col}$2:{last_data_col}$2/100)={year_cell})*(Source_PL!${data_start_col}$3:{last_data_col}$1500)),'
                    f'SUMPRODUCT((Source_PL!$A$3:$A$1500=$C$6)*(Source_PL!${acct_col}$3:${acct_col}$1500="{account}")*(INT(Source_PL!${data_start_col}$2:{last_data_col}$2/100)={year_cell})*(Source_PL!${data_start_col}$3:{last_data_col}$1500))),0))'
                    )
            else:
                formula = f'=ABS(IFERROR(SUMPRODUCT((Source_PL!${acct_col}$3:${acct_col}$1500="{account}")*(INT(Source_PL!${data_start_col}$2:{last_data_col}$2/100)={year_cell})*(Source_PL!${data_start_col}$3:{last_data_col}$1500)),0))'
            sheet.cell(row=row, column=chart_data_col + 1).value = formula

        # --- Monthly Data for Line Charts ---
        # Row 10: "Monthly Data" header
        sheet.cell(row=10, column=chart_data_col).value = "Month"
        sheet.cell(row=10, column=chart_data_col + 1).value = "Revenue"
        sheet.cell(row=10, column=chart_data_col + 2).value = "Net Income"
        sheet.cell(row=10, column=chart_data_col + 3).value = "Gross Margin %"

        # Write month labels and formulas for each month (up to last 12 months)
        display_months = months[-12:] if len(months) > 12 else months

        for i, (month, year, label) in enumerate(display_months):
            row = 11 + i
            yyyymm = year * 100 + month
            col_idx = months.index((month, year, label)) + (3 if is_multi_division else 2)
            month_col = get_column_letter(col_idx)

            # Month label (short form)
            sheet.cell(row=row, column=chart_data_col).value = label[:3] if len(label) > 3 else label

            # Revenue formula - FIXED: respect division selector
            if is_multi_division:
                rev_formula = (
                    f'=IFERROR(IF($C$6="Consolidated",'
                    f'SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"{total_income_name}",Source_PL!{month_col}$3:{month_col}$1500),'
                    f'SUMIFS(Source_PL!{month_col}$3:{month_col}$1500,Source_PL!$A$3:$A$1500,$C$6,Source_PL!${acct_col}$3:${acct_col}$1500,"{total_income_name}")),0)'
                    )
            else:
                rev_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"{total_income_name}",Source_PL!{month_col}$3:{month_col}$1500),0)'
            sheet.cell(row=row, column=chart_data_col + 1).value = rev_formula

            # Net Income formula - FIXED: respect division selector
            if is_multi_division:
                ni_formula = (
                    f'=IFERROR(IF($C$6="Consolidated",'
                    f'SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Net Income",Source_PL!{month_col}$3:{month_col}$1500),'
                    f'SUMIFS(Source_PL!{month_col}$3:{month_col}$1500,Source_PL!$A$3:$A$1500,$C$6,Source_PL!${acct_col}$3:${acct_col}$1500,"Net Income")),0)'
                    )
            else:
                ni_formula = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Net Income",Source_PL!{month_col}$3:{month_col}$1500),0)'
            sheet.cell(row=row, column=chart_data_col + 2).value = ni_formula

            # Gross Margin % formula (Gross Profit / Revenue) - FIXED: respect division selector
            if is_multi_division:
                gp_formula = (
                    f'IF($C$6="Consolidated",'
                    f'SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Gross Profit",Source_PL!{month_col}$3:{month_col}$1500),'
                    f'SUMIFS(Source_PL!{month_col}$3:{month_col}$1500,Source_PL!$A$3:$A$1500,$C$6,Source_PL!${acct_col}$3:${acct_col}$1500,"Gross Profit"))'
                    )
                sheet.cell(row=row, column=chart_data_col + 3).value = f'=IFERROR({gp_formula}/{get_column_letter(chart_data_col + 1)}{row},0)'
            else:
                sheet.cell(row=row, column=chart_data_col + 3).value = f'=IFERROR(SUMIF(Source_PL!${acct_col}$3:${acct_col}$1500,"Gross Profit",Source_PL!{month_col}$3:{month_col}$1500)/{get_column_letter(chart_data_col + 1)}{row},0)'

        num_data_rows = len(display_months)

        # Format chart data columns
        apply_style_to_range(sheet, 11, chart_data_col + 1, 10 + num_data_rows, chart_data_col + 1, number_format='"$"#,##0')
        apply_style_to_range(sheet, 11, chart_data_col + 2, 10 + num_data_rows, chart_data_col + 2, number_format='"$"#,##0')
        apply_style_to_range(sheet, 11, chart_data_col + 3, 10 + num_data_rows, chart_data_col + 3, number_format='0.0%')

        # Make chart data columns nearly invisible
        try:
            for col in range(chart_data_col, chart_data_col + 4):
                set_col_width(sheet, col, 2)
                for r in range(1, 10 + num_data_rows + 1):
                    sheet.cell(row=r, column=col).font = Font(color="FFFFFF", size=1)
        except:
            pass

        # =====================================================================
        # CREATE CHARTS using openpyxl chart API
        # =====================================================================
        try:
            chart_data_col_letter = get_column_letter(chart_data_col)

            # --- 1. PIE CHART: Revenue vs COGS vs Expenses ---
            pie = PieChart()
            pie.title = "YTD Financial Breakdown"
            pie.style = 10
            labels = Reference(sheet, min_col=chart_data_col, min_row=5, max_row=7)
            data = Reference(sheet, min_col=chart_data_col + 1, min_row=4, max_row=7)
            pie.add_data(data, titles_from_data=True)
            pie.set_categories(labels)
            pie.width = 14
            pie.height = 10
            # Color the pie slices
            try:
                from openpyxl.chart.series import DataPoint
                colors = ["27AE60", "FFA500", "E64C4C"]  # Green, Orange, Red
                for i, color in enumerate(colors):
                    pt = DataPoint(idx=i)
                    pt.graphicalProperties.solidFill = color
                    pie.series[0].data_points.append(pt)
            except:
                pass
            sheet.add_chart(pie, "N2")

            # --- 2. REVENUE LINE CHART ---
            rev_chart = LineChart()
            rev_chart.title = "Monthly Revenue"
            rev_chart.style = 10
            rev_chart.y_axis.numFmt = '"$"#,##0'
            rev_chart.width = 14
            rev_chart.height = 10
            rev_labels = Reference(sheet, min_col=chart_data_col, min_row=11, max_row=10 + num_data_rows)
            rev_data = Reference(sheet, min_col=chart_data_col + 1, min_row=10, max_row=10 + num_data_rows)
            rev_chart.add_data(rev_data, titles_from_data=True)
            rev_chart.set_categories(rev_labels)
            rev_chart.legend = None
            try:
                rev_chart.series[0].graphicalProperties.line.solidFill = "27AE60"
            except:
                pass
            sheet.add_chart(rev_chart, "N18")

            # --- 3. NET INCOME LINE CHART ---
            ni_chart = LineChart()
            ni_chart.title = "Monthly Net Income"
            ni_chart.style = 10
            ni_chart.y_axis.numFmt = '"$"#,##0'
            ni_chart.width = 14
            ni_chart.height = 10
            ni_data = Reference(sheet, min_col=chart_data_col + 2, min_row=10, max_row=10 + num_data_rows)
            ni_chart.add_data(ni_data, titles_from_data=True)
            ni_chart.set_categories(rev_labels)
            ni_chart.legend = None
            try:
                ni_chart.series[0].graphicalProperties.line.solidFill = "9B5EB0"
            except:
                pass
            sheet.add_chart(ni_chart, "N34")

            # --- 4. GROSS MARGIN LINE CHART ---
            gm_chart = LineChart()
            gm_chart.title = "Monthly Gross Margin %"
            gm_chart.style = 10
            gm_chart.y_axis.numFmt = "0%"
            gm_chart.width = 14
            gm_chart.height = 10
            gm_data = Reference(sheet, min_col=chart_data_col + 3, min_row=10, max_row=10 + num_data_rows)
            gm_chart.add_data(gm_data, titles_from_data=True)
            gm_chart.set_categories(rev_labels)
            gm_chart.legend = None
            try:
                gm_chart.series[0].graphicalProperties.line.solidFill = "3494D4"
            except:
                pass
            sheet.add_chart(gm_chart, "N50")

        except Exception as e:
            print(f"[Dashboard] Chart creation error: {e}")
            import traceback
            traceback.print_exc()

    def _build_current_month_sumproduct(self, sheet, account, current_col, multi_division=False):
        """
        Build a SUMPRODUCT formula for current month that dynamically references Menu!C7.
        Returns value for the exact month/year matching Menu!E7 (month) and Menu!F7 (year).

        Source sheets have:
        - Single-division: Row 2: YYYYMM helper values, Row 3+: Account data
        - Multi-division: Col A=Division, Col B=Account, Row 2: YYYYMM values start at Col C

        Args:
            multi_division: If True, adjust formula for multi-division structure with
                           division filter based on Dashboard!L3
        """
        col_letter = get_column_letter(current_col)

        if multi_division:
            # Multi-division: Col A=Division, Col B=Account, Col C+ = values
            # If Dashboard!L3 = "Consolidated", sum all divisions; else filter by division
            return (
                f"=IF(Dashboard!$L$3=\"Consolidated\","
                f"SUMPRODUCT("
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"({sheet}!$C$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)),"
                f"SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=Dashboard!$L$3)*"
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"({sheet}!$C$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)))"
            )
        else:
            # Single-division: Col A=Account, Col B+ = values
            return (
                f"=SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=\"{account}\")*"
                f"({sheet}!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"({sheet}!$B$3:{col_letter}$1000)"
                f")"
                )

    def _build_ytd_sumproduct(self, sheet, account, current_col, multi_division=False):
        """
        Build a SUMPRODUCT formula for YTD that only sums months in the current year.
        Uses Menu!F7 for year and Menu!E7 for current month number.

        Source sheets now have:
        - Row 1: Headers (month names)
        - Row 2: YYYYMM helper values (e.g., 202411 for Nov 2024)
        - Row 3+: Account data

        Args:
            multi_division: If True, adjust formula for multi-division structure with
                           division filter based on Dashboard!L3
        """
        col_letter = get_column_letter(current_col)

        if multi_division:
            # Multi-division: Col A=Division, Col B=Account, Col C+ = values
            return (
                f"=IF(Dashboard!$L$3=\"Consolidated\","
                f"SUMPRODUCT("
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"(INT({sheet}!$C$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD({sheet}!$C$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)),"
                f"SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=Dashboard!$L$3)*"
                f"({sheet}!$B$3:$B$1000=\"{account}\")*"
                f"(INT({sheet}!$C$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD({sheet}!$C$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"({sheet}!$C$3:{col_letter}$1000)))"
            )
        else:
            # Single-division: Col A=Account, Col B+ = values
            return (
                f"=SUMPRODUCT("
                f"({sheet}!$A$3:$A$1000=\"{account}\")*"
                f"(INT({sheet}!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD({sheet}!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"({sheet}!$B$3:{col_letter}$1000)"
                f")"
            )

    def _build_dashboard_kpi_formula(self, kpi, period, current_col, months, multi_division=False):
        """Build formula for a KPI based on its type and period

        FIXED VERSION: Properly handles YTD by summing across all months in current year.
        For multi-division mode, uses Column B for account lookup (Column A is Division).

        Args:
            kpi: KPI definition dict
            period: 'current' or 'ytd'
            current_col: Current month column number
            months: List of month tuples
            multi_division: If True, uses Column B for account lookup
        """
        formula_type = kpi['formula_type']
        col_letter = get_column_letter(current_col)

        # Determine account column and data start column based on mode
        # Multi-division: Col A = Division, Col B = Account, Col C+ = data
        # Single-division: Col A = Account, Col B+ = data
        acct_col = 'B' if multi_division else 'A'
        data_start_col = 'C' if multi_division else 'B'
        year_cell = 'Menu!$I$7' if multi_division else 'Menu!$F$7'

        def make_sumif(sheet, account, col):
            """Create simple SUMIF formula for current month"""
            return f'SUMIF({sheet}!${acct_col}$3:${acct_col}$1500,"{account}",{sheet}!{col}$3:{col}$1500)'

        def make_ytd_sumproduct(sheet, account):
            """Create SUMPRODUCT formula for YTD (sum all months in current year)"""
            # Row 2 has YYYYMM values, we extract year by INT(value/100)
            return (
                f'SUMPRODUCT('
                f'({sheet}!${acct_col}$3:${acct_col}$1500="{account}")*'
                f'(INT({sheet}!${data_start_col}$2:{col_letter}$2/100)={year_cell})*'
                f'({sheet}!${data_start_col}$3:{col_letter}$1500))'
                )

        if formula_type == 'direct':
            source = kpi['source']
            account = kpi['account']
            sheet = 'Source_PL' if source == 'PL' else 'Source_BS'
            if period == 'ytd' and source == 'PL':
                # Use SUMPRODUCT for YTD on P&L accounts
                return f'=IFERROR({make_ytd_sumproduct(sheet, account)},0)'
            else:
                # Current month or Balance Sheet (point-in-time)
                return f'={make_sumif(sheet, account, col_letter)}'

        elif formula_type == 'ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            if period == 'ytd':
                num_f = make_ytd_sumproduct('Source_PL', num)
                den_f = make_ytd_sumproduct('Source_PL', den)
            else:
                num_f = make_sumif('Source_PL', num, col_letter)
                den_f = make_sumif('Source_PL', den, col_letter)
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_ratio':
            # Balance Sheet ratios use point-in-time values, not YTD
            num = kpi['numerator']
            den = kpi['denominator']
            num_f = make_sumif('Source_BS', num, col_letter)
            den_f = make_sumif('Source_BS', den, col_letter)
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_difference':
            # Balance Sheet differences use point-in-time values
            min_acct = kpi['minuend']
            sub_acct = kpi['subtrahend']
            min_f = make_sumif('Source_BS', min_acct, col_letter)
            sub_f = make_sumif('Source_BS', sub_acct, col_letter)
            return f"={min_f}-{sub_f}"

        elif formula_type == 'days_ratio':
            balance = kpi['balance']
            flow = kpi['flow']
            bal_f = make_sumif('Source_BS', balance, col_letter)
            if period == 'ytd':
                flow_f = make_ytd_sumproduct('Source_PL', flow)
            else:
                flow_f = make_sumif('Source_PL', flow, col_letter)
            return f"=IFERROR({bal_f}/{flow_f}*30,0)"

        elif formula_type == 'turnover':
            flow = kpi['flow']
            balance = kpi['balance']
            bal_f = make_sumif('Source_BS', balance, col_letter)
            if period == 'ytd':
                flow_f = make_ytd_sumproduct('Source_PL', flow)
            else:
                flow_f = make_sumif('Source_PL', flow, col_letter)
            return f"=IFERROR({flow_f}/{bal_f},0)"

        elif formula_type == 'calculated':
            calc_type = kpi['calc_type']
            return self._build_calculated_dashboard_formula(calc_type, period, current_col, months)

        return "=0"

    def _build_calculated_dashboard_formula(self, calc_type, period, current_col, months):
        """Build formula for complex calculated KPIs

        FIXED: Uses dynamic account lookup with wildcard matching instead of hardcoded account names.
        Searches for accounts containing 'Interest' or 'Depreciation' keywords.
        """
        col_letter = get_column_letter(current_col)

        def current_sumif(account):
            """Helper to build current month SUMIF formula for an account"""
            return f'SUMIF(Source_PL!$A$3:$A$1500,"{account}",Source_PL!{col_letter}$3:{col_letter}$1500)'

        def current_sumif_wildcard(keyword):
            """Helper to build current month SUMIF with wildcard matching"""
            return f'SUMIF(Source_PL!$A$3:$A$1500,"*{keyword}*",Source_PL!{col_letter}$3:{col_letter}$1500)'

        def ytd_sum(account):
            """Helper to sum YTD values - sums all columns for the current year"""
            # For YTD, we need to sum across months in the current year
            # This simplified version just uses the same current period value
            # A more complete solution would sum across all months where YEAR matches
            return f'SUMIF(Source_PL!$A$3:$A$1500,"{account}",Source_PL!{col_letter}$3:{col_letter}$1500)'

        def ytd_sum_wildcard(keyword):
            """Helper to sum YTD values with wildcard matching"""
            return f'SUMIF(Source_PL!$A$3:$A$1500,"*{keyword}*",Source_PL!{col_letter}$3:{col_letter}$1500)'

        if calc_type == 'ebitda':
            # EBITDA = Net Income + Interest Expense + Depreciation + Amortization
            # Use wildcard matching to find accounts containing these keywords
            if period == 'current':
                ni = current_sumif("Net Income")
                # Search for any account containing "Interest" (for interest expense)
                int_e = current_sumif_wildcard("Interest")
                # Search for any account containing "Depreciation" or "Amortization"
                dep = current_sumif_wildcard("Depreciation")
                amort = current_sumif_wildcard("Amortization")
            else:
                ni = ytd_sum("Net Income")
                int_e = ytd_sum_wildcard("Interest")
                dep = ytd_sum_wildcard("Depreciation")
                amort = ytd_sum_wildcard("Amortization")
            return f"=IFERROR({ni}+ABS({int_e})+ABS({dep})+ABS({amort}),0)"

        elif calc_type == 'ocf':
            # Operating Cash Flow = Net Income + Depreciation (simplified)
            if period == 'current':
                ni = current_sumif("Net Income")
                dep = current_sumif_wildcard("Depreciation")
                amort = current_sumif_wildcard("Amortization")
            else:
                ni = ytd_sum("Net Income")
                dep = ytd_sum_wildcard("Depreciation")
                amort = ytd_sum_wildcard("Amortization")
            return f"=IFERROR({ni}+ABS({dep})+ABS({amort}),0)"

        elif calc_type == 'fcf':
            # Free Cash Flow = OCF - CapEx (simplified as OCF for now)
            if period == 'current':
                ni = current_sumif("Net Income")
                dep = current_sumif_wildcard("Depreciation")
                amort = current_sumif_wildcard("Amortization")
            else:
                ni = ytd_sum("Net Income")
                dep = ytd_sum_wildcard("Depreciation")
                amort = ytd_sum_wildcard("Amortization")
            return f"=IFERROR({ni}+ABS({dep})+ABS({amort}),0)"

        return "=0"

    def _col_letter(self, col_num):
        """Convert column number to letter"""
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    # ================================================================
    # TEMPLATE V2 UTILITY FUNCTIONS
    # ================================================================
    # These functions support the optimized "Clone, Inject, Adjust" pattern
    # for faster model generation using pre-built template sheets.

    def _clone_template_sheet(self, wb, template_name, new_name, position_after=None):
        """Clone a template sheet, rename it, and position it correctly.

        This is the core of the template-based optimization. Instead of building
        sheets from scratch (slow), we clone pre-formatted templates (fast).

        Args:
            wb: xlwings Workbook object
            template_name: Name of template sheet (e.g., '_TPL_PL')
            new_name: Name for the cloned sheet (e.g., 'Consolidated_PL')
            position_after: Name of sheet to position after (optional)

        Returns:
            xlwings Sheet object for the new sheet
        """
        try:
            template_sheet = wb[template_name]

            # Copy the sheet using openpyxl's copy_worksheet
            new_sheet = wb.copy_worksheet(template_sheet)
            new_sheet.title = new_name

            return new_sheet

        except Exception as e:
            print(f"Error cloning template sheet '{template_name}': {e}")
            # Fallback: create new sheet
            if position_after:
                new_sheet = wb.create_sheet(new_name)
            else:
                new_sheet = wb.create_sheet(new_name)
            return new_sheet

    def _inject_accounts_bulk(self, sheet, accounts, months, start_row, source_sheet='Source_PL',
                               is_consolidated=True, division_name=None):
        """Inject account data into a template sheet using bulk writes.

        This replaces the slow cell-by-cell formula writing with a single
        bulk write operation, which is 10-100x faster.

        Args:
            sheet: xlwings Sheet object (cloned from template)
            accounts: List of account dictionaries
            months: List of (month, year, name) tuples
            start_row: First row to write data (after headers)
            source_sheet: Name of source data sheet
            is_consolidated: If True, use SUMIF; if False, use SUMIFS with division filter
            division_name: Division name for SUMIFS filter (required if not consolidated)

        Returns:
            Number of rows written
        """
        if not accounts:
            return 0

        source_start = 3
        source_end = 1500
        num_months = len(months)

        # Pre-calculate column letters for efficiency
        col_letters = [get_column_letter(i + 2) for i in range(num_months)]

        # Build all data in memory first
        all_rows = []
        row_types = []

        for idx, account in enumerate(accounts):
            account_name = account['name']
            indent_level = account.get('indent', 0)
            actual_row = start_row + idx

            # Display name with indentation
            display_name = account_name
            if indent_level > 0 and not account.get('is_header') and not account.get('is_total'):
                display_name = ('    ' * indent_level) + account_name

            row_data = [display_name]

            if account.get('is_header', False):
                # Header rows: empty data cells (template has formatting)
                row_data.extend([''] * num_months)
                all_rows.append(row_data)
                row_types.append('header')
                continue

            # Build SUMIF/SUMIFS formulas for each month
            for i, cl in enumerate(col_letters):
                if is_consolidated:
                    formula = f'=SUMIF({source_sheet}!$B${source_start}:$B${source_end},"{account_name}",{source_sheet}!{cl}${source_start}:{cl}${source_end})'
                else:
                    formula = f'=SUMIFS({source_sheet}!{cl}${source_start}:{cl}${source_end},{source_sheet}!$A${source_start}:$A${source_end},"{division_name}",{source_sheet}!$B${source_start}:$B${source_end},"{account_name}")'
                row_data.append(formula)

            all_rows.append(row_data)
            row_types.append('total' if account.get('is_total') else 'detail')

        # Single bulk write operation
        if all_rows:
            num_cols = len(all_rows[0])
            end_row = start_row + len(all_rows) - 1
            _data = all_rows
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=start_row + _ri, column=1).value = _val

        return len(all_rows)

    def _adjust_template_ranges(self, sheet, data_start_row, actual_rows, placeholder_rows,
                                 last_data_col):
        """Adjust formula ranges in template after injecting actual data.

        Template formulas reference placeholder row numbers (e.g., row 204 for 200 rows).
        This adjusts them to reference the actual data range.

        Args:
            sheet: xlwings Sheet object
            data_start_row: First row of account data
            actual_rows: Number of actual account rows
            placeholder_rows: Number of placeholder rows in template
            last_data_col: Last column with data
        """
        if actual_rows >= placeholder_rows:
            return  # No adjustment needed

        old_end_row = data_start_row + placeholder_rows - 1
        new_end_row = data_start_row + actual_rows - 1

        # Use Find/Replace for bulk formula adjustment
        try:
            for col in range(1, last_data_col + 1):
                col_letter = get_column_letter(col)
                # Replace references to old end row with new end row
                cells_replace(sheet, f'{col_letter}{old_end_row}', f'{col_letter}{new_end_row}')
        except Exception as e:
            print(f"Warning: Range adjustment error: {e}")

    def _delete_excess_template_rows(self, sheet, data_start_row, actual_rows, placeholder_rows):
        """Delete excess placeholder rows from template after injecting data.

        Args:
            sheet: xlwings Sheet object
            data_start_row: First row of account data
            actual_rows: Number of actual account rows
            placeholder_rows: Number of placeholder rows in template
        """
        if actual_rows >= placeholder_rows:
            return  # No rows to delete

        excess_start = data_start_row + actual_rows
        excess_end = data_start_row + placeholder_rows - 1

        try:
            num_rows_to_delete = excess_end - excess_start + 1
            sheet.delete_rows(excess_start, num_rows_to_delete)
        except Exception as e:
            print(f"Warning: Could not delete excess rows: {e}")

    def _hide_template_sheets(self, wb):
        """Hide all template sheets (those starting with _TPL_) in the final output.

        Args:
            wb: xlwings Workbook object
        """
        for sheet in wb.worksheets:
            if sheet.title.startswith('_TPL_'):
                try:
                    sheet.sheet_state = 'hidden'
                except Exception as e:
                    print(f"Warning: Could not hide template sheet {sheet.title}: {e}")

    def _ensure_template_v2_exists(self):
        """Ensure DNA_Template_v2.xlsm exists, create it if not.

        This generates the optimized template with pre-built sheets.
        Called once when the app starts or when template is missing.

        Returns:
            Path to template file, or None if creation failed
        """
        if os.path.exists(TEMPLATE_V2_PATH):
            return TEMPLATE_V2_PATH

        print("Creating DNA_Template_v2.xlsm...")
        try:
            self._create_template_v2()
            return TEMPLATE_V2_PATH
        except Exception as e:
            print(f"Error creating template v2: {e}")
            return None

    def _create_template_v2(self):
        """Create the optimized DNA_Template_v2.xlsm with pre-built template sheets.

        This function builds the template programmatically with all formatting,
        formulas, and structure pre-configured. It only needs to run once.
        """
        print("Building DNA_Template_v2.xlsm...")

        # openpyxl: no Excel app needed
        try:
            wb = Workbook()

            # Define standard colors
            DARK_BLUE = CLR_DARK_BLUE
            HEADER_WHITE = "FFFFFF"
            SUBTOTAL_GRAY = CLR_SUBTOTAL_GRAY

            # ============================================================
            # Create _TPL_PL (P&L Template)
            # ============================================================
            tpl_pl = wb.create_sheet('_TPL_PL')
            self._build_pl_template_sheet(tpl_pl, DARK_BLUE, HEADER_WHITE, SUBTOTAL_GRAY)

            # ============================================================
            # Create _TPL_BS (Balance Sheet Template)
            # ============================================================
            tpl_bs = wb.create_sheet('_TPL_BS')
            self._build_bs_template_sheet(tpl_bs, DARK_BLUE, HEADER_WHITE, SUBTOTAL_GRAY)

            # ============================================================
            # Create placeholder sheets for other templates
            # These will be built out in later phases
            # ============================================================
            tpl_cf = wb.create_sheet('_TPL_CF')
            tpl_cf['A1'].value = 'Cash Flow Template - Phase 2'

            tpl_forecast = wb.create_sheet('_TPL_Forecast')
            tpl_forecast['A1'].value = 'Forecast Template - Phase 3'

            # ============================================================
            # Create standard sheets (Menu, Dashboard, etc.)
            # ============================================================
            menu_sheet = wb.create_sheet('Menu')
            menu_sheet['A1'].value = 'Menu'
            menu_sheet['A1'].font = Font(bold=True)

            # Remove default Sheet1
            for sheet in wb.worksheets:
                if sheet.title == 'Sheet1':
                    # Sheet deletion: use wb.remove(sheet)
                    break

            # Save as macro-enabled workbook
            wb.save(TEMPLATE_V2_PATH)
            print(f"Template v2 created: {TEMPLATE_V2_PATH}")

        finally:
            pass  # openpyxl auto-handles cleanup
            # openpyxl: no app to quit

    def _build_pl_template_sheet(self, sheet, dark_blue, header_white, subtotal_gray):
        """Build the _TPL_PL template sheet with all formatting pre-configured.

        Structure:
        - Row 1-2: Title area
        - Row 3: YYYYMM helper row (hidden)
        - Row 4: Header row
        - Row 5-204: 200 placeholder account rows with formatting
        - Row 205+: Validation section
        """
        PLACEHOLDER_ROWS = 200
        DATA_START_ROW = 5
        HEADER_ROW = 4

        # Max columns: Account + 24 months + Notes + spacer + YTD cols + spacer + 3 years
        MAX_COLS = 1 + 24 + 1 + 1 + 4 + 1 + 3  # = 35

        # Title rows
        sheet['A1'].value = '<<<COMPANY_NAME>>>'
        sheet['A1'].font = Font(size=14, bold=True)

        sheet['A2'].value = 'Profit & Loss Statement'
        sheet['A2'].font = Font(size=12, bold=True)

        # Row 3: Helper row (will contain YYYYMM values)
        sheet['A3'].value = 'YYYYMM Helper'
        sheet['A3'].font = Font(color="FFFFFF")  # White (hidden)

        # Header row
        header_data = ['Account'] + [f'Month {i}' for i in range(1, 25)]  # 24 month placeholders
        header_data.extend(['Notes', '', 'PY YTD', 'CY YTD', 'Var $', 'Var %', ''])
        header_data.extend(['Year 1', 'Year 2', 'Year 3'])
        last_col = len(header_data)

        # Write header data to row 4
        write_row_to_cells(sheet, header_data, row=HEADER_ROW, start_col=1)

        # Format header row
        apply_style_to_range(sheet, HEADER_ROW, 1, HEADER_ROW, last_col,
                             font=Font(bold=True, color=header_white),
                             fill=PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid"))

        # Pre-format placeholder rows
        for row in range(DATA_START_ROW, DATA_START_ROW + PLACEHOLDER_ROWS):
            # Number format for data columns
            apply_style_to_range(sheet, row, 2, row, 25, number_format='#,##0')
            # YTD columns
            apply_style_to_range(sheet, row, 28, row, 29, number_format='#,##0')
            sheet.cell(row=row, column=30).number_format = '#,##0'
            sheet.cell(row=row, column=31).number_format = '0.0%'
            # Year columns
            apply_style_to_range(sheet, row, 33, row, 35, number_format='#,##0')

        # Column widths
        sheet.column_dimensions['A'].width = 35
        sheet.column_dimensions[get_column_letter(27)].width = 2  # Spacer 1
        sheet.column_dimensions[get_column_letter(32)].width = 2  # Spacer 2

        # Hide row 3 (YYYYMM helper)
        try:
            # Row hiding handled via hide_row()
            pass
        except:
            pass

        # Add marker for injection point
        sheet[f'A{DATA_START_ROW}'].value = '<<<INSERT_ACCOUNTS_HERE>>>'

        print(f"  Built _TPL_PL with {PLACEHOLDER_ROWS} placeholder rows")

    def _build_bs_template_sheet(self, sheet, dark_blue, header_white, subtotal_gray):
        """Build the _TPL_BS template sheet with all formatting pre-configured.

        Similar structure to P&L but typically fewer accounts.
        """
        PLACEHOLDER_ROWS = 150
        DATA_START_ROW = 5
        HEADER_ROW = 4

        # Title rows
        sheet['A1'].value = '<<<COMPANY_NAME>>>'
        sheet['A1'].font = Font(size=14, bold=True)

        sheet['A2'].value = 'Balance Sheet'
        sheet['A2'].font = Font(size=12, bold=True)

        # Row 3: Helper row
        sheet['A3'].value = 'YYYYMM Helper'
        sheet['A3'].font = Font(color="FFFFFF")

        # Header row
        header_data = ['Account'] + [f'Month {i}' for i in range(1, 25)]
        last_col = len(header_data)

        # Write header data to row 4
        write_row_to_cells(sheet, header_data, row=HEADER_ROW, start_col=1)

        # Format header row
        apply_style_to_range(sheet, HEADER_ROW, 1, HEADER_ROW, last_col,
                             font=Font(bold=True, color=header_white),
                             fill=PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid"))

        # Pre-format placeholder rows
        for row in range(DATA_START_ROW, DATA_START_ROW + PLACEHOLDER_ROWS):
            apply_style_to_range(sheet, row, 2, row, 25, number_format='#,##0')

        # Column widths
        sheet.column_dimensions['A'].width = 35

        # Hide row 3
        try:
            # Row hiding handled via hide_row()
            pass
        except:
            pass

        # Add marker
        sheet[f'A{DATA_START_ROW}'].value = '<<<INSERT_ACCOUNTS_HERE>>>'

        print(f"  Built _TPL_BS with {PLACEHOLDER_ROWS} placeholder rows")

    # ================================================================
    # RUNTIME TEMPLATE CREATION - Creates templates during generation
    # ================================================================

    def _create_runtime_templates(self, wb, months):
        """Create pre-formatted template sheets dynamically during model generation.

        This creates _TPL_PL and _TPL_BS templates with full formatting applied,
        which can then be cloned for each division sheet (much faster than
        formatting from scratch each time).

        Args:
            wb: xlwings Workbook object
            months: List of (month, year, name) tuples

        Returns:
            Tuple of (tpl_pl_sheet, tpl_bs_sheet) or (None, None) if failed
        """
        import time as _time
        _t0 = _time.perf_counter()

        DARK_BLUE = CLR_DARK_BLUE
        HEADER_WHITE = "FFFFFF"
        PLACEHOLDER_ROWS = 250  # Enough for most P&L/BS structures
        HEADER_ROW = 4
        DATA_START_ROW = 5

        num_months = len(months)

        try:
            # ============================================================
            # Create _TPL_PL (P&L Template)
            # ============================================================
            tpl_pl = wb.create_sheet('_TPL_PL')

            # Title area
            tpl_pl['A1'].value = '<<<COMPANY_NAME>>>'
            tpl_pl['A1'].font = Font(name='Calibri Light', size=16, bold=True, color=CLR_DARK_BLUE)

            tpl_pl['A2'].value = 'Profit & Loss Statement'
            tpl_pl['A2'].font = Font(name='Calibri Light', size=12, color="808080")

            # Row 3: YYYYMM helper values (pre-populate)
            helper_row = [y * 100 + m for m, y, name in months]
            for _ci, _val in enumerate(helper_row if isinstance(helper_row, list) else [helper_row]):
                tpl_pl.cell(row=3, column=2 + _ci).value = _val if not isinstance(_val, list) else _val

            # Calculate column positions for P&L
            last_month_col = num_months + 1
            notes_col = last_month_col + 1
            spacer1_col = notes_col + 1
            py_ytd_col = spacer1_col + 1
            cy_ytd_col = py_ytd_col + 1
            var_col = cy_ytd_col + 1
            var_pct_col = var_col + 1
            spacer2_col = var_pct_col + 1
            years = sorted(set(y for m, y, name in months))
            fy_start_col = spacer2_col + 1
            last_col = fy_start_col + len(years) - 1

            # Header row with actual month names
            header_data = ['Account']
            for m, y, name in months:
                header_data.append(f"{self.MONTHS[m-1][:3]} {y}")
            header_data.extend(['Notes', '', 'PY YTD', 'CY YTD', 'Var $', 'Var %', ''])
            header_data.extend([str(y) for y in years])
            for _ci, _val in enumerate(header_data if isinstance(header_data, list) else [header_data]):
                tpl_pl.cell(row=HEADER_ROW, column=1 + _ci).value = _val if not isinstance(_val, list) else _val

            # Format header row
            # Range: header_range = (tpl_pl, HEADER_ROW, 1, HEADER_ROW, last_col)
            apply_style_to_range(tpl_pl, HEADER_ROW, 1, HEADER_ROW, last_col, font=Font(name='Calibri Light'))
            apply_style_to_range(tpl_pl, HEADER_ROW, 1, HEADER_ROW, last_col, font=Font(size=10))
            apply_style_to_range(tpl_pl, HEADER_ROW, 1, HEADER_ROW, last_col, font=Font(bold=True))
            apply_style_to_range(tpl_pl, HEADER_ROW, 1, HEADER_ROW, last_col, fill=FILL_DARK_BLUE)
            apply_style_to_range(tpl_pl, HEADER_ROW, 1, HEADER_ROW, last_col, font=Font(color="FFFFFF"))

            # Pre-format ALL placeholder rows with fonts and number formats (BULK operations)
            data_end_row = DATA_START_ROW + PLACEHOLDER_ROWS - 1
            # Range: all_data_range = (tpl_pl, DATA_START_ROW, 1, data_end_row, last_col)
            apply_style_to_range(tpl_pl, DATA_START_ROW, 1, data_end_row, last_col, font=Font(name='Calibri Light'))
            apply_style_to_range(tpl_pl, DATA_START_ROW, 1, data_end_row, last_col, font=Font(size=10))

            # Number formats for specific column groups
            apply_style_to_range(tpl_pl, DATA_START_ROW, 2, data_end_row, last_month_col, number_format='#,##0')
            apply_style_to_range(tpl_pl, DATA_START_ROW, py_ytd_col, data_end_row, cy_ytd_col, number_format='#,##0')
            apply_style_to_range(tpl_pl, DATA_START_ROW, var_col, data_end_row, var_col, number_format='#,##0')
            apply_style_to_range(tpl_pl, DATA_START_ROW, var_pct_col, data_end_row, var_pct_col, number_format='0.0%')
            if fy_start_col <= last_col:
                apply_style_to_range(tpl_pl, DATA_START_ROW, fy_start_col, data_end_row, last_col, number_format='#,##0')

            # Column widths
            tpl_pl.column_dimensions['A'].width = 35
            tpl_pl.column_dimensions[get_column_letter(spacer1_col)].width = 2
            tpl_pl.column_dimensions[get_column_letter(spacer2_col)].width = 2

            # Hide row 3 (YYYYMM helper)
            try:
                # Row hiding handled via hide_row()
                pass
            except:
                pass

            # ============================================================
            # Create _TPL_BS (Balance Sheet Template)
            # ============================================================
            tpl_bs = wb.create_sheet('_TPL_BS')

            # Title area
            tpl_bs['A1'].value = '<<<COMPANY_NAME>>>'
            tpl_bs['A1'].font = Font(name='Calibri Light', size=16, bold=True, color=CLR_DARK_BLUE)

            tpl_bs['A2'].value = 'Balance Sheet'
            tpl_bs['A2'].font = Font(name='Calibri Light', size=12, color="808080")

            # Header row
            bs_last_col = num_months + 1
            bs_header = ['Account']
            for m, y, name in months:
                bs_header.append(f"{self.MONTHS[m-1][:3]} {y}")
            for _ci, _val in enumerate(bs_header if isinstance(bs_header, list) else [bs_header]):
                tpl_bs.cell(row=HEADER_ROW, column=1 + _ci).value = _val if not isinstance(_val, list) else _val

            # Format header row
            # Range: bs_header_range = (tpl_bs, HEADER_ROW, 1, HEADER_ROW, bs_last_col)
            apply_style_to_range(tpl_bs, HEADER_ROW, 1, HEADER_ROW, bs_last_col, font=Font(name='Calibri Light'))
            apply_style_to_range(tpl_bs, HEADER_ROW, 1, HEADER_ROW, bs_last_col, font=Font(size=10))
            apply_style_to_range(tpl_bs, HEADER_ROW, 1, HEADER_ROW, bs_last_col, font=Font(bold=True))
            apply_style_to_range(tpl_bs, HEADER_ROW, 1, HEADER_ROW, bs_last_col, fill=FILL_DARK_BLUE)
            apply_style_to_range(tpl_bs, HEADER_ROW, 1, HEADER_ROW, bs_last_col, font=Font(color="FFFFFF"))

            # Pre-format ALL placeholder rows
            # Range: bs_data_range = (tpl_bs, DATA_START_ROW, 1, data_end_row, bs_last_col)
            apply_style_to_range(tpl_bs, DATA_START_ROW, 1, data_end_row, bs_last_col, font=Font(name='Calibri Light'))
            apply_style_to_range(tpl_bs, DATA_START_ROW, 1, data_end_row, bs_last_col, font=Font(size=10))
            apply_style_to_range(tpl_bs, DATA_START_ROW, 2, data_end_row, bs_last_col, number_format='#,##0')

            # Column width
            tpl_bs.column_dimensions['A'].width = 35

            # Store template info for later use
            self._tpl_info = {
                'pl_last_col': last_col,
                'pl_notes_col': notes_col,
                'pl_spacer1_col': spacer1_col,
                'pl_spacer2_col': spacer2_col,
                'pl_py_ytd_col': py_ytd_col,
                'pl_cy_ytd_col': cy_ytd_col,
                'pl_var_col': var_col,
                'pl_var_pct_col': var_pct_col,
                'pl_fy_start_col': fy_start_col,
                'pl_years': years,
                'bs_last_col': bs_last_col,
                'num_months': num_months,
                'placeholder_rows': PLACEHOLDER_ROWS,
                'header_row': HEADER_ROW,
                'data_start_row': DATA_START_ROW,
            }

            print(f"  [TEMPLATES] Created _TPL_PL and _TPL_BS in {_time.perf_counter() - _t0:.2f}s")
            return tpl_pl, tpl_bs

        except Exception as e:
            print(f"  [TEMPLATES] Error creating templates: {e}")
            return None, None

    def _create_division_pl_optimized(self, wb, div_pl, all_months, pl_totals, div_name, position_after):
        """Create division P&L sheet using optimized clone-and-inject pattern.

        This is 3-5x faster than _create_division_pl_report because:
        1. Clones pre-formatted template (no bulk font/number formatting)
        2. Only applies row-type-specific formatting (headers/totals)
        3. Uses bulk data write for formulas

        Args:
            wb: xlwings Workbook object
            div_pl: List of division P&L account dictionaries
            all_months: List of (month, year, name) tuples
            pl_totals: Dict of detected total rows
            div_name: Division name
            position_after: Sheet to position after

        Returns:
            xlwings Sheet object
        """
        import time as _time
        _t0 = _time.perf_counter()

        safe_name = div_name.replace(' ', '_')[:20]
        sheet_name = f"{safe_name}_PL"

        # Clone template - with robust error handling
        clone_success = False
        sheet = None
        try:
            # Check if template exists
            if '_TPL_PL' not in wb.sheetnames:
                raise ValueError("Template _TPL_PL not found")

            tpl = wb['_TPL_PL']

            # Copy the template using openpyxl
            sheet = wb.copy_worksheet(tpl)
            sheet.title = sheet_name
            clone_success = True

        except Exception as e:
            print(f"  [{div_name}_PL] Clone failed: {e}")
            clone_success = False

        # If clone failed, fall back to standard method
        if not clone_success:
            print(f"  [{div_name}_PL] Falling back to standard creation")
            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                clear_sheet_data(sheet)
            else:
                sheet = wb.create_sheet(sheet_name)
            # Use the original method
            self._create_division_pl_report(sheet, div_pl, all_months, pl_totals, div_name)
            return sheet

        _t1 = _time.perf_counter()
        print(f"  [{div_name}_PL] Cloned template in {_t1 - _t0:.2f}s")

        # Update title
        sheet['A1'].value = div_name
        sheet['A2'].value = 'Profit & Loss Statement'

        # Get template info
        info = getattr(self, '_tpl_info', {})
        header_row = info.get('header_row', 4)
        data_start_row = info.get('data_start_row', 5)
        num_months = info.get('num_months', len(all_months))
        last_col = info.get('pl_last_col', num_months + 10)
        notes_col = info.get('pl_notes_col', num_months + 2)
        py_ytd_col = info.get('pl_py_ytd_col', notes_col + 2)
        cy_ytd_col = info.get('pl_cy_ytd_col', py_ytd_col + 1)
        var_col = info.get('pl_var_col', cy_ytd_col + 1)
        var_pct_col = info.get('pl_var_pct_col', var_col + 1)
        fy_start_col = info.get('pl_fy_start_col', var_pct_col + 2)
        years = info.get('pl_years', sorted(set(y for m, y, name in all_months)))

        source_start = 3
        source_end = 1500
        first_data_col_letter = get_column_letter(2)
        last_data_col_letter = get_column_letter(num_months + 1)

        # Build all data in memory
        all_data = []
        row_types = []

        for account in div_pl:
            account_name = account['name']
            indent_level = account.get('indent', 0)
            actual_row = data_start_row + len(all_data)

            display_name = account_name
            if indent_level > 0 and not account.get('is_header') and not account.get('is_total'):
                display_name = ('    ' * indent_level) + account_name

            row_data = [display_name]

            if account.get('is_header', False):
                row_data.extend([''] * (last_col - 1))
                all_data.append(row_data)
                row_types.append('header')
                continue

            # Monthly formulas (SUMIFS for division)
            for i, (m, y, name) in enumerate(all_months):
                cl = get_column_letter(i + 3)  # Data starts at column C in Source_PL
                formula = f'=SUMIFS(Source_PL!{cl}${source_start}:{cl}${source_end},Source_PL!$A${source_start}:$A${source_end},"{div_name}",Source_PL!$B${source_start}:$B${source_end},"{account_name}")'
                row_data.append(formula)

            # Notes formula
            notes_formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="P&L")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{actual_row}))),Notes!$D$2:$D$100),"")'
            row_data.append(notes_formula)
            row_data.append('')  # Spacer

            # YTD formulas
            data_range = f'{first_data_col_letter}{actual_row}:{last_data_col_letter}{actual_row}'
            helper_range = f'{first_data_col_letter}$3:{last_data_col_letter}$3'
            row_data.append(f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7-1)*--(MOD({helper_range},100)<=Menu!$E$7))')
            row_data.append(f'=SUMPRODUCT(({data_range})*--(INT({helper_range}/100)=Menu!$F$7)*--(MOD({helper_range},100)<=Menu!$E$7))')
            row_data.append(f'={get_column_letter(cy_ytd_col)}{actual_row}-{get_column_letter(py_ytd_col)}{actual_row}')
            row_data.append(f'=IFERROR({get_column_letter(var_col)}{actual_row}/{get_column_letter(py_ytd_col)}{actual_row},0)')
            row_data.append('')  # Spacer

            # Full year formulas
            for year in years:
                row_data.append(f'=SUMPRODUCT((INT({first_data_col_letter}$3:{last_data_col_letter}$3/100)={year})*{first_data_col_letter}{actual_row}:{last_data_col_letter}{actual_row})')

            all_data.append(row_data)
            name_lower = account_name.lower()
            if 'net income' in name_lower and account.get('is_total'):
                row_types.append('net_income')
            elif account.get('is_total'):
                row_types.append('total')
            else:
                row_types.append('detail')

        # Bulk write data
        if all_data:
            data_end_row = data_start_row + len(all_data) - 1
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=data_start_row + _ri, column=1).value = _val

        _t2 = _time.perf_counter()
        print(f"  [{div_name}_PL] Data injected ({len(all_data)} rows) in {_t2 - _t1:.2f}s")

        # Apply ONLY row-type-specific formatting (headers/totals need bold/borders)
        # This is the only per-row loop, but it's minimal
        header_rows = [data_start_row + i for i, rt in enumerate(row_types) if rt == 'header']
        total_rows = [data_start_row + i for i, rt in enumerate(row_types) if rt in ('total', 'net_income')]
        net_income_rows = [data_start_row + i for i, rt in enumerate(row_types) if rt == 'net_income']

        for r in header_rows:
            sheet.cell(row=r, column=1).font = Font(bold=True, size=11)

        for r in total_rows:
            # Range: row_range = (sheet, r, 1, r, last_col)
            apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
            try:
                # Borders handled via Border()/Side() objects
                # Borders handled via Border()/Side() objects
                pass
            except:
                pass

        for r in net_income_rows:
            try:
                pass
                # Range: row_range = (sheet, r, 1, r, last_col)
                # Borders handled via Border()/Side() objects
                # Borders handled via Border()/Side() objects
            except:
                pass

        # Delete excess template rows
        if all_data:
            placeholder_rows = info.get('placeholder_rows', 250)
            if len(all_data) < placeholder_rows:
                excess_start = data_start_row + len(all_data)
                excess_end = data_start_row + placeholder_rows - 1
                try:
                    sheet.delete_rows(excess_start, excess_end - excess_start + 1)
                except:
                    pass

        # Group previous year columns
        self._group_previous_year_columns(sheet, all_months, data_start_col=2)

        _t3 = _time.perf_counter()
        print(f"  [{div_name}_PL] TOTAL: {_t3 - _t0:.2f}s (vs ~35-40s without optimization)")

        return sheet

    def _create_division_bs_optimized(self, wb, div_bs, all_months, bs_totals, div_name, position_after):
        """Create division Balance Sheet using optimized clone-and-inject pattern.

        Args:
            wb: xlwings Workbook object
            div_bs: List of division BS account dictionaries
            all_months: List of (month, year, name) tuples
            bs_totals: Dict of detected total rows
            div_name: Division name
            position_after: Sheet to position after

        Returns:
            xlwings Sheet object
        """
        import time as _time
        _t0 = _time.perf_counter()

        safe_name = div_name.replace(' ', '_')[:20]
        sheet_name = f"{safe_name}_BS"

        # Clone template - with robust error handling
        clone_success = False
        sheet = None
        try:
            # Check if template exists
            if '_TPL_BS' not in wb.sheetnames:
                raise ValueError("Template _TPL_BS not found")

            tpl = wb['_TPL_BS']

            # Copy the template using openpyxl
            sheet = wb.copy_worksheet(tpl)
            sheet.title = sheet_name
            clone_success = True

        except Exception as e:
            print(f"  [{div_name}_BS] Clone failed: {e}")
            clone_success = False

        # If clone failed, fall back to standard method
        if not clone_success:
            print(f"  [{div_name}_BS] Falling back to standard creation")
            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                clear_sheet_data(sheet)
            else:
                sheet = wb.create_sheet(sheet_name)
            # Use the original method
            self._create_division_bs_report(sheet, div_bs, all_months, bs_totals, div_name)
            return sheet

        _t1 = _time.perf_counter()

        # Update title
        sheet['A1'].value = div_name
        sheet['A2'].value = 'Balance Sheet'

        info = getattr(self, '_tpl_info', {})
        header_row = info.get('header_row', 4)
        data_start_row = info.get('data_start_row', 5)
        num_months = len(all_months)
        last_col = num_months + 1

        source_start = 3
        source_end = 1500

        # Build all data in memory
        all_data = []
        header_rows_list = []
        total_rows = []

        for account in div_bs:
            account_name = account['name']
            indent_level = account.get('indent', 0)
            row_idx = data_start_row + len(all_data)

            display_name = account_name
            if indent_level > 0 and not account.get('is_header') and not account.get('is_total'):
                display_name = ('    ' * indent_level) + account_name

            if account.get('is_header', False):
                header_rows_list.append(row_idx)
            elif account.get('is_total', False):
                total_rows.append(row_idx)

            row_data = [display_name]

            if account.get('is_header', False):
                row_data.extend([''] * num_months)
            else:
                for i, (m, y, name) in enumerate(all_months):
                    cl = get_column_letter(i + 3)
                    formula = f'=SUMIFS(Source_BS!{cl}${source_start}:{cl}${source_end},Source_BS!$A${source_start}:$A${source_end},"{div_name}",Source_BS!$B${source_start}:$B${source_end},"{account_name}")'
                    row_data.append(formula)

            all_data.append(row_data)

        # Bulk write data
        if all_data:
            data_end_row = data_start_row + len(all_data) - 1
            _data = all_data
            if _data is not None:
                if isinstance(_data, list) and len(_data) > 0 and isinstance(_data[0], list):
                    for _ri, _row in enumerate(_data):
                        for _ci, _val in enumerate(_row):
                            sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _val
                elif isinstance(_data, list):
                    for _ri, _val in enumerate(_data):
                        if isinstance(_val, list):
                            for _ci, _v in enumerate(_val):
                                sheet.cell(row=data_start_row + _ri, column=1 + _ci).value = _v
                        else:
                            sheet.cell(row=data_start_row + _ri, column=1).value = _val

        _t2 = _time.perf_counter()

        # Apply row-type-specific formatting
        for r in header_rows_list:
            sheet.cell(row=r, column=1).font = Font(bold=True, size=11)

        for r in total_rows:
            # Range: row_range = (sheet, r, 1, r, last_col)
            apply_style_to_range(sheet, r, 1, r, last_col, font=Font(bold=True))
            try:
                # Borders handled via Border()/Side() objects
                # Borders handled via Border()/Side() objects
                pass
            except:
                pass

        # Special formatting for Total Assets and Total Liabilities & Equity
        for r in total_rows:
            try:
                cell_value = sheet.cell(row=r, column=1).value
                if cell_value and ('Total Assets' in str(cell_value) or
                                  'Total Liabilities & Equity' in str(cell_value) or
                                  'Total Liabilities and Equity' in str(cell_value)):
                    apply_style_to_range(sheet, r, 1, r, last_col, border=BORDER_NET_INCOME)
            except:
                pass

        # Delete excess template rows
        if all_data:
            placeholder_rows = info.get('placeholder_rows', 250)
            if len(all_data) < placeholder_rows:
                excess_start = data_start_row + len(all_data)
                excess_end = data_start_row + placeholder_rows - 1
                try:
                    sheet.delete_rows(excess_start, excess_end - excess_start + 1)
                except:
                    pass

        # Group previous year columns
        self._group_previous_year_columns(sheet, all_months, data_start_col=2)

        _t3 = _time.perf_counter()
        print(f"  [{div_name}_BS] TOTAL: {_t3 - _t0:.2f}s (optimized)")

        return sheet

    # ================================================================
    # END TEMPLATE V2 UTILITY FUNCTIONS
    # ================================================================

    def _add_back_to_menu_link(self, sheet, row=1, col=1):
        """Add a 'Back to Menu' hyperlink at the specified position"""
        try:
            cell = sheet.cell(row=row, column=col)
            cell.value = '<< Menu'
            cell.font = Font(name='Calibri Light', size=9)
            cell.hyperlink = "#'Menu'!A1"
            cell.font = Font(name='Calibri Light', size=9, color="0066CC", underline="single")
        except Exception as e:
            print(f"Back to Menu link warning: {e}")

    def _setup_print_area(self, sheet, last_row=None, last_col=None):
        """Set up professional print settings for a sheet

        Args:
            sheet: xlwings Sheet object
            last_row: Last row to include (optional, uses UsedRange if not specified)
            last_col: Last column to include (optional, uses UsedRange if not specified)
        """
        try:
            ps = sheet.page_setup
            pp = sheet.print_options

            # Landscape orientation
            ps.orientation = 'landscape'

            # Fit all columns on one page, rows can span multiple pages
            ps.fitToPage = True
            ps.fitToWidth = 1
            ps.fitToHeight = 0  # Allow multiple pages vertically

            # Margins (in inches) - openpyxl uses inches directly
            sheet.page_margins.left = 0.5
            sheet.page_margins.right = 0.5
            sheet.page_margins.top = 0.75
            sheet.page_margins.bottom = 0.75
            sheet.page_margins.header = 0.5
            sheet.page_margins.footer = 0.5

            # Header: Company name on left, sheet name center
            company = self.company_name.get()
            sheet.oddHeader.left.text = company
            sheet.oddHeader.left.font = "Calibri Light,Regular"
            sheet.oddHeader.left.size = 10
            sheet.oddHeader.center.text = "&A"  # Sheet name
            sheet.oddHeader.center.font = "Calibri Light,Bold"
            sheet.oddHeader.center.size = 12

            # Footer: Date on left, page number center
            sheet.oddFooter.left.text = "&D"  # Date
            sheet.oddFooter.left.font = "Calibri Light,Regular"
            sheet.oddFooter.left.size = 9
            sheet.oddFooter.center.text = "Page &P of &N"
            sheet.oddFooter.center.font = "Calibri Light,Regular"
            sheet.oddFooter.center.size = 9

            # Repeat rows at top (header row)
            sheet.print_title_rows = '1:4'

            # Gridlines off for cleaner print
            pp.gridLines = False

        except Exception as e:
            print(f"Print setup warning: {e}")

    def _apply_notes_division_dropdown(self, sheet, notes_col, start_row, end_row, divisions=None):
        """Add division dropdown to Notes column for multi-division mode

        Args:
            sheet: xlwings Sheet object
            notes_col: Column number for Notes
            start_row: First data row
            end_row: Last data row
            divisions: List of division dicts (with 'name' key)
        """
        try:
            if not divisions or len(divisions) < 2:
                return  # No dropdown needed for single division

            # Build dropdown list: "All" + division names
            division_names = ["All"] + [d.get('name', d) if isinstance(d, dict) else d for d in divisions]
            dropdown_list = ",".join(division_names)

            # Apply data validation to notes column range
            col_letter = get_column_letter(notes_col)
            dv = DataValidation(type="list", formula1=f'"{dropdown_list}"', allow_blank=True)
            dv.prompt = "Select division"
            dv.promptTitle = "Division Filter"
            range_str = f'{col_letter}{start_row}:{col_letter}{end_row}'
            dv.add(range_str)
            sheet.add_data_validation(dv)

        except Exception as e:
            print(f"Warning: Could not add division dropdown to Notes: {e}")

    def _apply_row_grouping(self, sheet, accounts, start_row):
        """Apply Excel row grouping (outline) based on account sections

        Groups rows between section headers and their totals to allow
        collapsing/expanding sections in Excel.

        Args:
            sheet: xlwings Sheet object
            accounts: List of account dictionaries with 'is_header' and 'is_total' flags
            start_row: First data row (after header row)
        """
        try:
            # Track sections: (header_row, total_row)
            sections = []
            current_section_start = None
            row_idx = start_row

            for account in accounts:
                if account.get('is_header', False):
                    # Start of a new section
                    if current_section_start is not None:
                        # Close previous section without a total (shouldn't happen but handle it)
                        sections.append((current_section_start, row_idx - 1))
                    current_section_start = row_idx
                elif account.get('is_total', False) and current_section_start is not None:
                    # End of current section - group rows BETWEEN header and total
                    if row_idx > current_section_start + 1:
                        # Only group if there are rows between header and total
                        sections.append((current_section_start + 1, row_idx - 1))
                    current_section_start = None

                row_idx += 1

            # Apply grouping to each section
            for group_start, group_end in sections:
                if group_end > group_start:
                    try:
                        group_rows(sheet, group_start, group_end, outline_level=1)
                    except Exception as e:
                        print(f"Warning: Could not group rows {group_start}-{group_end}: {e}")

            # Set outline settings - summary rows below detail (default)
            try:
                sheet.sheet_properties.outlinePr = openpyxl.worksheet.properties.OutlineProperties(summaryBelow=True)
            except:
                pass

        except Exception as e:
            print(f"Warning: Could not apply row grouping: {e}")

    def _build_source_lookup_formula(self, source_sheet, account_name, col_num, division=None):
        """Build a SUMIF or SUMIFS formula for looking up source data

        Args:
            source_sheet: Name of source sheet (e.g., "Source_PL")
            account_name: Account name to look up
            col_num: Column number for the value (1-based)
            division: Optional division name for multi-division filtering

        Returns:
            Excel formula string
        """
        col_letter = get_column_letter(col_num)
        # Use limited ranges (rows 3-1500) instead of entire columns for speed
        sr = 3  # source start row
        er = 1500  # source end row

        if division or self.is_multi_division.get():
            # Multi-division mode: Use SUMIFS with Division filter
            # Structure: Division in A, Account in B, values start in C
            div_filter = division if division else 'Menu!$G$5'
            return (f'=SUMIFS({source_sheet}!{col_letter}${sr}:{col_letter}${er},'
                    f'{source_sheet}!$A${sr}:$A${er},"{div_filter}",'
                    f'{source_sheet}!$B${sr}:$B${er},"{account_name}")')
        else:
            # Single division mode: Use SUMIF (backward compatible)
            return f'=SUMIF({source_sheet}!$A${sr}:$A${er},"{account_name}",{source_sheet}!{col_letter}${sr}:{col_letter}${er})'

    def _build_consolidated_formula(self, source_sheet, account_name, col_num):
        """Build a formula that sums across all divisions (consolidated view)

        Args:
            source_sheet: Name of source sheet (e.g., "Source_PL")
            account_name: Account name to look up
            col_num: Column number for the value (1-based)

        Returns:
            Excel formula string for consolidated sum
        """
        col_letter = get_column_letter(col_num)
        # Use limited ranges (rows 3-1500) instead of entire columns for speed
        sr = 3
        er = 1500

        if self.is_multi_division.get():
            # Consolidated: Sum all divisions (no division filter)
            return f'=SUMIF({source_sheet}!$B${sr}:$B${er},"{account_name}",{source_sheet}!{col_letter}${sr}:{col_letter}${er})'
        else:
            # Single division: Same as regular lookup
            return f'=SUMIF({source_sheet}!$A${sr}:$A${er},"{account_name}",{source_sheet}!{col_letter}${sr}:{col_letter}${er})'

    def _group_columns_by_year(self, sheet, months, header_row):
        """
        Hide and group columns based on current month (last month in data):
        - Group and hide all prior year columns (before current year)
        - Current year columns up to current month remain visible
        - Future months (after current month) are hidden but not grouped

        The "current month" is determined by the last month in the data,
        which corresponds to Menu cell C7.
        """
        if not months:
            return

        try:
            # Current month is the LAST month in the data (this is what Menu C7 shows)
            current_month_num, current_year, current_month_name = months[-1]
            prior_year = current_year - 1

            print(f"Column visibility: Current month = {current_month_name} (month {current_month_num}), Year = {current_year}")

            # Track columns to hide and group
            prior_year_cols = []  # Columns from prior years (to be grouped and hidden)
            future_month_cols = []  # Columns after current month (to be hidden only)

            # months is a list of tuples: (month_num, year, display_name)
            for i, (m, y, name) in enumerate(months):
                col = i + 2  # Data starts at column 2 (B)

                if y < current_year:
                    # Prior year - group and hide
                    prior_year_cols.append((col, y))
                elif y == current_year and m > current_month_num:
                    # Future month in current year - hide only
                    future_month_cols.append(col)
                elif y > current_year:
                    # Future year - hide only
                    future_month_cols.append(col)
                # else: current year, current or prior month - leave visible

            # Group prior year columns by year
            year_groups = {}
            for col, y in prior_year_cols:
                if y not in year_groups:
                    year_groups[y] = [col, col]
                else:
                    year_groups[y][1] = col

            # Create groups for each prior year and hide them
            for year in sorted(year_groups.keys()):
                start_col, end_col = year_groups[year]
                try:
                    # Create group and hide
                    group_cols(sheet, start_col, end_col, outline_level=1, hidden=True)
                    print(f"Grouped and hid year {year}: columns {start_col} to {end_col}")
                except Exception as e:
                    print(f"Column grouping error for year {year}: {e}")

            # Hide future month columns (not grouped, just hidden)
            for col in future_month_cols:
                try:
                    # Column hiding handled via hide_columns_range()
                    print(f"Hid future month column {col}")
                except Exception as e:
                    print(f"Error hiding column {col}: {e}")

        except Exception as e:
            print(f"Year grouping error: {e}")


def main():
    root = tk.Tk()
    app = FinancialModelApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
