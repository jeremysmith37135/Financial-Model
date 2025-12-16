"""
Financial Model Excel Generator
Generates a CFO-grade financial model from P&L and Balance Sheet data
"""

import io
import os
import re
import zipfile
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference


class FinancialModelGenerator:
    """Generates CFO-grade financial model Excel workbook"""

    MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    def __init__(self, company_name: str, fiscal_year_start: int,
                 first_display_month: int, first_display_year: int,
                 pl_data: pd.DataFrame, bs_data: pd.DataFrame):
        self.company_name = company_name
        self.fiscal_year_start = fiscal_year_start
        self.first_display_month = first_display_month
        self.first_display_year = first_display_year
        self.pl_raw = pl_data
        self.bs_raw = bs_data

        self.pl_accounts: List[Dict] = []
        self.bs_accounts: List[Dict] = []
        self.months: List[Tuple[int, int]] = []

        self.wb: Optional[Workbook] = None
        self._define_styles()

    def _define_styles(self):
        """Define reusable styles"""
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        self.bottom_border = Border(bottom=Side(style='thin'))
        self.top_border = Border(top=Side(style='thin'))
        self.double_bottom = Border(bottom=Side(style='double'))

        # Fills
        self.header_fill = PatternFill(start_color='16213E', end_color='16213E', fill_type='solid')
        self.total_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
        self.subtotal_fill = PatternFill(start_color='ECECEC', end_color='ECECEC', fill_type='solid')
        self.source_fill = PatternFill(start_color='1A1A1A', end_color='1A1A1A', fill_type='solid')
        self.light_blue_fill = PatternFill(start_color='E6F2FF', end_color='E6F2FF', fill_type='solid')

        # Fonts
        self.header_font = Font(name='Calibri Light', size=10, bold=True, color='FFFFFF')
        self.title_font = Font(name='Calibri Light', size=14, bold=True)
        self.subtitle_font = Font(name='Calibri Light', size=12, bold=True)
        self.bold_font = Font(name='Calibri Light', size=10, bold=True)
        self.normal_font = Font(name='Calibri Light', size=10)
        self.small_font = Font(name='Calibri Light', size=9, color='666666')

        # Alignments
        self.right_align = Alignment(horizontal='right', vertical='center')
        self.left_align = Alignment(horizontal='left', vertical='center')
        self.center_align = Alignment(horizontal='center', vertical='center')
        self.wrap_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # Dashboard-specific styles
        self.section_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        self.kpi_header_fill = PatternFill(start_color='34495E', end_color='34495E', fill_type='solid')
        self.alt_row_fill = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
        self.green_fill = PatternFill(start_color='27AE60', end_color='27AE60', fill_type='solid')
        self.yellow_fill = PatternFill(start_color='F39C12', end_color='F39C12', fill_type='solid')
        self.red_fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
        self.control_fill = PatternFill(start_color='E8F4FD', end_color='E8F4FD', fill_type='solid')

        self.section_font = Font(name='Calibri Light', size=14, bold=True, color='FFFFFF')
        self.kpi_name_font = Font(name='Calibri Light', size=11, bold=True, color='2C3E50')
        self.kpi_value_font = Font(name='Calibri Light', size=14, bold=True, color='16213E')

    def _auto_fit_column(self, ws, column: str, min_width: float = 8, max_width: float = 60, padding: float = 2):
        """Auto-fit column width based on content"""
        max_length = min_width
        col_letter = column if isinstance(column, str) else get_column_letter(column)

        for cell in ws[col_letter]:
            if cell.value:
                # Get the display value
                cell_value = str(cell.value)
                # Skip formulas - use a reasonable estimate
                if cell_value.startswith('='):
                    continue
                # Calculate length (account for font size differences)
                cell_length = len(cell_value)
                # Account for bold text being slightly wider
                if cell.font and cell.font.bold:
                    cell_length *= 1.1
                max_length = max(max_length, cell_length)

        # Apply padding and cap at max
        final_width = min(max_length + padding, max_width)
        ws.column_dimensions[col_letter].width = final_width
        return final_width

    def _auto_fit_columns(self, ws, columns: List[str], min_width: float = 8, max_width: float = 60, padding: float = 2):
        """Auto-fit multiple columns"""
        for col in columns:
            self._auto_fit_column(ws, col, min_width, max_width, padding)

    def generate(self) -> io.BytesIO:
        """Generate the Excel workbook"""
        self._parse_pl_data()
        self._parse_bs_data()

        self.wb = Workbook()
        default_sheet = self.wb.active
        self.wb.remove(default_sheet)

        # Create sheets
        self._create_menu_sheet()
        self._create_source_pl_sheet()
        self._create_source_bs_sheet()
        self._create_pl_sheet()
        self._create_bs_sheet()
        self._create_cash_flow_sheet()
        self._create_helper_notes_sheet()
        self._create_diagnostics_sheet()
        self._create_vba_instructions_sheet()

        # Create Dashboard sheets
        self._create_dashboard_control_sheet()
        self._create_dashboard_sheet()

        # Reorder sheets: Dashboard first, Menu at end before source sheets
        # Sheet order will be: Dashboard, P&L, Balance Sheet, Cash Flow, Helper_Notes,
        # Diagnostics, VBA Instructions, Dashboard_Control, Menu, Source P&L, Source BS
        self._reorder_sheets()

        # Set Dashboard as the active sheet when file opens
        self.wb.active = self.wb['Dashboard']

        # Save and return
        buffer = io.BytesIO()
        self.wb.save(buffer)
        buffer.seek(0)

        return buffer

    def _reorder_sheets(self):
        """Reorder sheets to put Dashboard first and Menu at the end"""
        # Define desired order: Dashboard first, Menu before source sheets
        desired_order = [
            'Dashboard',
            'P&L',
            'Balance Sheet',
            'Cash Flow',
            'Helper_Notes',
            'Diagnostics',
            'VBA Instructions',
            'Dashboard_Control',
            'Menu',
            'Source P&L',
            'Source BS'
        ]

        # Create new order based on what sheets exist
        current_sheets = [ws.title for ws in self.wb.worksheets]
        new_order = []

        for sheet_name in desired_order:
            if sheet_name in current_sheets:
                new_order.append(sheet_name)

        # Add any sheets not in our desired order at the end
        for sheet_name in current_sheets:
            if sheet_name not in new_order:
                new_order.append(sheet_name)

        # Reorder by moving sheets
        for i, sheet_name in enumerate(new_order):
            ws = self.wb[sheet_name]
            self.wb.move_sheet(ws, offset=i - self.wb.index(ws))

    def _parse_month_header(self, header) -> Optional[Tuple[int, int]]:
        """Parse month header like 'January 2024' into (month_num, year)"""
        if pd.isna(header):
            return None

        header = str(header).strip()

        # Skip range formats like "January-December, 2024"
        if '-' in header and any(m.lower() in header.lower() for m in self.MONTHS[1:]):
            return None

        # Match "Month Year" format exactly
        for i, month in enumerate(self.MONTHS, 1):
            if header.lower().startswith(month.lower()):
                parts = header.split()
                if len(parts) == 2:  # Exactly "Month Year"
                    try:
                        year = int(parts[-1])
                        if 2000 <= year <= 2100:  # Sanity check
                            return (i, year)
                    except ValueError:
                        pass

        # Try MM/YYYY or M/YYYY format
        match = re.match(r'^(\d{1,2})[/-](\d{4})$', header)
        if match:
            month_num = int(match.group(1))
            if 1 <= month_num <= 12:
                return (month_num, int(match.group(2)))

        return None

    def _parse_pl_data(self):
        """Parse P&L data"""
        df = self.pl_raw.copy()

        header_row = None
        for idx in range(min(15, len(df))):
            row = df.iloc[idx]
            # Check each column for a month header
            for col_idx in range(len(row)):
                val = row.iloc[col_idx] if hasattr(row, 'iloc') else row[col_idx]
                if self._parse_month_header(val):
                    header_row = idx
                    break
            if header_row is not None:
                break

        if header_row is None:
            raise ValueError("Could not find month headers in P&L file")

        self.months = []
        month_cols = {}

        for col_idx in range(1, len(df.columns)):
            val = df.iloc[header_row, col_idx]
            parsed = self._parse_month_header(val)
            if parsed and pd.notna(val) and 'total' not in str(val).lower():
                self.months.append(parsed)
                month_cols[col_idx] = parsed

        data_start = header_row + 1
        for row_idx in range(data_start, len(df)):
            account_name = df.iloc[row_idx, 0]

            if pd.isna(account_name) or str(account_name).strip() == '':
                continue

            account_name = str(account_name).strip()

            if 'cash basis' in account_name.lower():
                continue

            is_total = (account_name.lower().startswith('total for') or
                       account_name == 'Net Income' or
                       account_name == 'Gross Profit' or
                       account_name == 'Net Operating Income' or
                       account_name == 'Net Other Income')
            is_header = (account_name in ['Income', 'Cost of Sales', 'Expense', 'Expenses',
                                         'Other Income', 'Other Expense', 'Other Expenses'] or
                        (any(cat in account_name for cat in ['Income', 'Expense', 'Cost'])
                         and not is_total and not any(c.isdigit() for c in account_name[:4])))

            values = {}
            for col_idx, (m, y) in month_cols.items():
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val):
                    try:
                        values[(m, y)] = float(val)
                    except (ValueError, TypeError):
                        values[(m, y)] = 0.0
                else:
                    values[(m, y)] = 0.0

            self.pl_accounts.append({
                'name': account_name,
                'is_total': is_total,
                'is_header': is_header,
                'values': values
            })

    def _parse_bs_data(self):
        """Parse Balance Sheet data"""
        df = self.bs_raw.copy()

        header_row = None
        for idx in range(min(15, len(df))):
            row = df.iloc[idx]
            for col_idx in range(len(row)):
                val = row.iloc[col_idx] if hasattr(row, 'iloc') else row[col_idx]
                if self._parse_month_header(val):
                    header_row = idx
                    break
            if header_row is not None:
                break

        if header_row is None:
            raise ValueError("Could not find month headers in Balance Sheet file")

        month_cols = {}
        for col_idx in range(1, len(df.columns)):
            val = df.iloc[header_row, col_idx]
            parsed = self._parse_month_header(val)
            if parsed and pd.notna(val):
                month_cols[col_idx] = parsed

        data_start = header_row + 1
        for row_idx in range(data_start, len(df)):
            account_name = df.iloc[row_idx, 0]

            if pd.isna(account_name) or str(account_name).strip() == '':
                continue

            account_name = str(account_name).strip()

            if 'cash basis' in account_name.lower():
                continue

            is_total = (account_name.lower().startswith('total for') or
                       account_name in ['Total for Assets', 'Total for Liabilities',
                                       'Total for Equity', 'Total for Liabilities and Equity'])
            is_header = account_name in ['Assets', 'Liabilities', 'Equity', 'Current Assets',
                                        'Fixed Assets', 'Other Assets', 'Current Liabilities',
                                        'Long-term Liabilities', 'Bank Accounts',
                                        'Accounts Receivable', 'Other Current Assets',
                                        'Accounts Payable', 'Other Current Liabilities']

            values = {}
            for col_idx, (m, y) in month_cols.items():
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val):
                    try:
                        values[(m, y)] = float(val)
                    except (ValueError, TypeError):
                        values[(m, y)] = 0.0
                else:
                    values[(m, y)] = 0.0

            self.bs_accounts.append({
                'name': account_name,
                'is_total': is_total,
                'is_header': is_header,
                'values': values
            })

    def _create_menu_sheet(self):
        """Create Menu/Control sheet"""
        ws = self.wb.create_sheet("Menu", 0)
        ws.sheet_view.showGridLines = False

        # Title
        ws['B2'] = self.company_name
        ws['B2'].font = Font(name='Calibri Light', size=28, bold=True, color='16213E')
        ws.merge_cells('B2:G2')

        ws['B3'] = "Financial Model"
        ws['B3'].font = Font(name='Calibri Light', size=16, color='666666')
        ws.merge_cells('B3:G3')

        ws['B4'] = f"Generated: {datetime.now().strftime('%B %d, %Y')}"
        ws['B4'].font = self.small_font
        ws.merge_cells('B4:G4')

        # Configuration Section
        ws['B7'] = "MODEL CONFIGURATION"
        ws['B7'].font = self.header_font
        ws['B7'].fill = self.header_fill
        ws.merge_cells('B7:D7')
        for col in range(2, 5):
            ws.cell(row=7, column=col).fill = self.header_fill

        config_items = [
            ('Company Name:', self.company_name),
            ('Fiscal Year Start:', self.MONTHS[self.fiscal_year_start - 1]),
            ('Current Month:', f"{self.MONTHS[self.months[-1][0] - 1]} {self.months[-1][1]}" if self.months else ''),
            ('First Display Month:', f"{self.MONTHS[self.first_display_month - 1]} {self.first_display_year}"),
        ]

        for idx, (label, value) in enumerate(config_items):
            row = 9 + idx
            ws[f'B{row}'] = label
            ws[f'B{row}'].font = self.bold_font
            ws[f'C{row}'] = value
            ws[f'C{row}'].font = self.normal_font
            ws[f'C{row}'].alignment = self.left_align

        # Current month dropdown
        if self.months:
            month_options = [f"{self.MONTHS[m-1]} {y}" for m, y in self.months]
            dv = DataValidation(type="list", formula1=f'"{",".join(month_options[-24:])}"', allow_blank=False)
            dv.add(ws['C11'])
            ws.add_data_validation(dv)

        # Create named ranges for config
        self._create_named_range('config_company', 'Menu', 'C9')
        self._create_named_range('config_fy_start', 'Menu', 'C10')
        self._create_named_range('config_current_month', 'Menu', 'C11')
        self._create_named_range('config_first_display', 'Menu', 'C12')

        # Actions Section
        ws['B15'] = "ACTIONS"
        ws['B15'].font = self.header_font
        ws['B15'].fill = self.header_fill
        ws.merge_cells('B15:D15')
        for col in range(2, 5):
            ws.cell(row=15, column=col).fill = self.header_fill

        actions = [
            ('Upload New P&L', 'Run macro: UploadPLFile'),
            ('Upload New Balance Sheet', 'Run macro: UploadBSFile'),
            ('Refresh All Formulas', 'Run macro: RefreshAllFormulas'),
            ('Save Variance Note', 'Run macro: SaveVarianceNote'),
            ('Run Diagnostics', 'Run macro: RunDiagnostics'),
        ]

        for idx, (action, desc) in enumerate(actions):
            row = 17 + idx
            ws[f'B{row}'] = action
            ws[f'B{row}'].font = self.bold_font
            ws[f'C{row}'] = desc
            ws[f'C{row}'].font = self.small_font

        # Navigation Section
        ws['B24'] = "QUICK NAVIGATION"
        ws['B24'].font = self.header_font
        ws['B24'].fill = self.header_fill
        ws.merge_cells('B24:D24')
        for col in range(2, 5):
            ws.cell(row=24, column=col).fill = self.header_fill

        nav_items = [
            ('P&L', "'P&L'!A1"),
            ('Balance Sheet', "'Balance Sheet'!A1"),
            ('Cash Flow', "'Cash Flow'!A1"),
            ('Source P&L', "'Source P&L'!A1"),
            ('Source BS', "'Source BS'!A1"),
            ('Diagnostics', "'Diagnostics'!A1"),
        ]

        for idx, (name, ref) in enumerate(nav_items):
            row = 26 + idx
            ws[f'B{row}'] = name
            ws[f'B{row}'].font = Font(name='Calibri Light', size=10, color='0563C1', underline='single')
            ws[f'B{row}'].hyperlink = f"#{ref}"

        # Instructions
        ws['B34'] = "INSTRUCTIONS"
        ws['B34'].font = self.header_font
        ws['B34'].fill = self.header_fill
        ws.merge_cells('B34:D34')
        for col in range(2, 5):
            ws.cell(row=34, column=col).fill = self.header_fill

        instructions = [
            "1. To add VBA macros: See 'VBA Instructions' sheet",
            "2. Upload new monthly data using the Upload actions",
            "3. Change Current Month to update variance notes display",
            "4. Use navigation links to move between sheets",
        ]

        for idx, inst in enumerate(instructions):
            row = 36 + idx
            ws[f'B{row}'] = inst
            ws[f'B{row}'].font = self.normal_font

        # Column widths
        ws.column_dimensions['A'].width = 3
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 15

        # Hidden config values
        ws['H3'] = self.fiscal_year_start
        ws['H4'] = self.first_display_month
        ws['H5'] = self.first_display_year
        ws.column_dimensions['H'].hidden = True

    def _create_named_range(self, name: str, sheet: str, cell: str):
        """Create a named range"""
        from openpyxl.workbook.defined_name import DefinedName
        # Handle cell references that may or may not have $ signs
        if ':' in cell:
            # Range reference like A1:M100
            parts = cell.split(':')
            ref = f"'{sheet}'!${parts[0].replace('$', '')}:${parts[1].replace('$', '')}"
        else:
            # Single cell reference
            ref = f"'{sheet}'!${cell.replace('$', '')}"
        defn = DefinedName(name=name, attr_text=ref)
        self.wb.defined_names[name] = defn

    def _create_source_pl_sheet(self):
        """Create Source P&L sheet"""
        ws = self.wb.create_sheet("Source P&L")
        ws.sheet_properties.tabColor = "000000"

        # Header
        ws['A1'] = "Account"
        ws['A1'].font = self.header_font
        ws['A1'].fill = self.source_fill

        for col_idx, (m, y) in enumerate(self.months, 2):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = f"{self.MONTHS[m-1]} {y}"
            cell.font = self.header_font
            cell.fill = self.source_fill
            cell.alignment = self.center_align

        # Data
        for row_idx, account in enumerate(self.pl_accounts, 2):
            ws.cell(row=row_idx, column=1).value = account['name']
            ws.cell(row=row_idx, column=1).font = self.normal_font

            for col_idx, (m, y) in enumerate(self.months, 2):
                val = account['values'].get((m, y), 0)
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = val
                cell.number_format = '#,##0'
                cell.font = self.normal_font

        # Column widths - auto-fit account name column
        self._auto_fit_column(ws, 'A', min_width=45, max_width=70)
        for col_idx in range(2, len(self.months) + 2):
            ws.column_dimensions[get_column_letter(col_idx)].width = 14

        # Named range
        last_row = len(self.pl_accounts) + 1
        last_col = get_column_letter(len(self.months) + 1)
        self._create_named_range('SourcePL_Data', 'Source P&L', f'A1:{last_col}{last_row}')

    def _create_source_bs_sheet(self):
        """Create Source BS sheet"""
        ws = self.wb.create_sheet("Source BS")
        ws.sheet_properties.tabColor = "000000"

        ws['A1'] = "Account"
        ws['A1'].font = self.header_font
        ws['A1'].fill = self.source_fill

        for col_idx, (m, y) in enumerate(self.months, 2):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = f"{self.MONTHS[m-1]} {y}"
            cell.font = self.header_font
            cell.fill = self.source_fill
            cell.alignment = self.center_align

        for row_idx, account in enumerate(self.bs_accounts, 2):
            ws.cell(row=row_idx, column=1).value = account['name']
            ws.cell(row=row_idx, column=1).font = self.normal_font

            for col_idx, (m, y) in enumerate(self.months, 2):
                val = account['values'].get((m, y), 0)
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = val
                cell.number_format = '#,##0'
                cell.font = self.normal_font

        # Column widths - auto-fit account name column
        self._auto_fit_column(ws, 'A', min_width=45, max_width=70)
        for col_idx in range(2, len(self.months) + 2):
            ws.column_dimensions[get_column_letter(col_idx)].width = 14

        last_row = len(self.bs_accounts) + 1
        last_col = get_column_letter(len(self.months) + 1)
        self._create_named_range('SourceBS_Data', 'Source BS', f'A1:{last_col}{last_row}')

    def _create_pl_sheet(self):
        """Create P&L report sheet"""
        ws = self.wb.create_sheet("P&L")

        # Title
        ws['A1'] = self.company_name
        ws['A1'].font = self.title_font
        ws.merge_cells('A1:E1')

        ws['A2'] = "Profit & Loss Statement"
        ws['A2'].font = self.subtitle_font
        ws.merge_cells('A2:E2')

        # Header row
        header_row = 4
        ws.cell(row=header_row, column=1).value = "Account"
        ws.cell(row=header_row, column=1).font = self.header_font
        ws.cell(row=header_row, column=1).fill = self.header_fill
        ws.cell(row=header_row, column=1).alignment = self.left_align

        # Month columns
        for col_idx, (m, y) in enumerate(self.months, 2):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.value = f"{self.MONTHS[m-1][:3]} {y}"
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.center_align

        # Summary columns
        spacer_col = len(self.months) + 2
        notes_col = spacer_col + 1
        ytd_start = notes_col + 2

        ws.cell(row=header_row, column=notes_col).value = "Notes"
        ws.cell(row=header_row, column=notes_col).font = self.header_font
        ws.cell(row=header_row, column=notes_col).fill = self.header_fill

        summary_headers = ['PY YTD', 'CY YTD', 'Var $', 'Var %']
        for idx, header in enumerate(summary_headers):
            cell = ws.cell(row=header_row, column=ytd_start + idx)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.center_align

        # Annual columns
        years = sorted(set(y for m, y in self.months))
        annual_start = ytd_start + len(summary_headers) + 1
        for idx, year in enumerate(years):
            cell = ws.cell(row=header_row, column=annual_start + idx)
            cell.value = f"FY {year}"
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.center_align

        # Data rows
        data_row = header_row + 1
        for account in self.pl_accounts:
            ws.cell(row=data_row, column=1).value = account['name']

            # Apply formatting
            if account['is_header']:
                ws.cell(row=data_row, column=1).font = self.bold_font
                ws.cell(row=data_row, column=1).border = self.top_border
            elif account['is_total']:
                ws.cell(row=data_row, column=1).font = self.bold_font
                fill = self.total_fill if 'net income' in account['name'].lower() else self.subtotal_fill
                for col in range(1, annual_start + len(years)):
                    ws.cell(row=data_row, column=col).fill = fill
                    ws.cell(row=data_row, column=col).border = self.top_border
            else:
                ws.cell(row=data_row, column=1).font = self.normal_font

            # Monthly data with VLOOKUP
            for col_idx, (m, y) in enumerate(self.months, 2):
                cell = ws.cell(row=data_row, column=col_idx)
                formula = f'=IFERROR(VLOOKUP($A{data_row},SourcePL_Data,{col_idx},FALSE),0)'
                cell.value = formula
                cell.number_format = '#,##0'
                cell.alignment = self.right_align
                if account['is_total']:
                    cell.font = self.bold_font

            # Notes formula with text wrap
            notes_cell = ws.cell(row=data_row, column=notes_col)
            notes_formula = f'=IFERROR(INDEX(Helper_Notes!$C:$C,MATCH($A{data_row}&config_current_month,Helper_Notes!$D:$D,0)),"")'
            notes_cell.value = notes_formula
            notes_cell.font = self.normal_font
            notes_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

            # YTD calculations
            first_col = get_column_letter(2)
            last_col = get_column_letter(len(self.months) + 1)

            # CY YTD
            cy_cell = ws.cell(row=data_row, column=ytd_start + 1)
            cy_cell.value = f'=SUM({first_col}{data_row}:{last_col}{data_row})'
            cy_cell.number_format = '#,##0'

            # PY YTD (placeholder)
            py_cell = ws.cell(row=data_row, column=ytd_start)
            py_cell.value = 0
            py_cell.number_format = '#,##0'

            # Variance $
            var_cell = ws.cell(row=data_row, column=ytd_start + 2)
            var_cell.value = f'={get_column_letter(ytd_start + 1)}{data_row}-{get_column_letter(ytd_start)}{data_row}'
            var_cell.number_format = '#,##0'

            # Variance %
            var_pct = ws.cell(row=data_row, column=ytd_start + 3)
            var_pct.value = f'=IFERROR({get_column_letter(ytd_start + 2)}{data_row}/{get_column_letter(ytd_start)}{data_row},0)'
            var_pct.number_format = '0.0%'

            # Annual totals
            for idx, year in enumerate(years):
                year_cols = [c for c, (m, y) in enumerate(self.months, 2) if y == year]
                if year_cols:
                    annual_cell = ws.cell(row=data_row, column=annual_start + idx)
                    refs = '+'.join([f'{get_column_letter(c)}{data_row}' for c in year_cols])
                    annual_cell.value = f'={refs}'
                    annual_cell.number_format = '#,##0'

            data_row += 1

        # Validation section
        val_row = data_row + 2
        ws.cell(row=val_row, column=1).value = "VALIDATION"
        ws.cell(row=val_row, column=1).font = self.bold_font

        ws.cell(row=val_row + 1, column=1).value = "Source P&L Total Check"
        ws.cell(row=val_row + 1, column=1).font = self.normal_font

        # Column widths - auto-fit account name column
        self._auto_fit_column(ws, 'A', min_width=45, max_width=70)
        for col in range(2, annual_start + len(years) + 1):
            if col == notes_col:
                # Notes column - wider to accommodate wrapped text
                ws.column_dimensions[get_column_letter(col)].width = 35
            else:
                ws.column_dimensions[get_column_letter(col)].width = 12

        # Group earlier months
        first_display_idx = self._get_first_display_index()
        if first_display_idx and first_display_idx > 0:
            ws.column_dimensions.group(
                get_column_letter(2),
                get_column_letter(first_display_idx + 1),
                hidden=True
            )

    def _create_bs_sheet(self):
        """Create Balance Sheet report sheet"""
        ws = self.wb.create_sheet("Balance Sheet")

        ws['A1'] = self.company_name
        ws['A1'].font = self.title_font
        ws.merge_cells('A1:E1')

        ws['A2'] = "Balance Sheet"
        ws['A2'].font = self.subtitle_font
        ws.merge_cells('A2:E2')

        header_row = 4
        ws.cell(row=header_row, column=1).value = "Account"
        ws.cell(row=header_row, column=1).font = self.header_font
        ws.cell(row=header_row, column=1).fill = self.header_fill

        for col_idx, (m, y) in enumerate(self.months, 2):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.value = f"{self.MONTHS[m-1][:3]} {y}"
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.center_align

        data_row = header_row + 1
        for account in self.bs_accounts:
            ws.cell(row=data_row, column=1).value = account['name']

            if account['is_header']:
                ws.cell(row=data_row, column=1).font = self.bold_font
                ws.cell(row=data_row, column=1).border = self.top_border
            elif account['is_total']:
                ws.cell(row=data_row, column=1).font = self.bold_font
                fill = self.total_fill if account['name'] in ['Total for Assets', 'Total for Liabilities and Equity'] else self.subtotal_fill
                for col in range(1, len(self.months) + 2):
                    ws.cell(row=data_row, column=col).fill = fill
            else:
                ws.cell(row=data_row, column=1).font = self.normal_font

            for col_idx, (m, y) in enumerate(self.months, 2):
                cell = ws.cell(row=data_row, column=col_idx)
                formula = f'=IFERROR(VLOOKUP($A{data_row},SourceBS_Data,{col_idx},FALSE),0)'
                cell.value = formula
                cell.number_format = '#,##0'
                cell.alignment = self.right_align

            data_row += 1

        # Validation
        val_row = data_row + 2
        ws.cell(row=val_row, column=1).value = "VALIDATION: Assets = L + E"
        ws.cell(row=val_row, column=1).font = self.bold_font

        # Column widths - auto-fit account name column
        self._auto_fit_column(ws, 'A', min_width=45, max_width=70)
        for col in range(2, len(self.months) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 14

        first_display_idx = self._get_first_display_index()
        if first_display_idx and first_display_idx > 0:
            ws.column_dimensions.group(
                get_column_letter(2),
                get_column_letter(first_display_idx + 1),
                hidden=True
            )

    def _create_cash_flow_sheet(self):
        """Create Cash Flow Statement sheet"""
        ws = self.wb.create_sheet("Cash Flow")

        ws['A1'] = self.company_name
        ws['A1'].font = self.title_font
        ws.merge_cells('A1:E1')

        ws['A2'] = "Cash Flow Statement"
        ws['A2'].font = self.subtitle_font
        ws.merge_cells('A2:E2')

        header_row = 4
        ws.cell(row=header_row, column=1).value = "Description"
        ws.cell(row=header_row, column=1).font = self.header_font
        ws.cell(row=header_row, column=1).fill = self.header_fill

        for col_idx, (m, y) in enumerate(self.months, 2):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.value = f"{self.MONTHS[m-1][:3]} {y}"
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.center_align

        # YTD column
        ytd_col = len(self.months) + 3
        ws.cell(row=header_row, column=ytd_col).value = "YTD"
        ws.cell(row=header_row, column=ytd_col).font = self.header_font
        ws.cell(row=header_row, column=ytd_col).fill = self.header_fill

        # Cash flow line items
        cf_structure = [
            ("CASH FLOWS FROM OPERATING ACTIVITIES", True, False, None),
            ("Net income", False, False, "PL_NetIncome"),
            ("Depreciation & amortization", False, False, None),
            ("Changes in working capital:", True, False, None),
            ("  (Increase) decrease in accounts receivable", False, False, "BS_AR_Change"),
            ("  (Increase) decrease in other current assets", False, False, "BS_OCA_Change"),
            ("  Increase (decrease) in accounts payable", False, False, "BS_AP_Change"),
            ("  Increase (decrease) in other current liabilities", False, False, "BS_OCL_Change"),
            ("Net cash from operating activities", False, True, "CF_Operating"),
            ("", False, False, None),
            ("CASH FLOWS FROM INVESTING ACTIVITIES", True, False, None),
            ("Capital expenditures", False, False, "BS_FixedAssets_Change"),
            ("Other investing activities", False, False, None),
            ("Net cash from investing activities", False, True, "CF_Investing"),
            ("", False, False, None),
            ("CASH FLOWS FROM FINANCING ACTIVITIES", True, False, None),
            ("Increase (decrease) in line of credit", False, False, None),
            ("Increase (decrease) in long-term debt", False, False, "BS_LTD_Change"),
            ("Shareholder distributions", False, False, "BS_Distributions"),
            ("Net cash from financing activities", False, True, "CF_Financing"),
            ("", False, False, None),
            ("NET CHANGE IN CASH", False, True, "CF_NetChange"),
            ("Beginning cash balance", False, False, "BS_BeginningCash"),
            ("Ending cash balance", False, True, "BS_EndingCash"),
            ("", False, False, None),
            ("VALIDATION: Ending cash = BS Cash", False, False, "CF_Validation"),
        ]

        data_row = header_row + 1
        for item_name, is_header, is_subtotal, calc_ref in cf_structure:
            cell = ws.cell(row=data_row, column=1)
            cell.value = item_name

            if is_header:
                cell.font = self.bold_font
                cell.border = self.top_border
            elif is_subtotal:
                cell.font = self.bold_font
                for col in range(1, ytd_col + 1):
                    ws.cell(row=data_row, column=col).fill = self.subtotal_fill
                    ws.cell(row=data_row, column=col).border = self.top_border
            else:
                cell.font = self.normal_font

            # Add placeholder values
            if not is_header and item_name:
                for col_idx in range(2, len(self.months) + 2):
                    val_cell = ws.cell(row=data_row, column=col_idx)
                    val_cell.value = 0
                    val_cell.number_format = '#,##0'
                    val_cell.alignment = self.right_align

                # YTD formula
                ytd_cell = ws.cell(row=data_row, column=ytd_col)
                first_col = get_column_letter(2)
                last_col = get_column_letter(len(self.months) + 1)
                ytd_cell.value = f'=SUM({first_col}{data_row}:{last_col}{data_row})'
                ytd_cell.number_format = '#,##0'

            data_row += 1

        # Column widths - auto-fit description column
        self._auto_fit_column(ws, 'A', min_width=45, max_width=70)
        for col in range(2, ytd_col + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14

    def _create_helper_notes_sheet(self):
        """Create Helper Notes sheet"""
        ws = self.wb.create_sheet("Helper_Notes")
        ws.sheet_properties.tabColor = "666666"

        headers = ['Account', 'Month', 'Note', 'Lookup Key']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.header_fill

        row = 2
        sample_accounts = [a for a in self.pl_accounts if not a['is_header']][:15]
        for account in sample_accounts:
            for m, y in self.months[-3:]:
                ws.cell(row=row, column=1).value = account['name']
                ws.cell(row=row, column=2).value = f"{self.MONTHS[m-1]} {y}"
                ws.cell(row=row, column=3).value = ""
                ws.cell(row=row, column=4).value = f'=A{row}&B{row}'
                row += 1

        # Column widths - auto-fit account column
        self._auto_fit_column(ws, 'A', min_width=45, max_width=70)
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 60
        ws.column_dimensions['D'].width = 55

    def _create_diagnostics_sheet(self):
        """Create Diagnostics sheet"""
        ws = self.wb.create_sheet("Diagnostics")
        ws.sheet_properties.tabColor = "FF6600"

        ws['A1'] = "Model Diagnostics"
        ws['A1'].font = self.title_font

        ws['A3'] = "DATA SUMMARY"
        ws['A3'].font = self.bold_font

        summary = [
            ("P&L Accounts:", len(self.pl_accounts)),
            ("Balance Sheet Accounts:", len(self.bs_accounts)),
            ("Months in Data:", len(self.months)),
            ("Date Range:", f"{self.MONTHS[self.months[0][0]-1]} {self.months[0][1]} - {self.MONTHS[self.months[-1][0]-1]} {self.months[-1][1]}" if self.months else "N/A"),
            ("Model Generated:", datetime.now().strftime('%Y-%m-%d %H:%M')),
        ]

        for idx, (label, value) in enumerate(summary):
            ws.cell(row=5 + idx, column=1).value = label
            ws.cell(row=5 + idx, column=1).font = self.normal_font
            ws.cell(row=5 + idx, column=2).value = value
            ws.cell(row=5 + idx, column=2).font = self.bold_font

        ws['A12'] = "VALIDATION CHECKS"
        ws['A12'].font = self.bold_font

        checks = [
            "P&L Net Income matches Balance Sheet",
            "Balance Sheet balances (Assets = L + E)",
            "Cash Flow ending balance = BS Cash",
            "All months have data",
        ]

        for idx, check in enumerate(checks):
            ws.cell(row=14 + idx, column=1).value = check
            ws.cell(row=14 + idx, column=2).value = "Run diagnostics to check"
            ws.cell(row=14 + idx, column=2).font = self.small_font

        ws['A20'] = "ERROR LOG"
        ws['A20'].font = self.bold_font

        ws['A21'] = "Timestamp"
        ws['B21'] = "Message"
        ws['C21'] = "Status"
        for col in range(1, 4):
            ws.cell(row=21, column=col).font = self.header_font
            ws.cell(row=21, column=col).fill = self.header_fill

        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 50
        ws.column_dimensions['C'].width = 20

    def _create_vba_instructions_sheet(self):
        """Create VBA Instructions sheet"""
        ws = self.wb.create_sheet("VBA Instructions")
        ws.sheet_properties.tabColor = "990000"

        ws['A1'] = "VBA Macro Installation Instructions"
        ws['A1'].font = self.title_font

        instructions = [
            "",
            "This workbook requires VBA macros for full functionality.",
            "Follow these steps to add the macros:",
            "",
            "STEP 1: Save as Macro-Enabled Workbook",
            "  - File > Save As",
            "  - Choose 'Excel Macro-Enabled Workbook (*.xlsm)'",
            "  - Note: The current .xlsx cannot contain macros",
            "",
            "STEP 2: Open VBA Editor",
            "  - Press Alt + F11",
            "  - Or: Developer tab > Visual Basic",
            "",
            "STEP 3: Import the VBA Module",
            "  - In VBA Editor: File > Import File",
            "  - Select the 'vba_code.bas' file (provided separately)",
            "  - Or: Insert > Module, then paste the code",
            "",
            "STEP 4: Enable Macros",
            "  - When opening the file, click 'Enable Content'",
            "  - Or: File > Options > Trust Center > Macro Settings",
            "",
            "AVAILABLE MACROS:",
            "  - UploadPLFile: Import new P&L data",
            "  - UploadBSFile: Import new Balance Sheet data",
            "  - RefreshAllFormulas: Recalculate all sheets",
            "  - SaveVarianceNote: Save a variance note",
            "  - RunDiagnostics: Check for data issues",
            "",
            "To run a macro:",
            "  - Press Alt + F8",
            "  - Select the macro name",
            "  - Click 'Run'",
        ]

        for idx, line in enumerate(instructions):
            ws.cell(row=3 + idx, column=1).value = line
            if line.startswith("STEP") or line.startswith("AVAILABLE"):
                ws.cell(row=3 + idx, column=1).font = self.bold_font
            else:
                ws.cell(row=3 + idx, column=1).font = self.normal_font

        ws.column_dimensions['A'].width = 70

    def _get_first_display_index(self) -> Optional[int]:
        """Get index of first month to display"""
        for idx, (m, y) in enumerate(self.months):
            if y == self.first_display_year and m >= self.first_display_month:
                return idx
            elif y > self.first_display_year:
                return idx
        return None

    # =========================================================================
    # DASHBOARD MODULE METHODS
    # =========================================================================

    # KPI Definitions - CFO-standard metrics and ratios
    KPI_DEFINITIONS = {
        'profitability': [
            {'name': 'Revenue', 'formula_type': 'direct', 'source': 'PL', 'account': 'Total for Income',
             'description': 'Total revenue generated', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Gross Profit', 'formula_type': 'direct', 'source': 'PL', 'account': 'Gross Profit',
             'description': 'Revenue minus cost of goods sold', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Gross Margin %', 'formula_type': 'ratio', 'numerator': 'Gross Profit', 'denominator': 'Total for Income',
             'description': 'Gross profit as percentage of revenue', 'default_target': 0.40, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'Net Income', 'formula_type': 'direct', 'source': 'PL', 'account': 'Net Income',
             'description': 'Bottom line profit after all expenses', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Net Margin %', 'formula_type': 'ratio', 'numerator': 'Net Income', 'denominator': 'Total for Income',
             'description': 'Net income as percentage of revenue', 'default_target': 0.10, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'EBITDA', 'formula_type': 'calculated', 'calc_type': 'ebitda',
             'description': 'Earnings before interest, taxes, depreciation, amortization', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'EBITDA Margin %', 'formula_type': 'ratio', 'numerator_calc': 'ebitda', 'denominator': 'Total for Income',
             'description': 'EBITDA as percentage of revenue', 'default_target': 0.15, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'Operating Income', 'formula_type': 'direct', 'source': 'PL', 'account': 'Net Operating Income',
             'description': 'Income from core operations', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Operating Margin %', 'formula_type': 'ratio', 'numerator': 'Net Operating Income', 'denominator': 'Total for Income',
             'description': 'Operating income as percentage of revenue', 'default_target': 0.12, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'Revenue Growth %', 'formula_type': 'growth', 'metric': 'Total for Income',
             'description': 'Period-over-period revenue growth rate', 'default_target': 0.10, 'format': '0.0%', 'higher_is_better': True},
        ],
        'liquidity': [
            {'name': 'Current Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities',
             'description': 'Current assets / current liabilities', 'default_target': 1.5, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Quick Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities',
             'description': '(Current assets - inventory) / current liabilities', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Cash Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Bank Accounts', 'denominator': 'Total for Current Liabilities',
             'description': 'Cash / current liabilities', 'default_target': 0.2, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Working Capital', 'formula_type': 'bs_difference', 'minuend': 'Total for Current Assets', 'subtrahend': 'Total for Current Liabilities',
             'description': 'Current assets minus current liabilities', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Cash Balance', 'formula_type': 'direct', 'source': 'BS', 'account': 'Total for Bank Accounts',
             'description': 'Total cash and bank accounts', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
        ],
        'efficiency': [
            {'name': 'AR Days (DSO)', 'formula_type': 'days_ratio', 'balance': 'Total for Accounts Receivable', 'flow': 'Total for Income',
             'description': 'Average days to collect receivables', 'default_target': 45, 'format': '0', 'higher_is_better': False},
            {'name': 'AP Days (DPO)', 'formula_type': 'days_ratio', 'balance': 'Total for Credit Cards', 'flow': 'Total for Cost of Sales',
             'description': 'Average days to pay payables', 'default_target': 30, 'format': '0', 'higher_is_better': True},
            {'name': 'Asset Turnover', 'formula_type': 'turnover', 'flow': 'Total for Income', 'balance': 'Total for Assets',
             'description': 'Revenue generated per dollar of assets', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': True},
            {'name': 'Cash Conversion Cycle', 'formula_type': 'calculated', 'calc_type': 'ccc',
             'description': 'Days from cash outflow to cash inflow', 'default_target': 30, 'format': '0', 'higher_is_better': False},
        ],
        'leverage': [
            {'name': 'Debt-to-Equity', 'formula_type': 'bs_ratio', 'numerator': 'Total for Liabilities', 'denominator': 'Total for Equity',
             'description': 'Total debt relative to equity', 'default_target': 1.0, 'format': '0.00', 'higher_is_better': False},
            {'name': 'Debt-to-Assets', 'formula_type': 'bs_ratio', 'numerator': 'Total for Liabilities', 'denominator': 'Total for Assets',
             'description': 'Percentage of assets financed by debt', 'default_target': 0.5, 'format': '0.0%', 'higher_is_better': False},
            {'name': 'Equity Ratio', 'formula_type': 'bs_ratio', 'numerator': 'Total for Equity', 'denominator': 'Total for Assets',
             'description': 'Percentage of assets financed by equity', 'default_target': 0.5, 'format': '0.0%', 'higher_is_better': True},
            {'name': 'Interest Coverage', 'formula_type': 'coverage', 'earnings': 'Net Operating Income', 'interest': '8000 Interest Expense',
             'description': 'Operating income / interest expense', 'default_target': 3.0, 'format': '0.0', 'higher_is_better': True},
        ],
        'cashflow': [
            {'name': 'Operating Cash Flow', 'formula_type': 'calculated', 'calc_type': 'ocf',
             'description': 'Cash generated from operations', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Free Cash Flow', 'formula_type': 'calculated', 'calc_type': 'fcf',
             'description': 'Operating cash flow minus capital expenditures', 'default_target': 0, 'format': '#,##0', 'higher_is_better': True},
            {'name': 'Cash Burn Rate', 'formula_type': 'calculated', 'calc_type': 'burn_rate',
             'description': 'Monthly cash consumption rate', 'default_target': 0, 'format': '#,##0', 'higher_is_better': False},
            {'name': 'Cash Runway (Months)', 'formula_type': 'calculated', 'calc_type': 'runway',
             'description': 'Months of cash remaining at current burn rate', 'default_target': 12, 'format': '0.0', 'higher_is_better': True},
        ],
    }

    def _create_dashboard_control_sheet(self):
        """Create the Dashboard Control page with KPI target settings"""
        ws = self.wb.create_sheet("Dashboard_Control")
        ws.sheet_properties.tabColor = "3498DB"
        ws.sheet_view.showGridLines = False

        # Title
        ws['B2'] = "Dashboard Control Panel"
        ws['B2'].font = Font(name='Calibri Light', size=24, bold=True, color='16213E')
        ws.merge_cells('B2:F2')

        ws['B3'] = "Configure KPI targets and thresholds"
        ws['B3'].font = self.small_font
        ws.merge_cells('B3:F3')

        # Instructions
        ws['B5'] = "INSTRUCTIONS"
        ws['B5'].font = self.section_font
        ws['B5'].fill = self.section_fill
        for col in range(2, 7):
            ws.cell(5, col).fill = self.section_fill

        instructions = [
            "1. Set target values for each KPI in the 'Target' column",
            "2. Yellow thresholds trigger caution alerts (% of target)",
            "3. Red thresholds trigger warning alerts (% of target)",
            "4. KPIs with 'User Defined' targets require your input",
            "5. Standard targets are industry benchmarks - adjust as needed"
        ]
        for i, instr in enumerate(instructions):
            ws.cell(7 + i, 2).value = instr
            ws.cell(7 + i, 2).font = self.normal_font

        # Headers
        header_row = 14
        ws.cell(header_row, 2).value = "KPI TARGETS AND THRESHOLDS"
        ws.cell(header_row, 2).font = self.section_font
        ws.cell(header_row, 2).fill = self.section_fill
        for col in range(2, 9):
            ws.cell(header_row, col).fill = self.section_fill

        headers = ['Category', 'KPI Name', 'Description', 'Target', 'Yellow %', 'Red %', 'Direction']
        header_row = 16
        for col, header in enumerate(headers, 2):
            cell = ws.cell(header_row, col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.kpi_header_fill
            cell.alignment = self.center_align
            cell.border = self.thin_border

        # Populate KPIs
        data_row = 17
        for category, kpis in self.KPI_DEFINITIONS.items():
            for kpi in kpis:
                ws.cell(data_row, 2).value = category.title()
                ws.cell(data_row, 2).font = self.normal_font

                ws.cell(data_row, 3).value = kpi['name']
                ws.cell(data_row, 3).font = self.bold_font

                ws.cell(data_row, 4).value = kpi['description']
                ws.cell(data_row, 4).font = self.small_font
                ws.cell(data_row, 4).alignment = self.wrap_align

                target_cell = ws.cell(data_row, 5)
                target_cell.value = kpi['default_target']
                if '%' in kpi['format']:
                    target_cell.number_format = '0.0%'
                elif '0.0' in kpi['format']:
                    target_cell.number_format = '0.00'
                else:
                    target_cell.number_format = '#,##0'
                target_cell.fill = self.control_fill
                target_cell.alignment = self.right_align
                target_cell.border = self.thin_border

                ws.cell(data_row, 6).value = 0.80
                ws.cell(data_row, 6).number_format = '0%'
                ws.cell(data_row, 6).fill = self.control_fill
                ws.cell(data_row, 6).alignment = self.center_align
                ws.cell(data_row, 6).border = self.thin_border

                ws.cell(data_row, 7).value = 0.60
                ws.cell(data_row, 7).number_format = '0%'
                ws.cell(data_row, 7).fill = self.control_fill
                ws.cell(data_row, 7).alignment = self.center_align
                ws.cell(data_row, 7).border = self.thin_border

                ws.cell(data_row, 8).value = "Higher" if kpi['higher_is_better'] else "Lower"
                ws.cell(data_row, 8).font = self.normal_font
                ws.cell(data_row, 8).alignment = self.center_align

                if data_row % 2 == 0:
                    for col in range(2, 9):
                        if not ws.cell(data_row, col).fill.start_color.rgb or ws.cell(data_row, col).fill.start_color.rgb == '00000000':
                            ws.cell(data_row, col).fill = self.alt_row_fill

                data_row += 1

        # Column widths - auto-fit KPI name and description columns
        ws.column_dimensions['A'].width = 3
        ws.column_dimensions['B'].width = 14
        self._auto_fit_column(ws, 'C', min_width=22, max_width=35)  # KPI Name
        self._auto_fit_column(ws, 'D', min_width=45, max_width=60)  # Description
        ws.column_dimensions['E'].width = 14
        ws.column_dimensions['F'].width = 10
        ws.column_dimensions['G'].width = 10
        ws.column_dimensions['H'].width = 10

        # Named range for targets
        self._create_named_range('KPI_Targets', 'Dashboard_Control', f'C17:E{data_row-1}')

    def _create_dashboard_sheet(self):
        """Create the main Dashboard sheet with KPIs and visualizations"""
        ws = self.wb.create_sheet("Dashboard", 0)
        ws.sheet_properties.tabColor = "27AE60"
        ws.sheet_view.showGridLines = False

        # Title
        ws['B2'] = self.company_name
        ws['B2'].font = Font(name='Calibri Light', size=28, bold=True, color='16213E')
        ws.merge_cells('B2:L2')

        ws['B3'] = "Executive Dashboard"
        ws['B3'].font = Font(name='Calibri Light', size=16, color='7F8C8D')
        ws.merge_cells('B3:L3')

        if self.months:
            current_month = self.months[-1]
            ws['B4'] = f"Current Period: {self.MONTHS[current_month[0]-1]} {current_month[1]}"
        else:
            ws['B4'] = "Current Period: N/A"
        ws['B4'].font = self.small_font

        ws['B5'] = f"Generated: {datetime.now().strftime('%B %d, %Y')}"
        ws['B5'].font = self.small_font

        # Summary Cards
        self._create_summary_cards(ws, 7)

        # KPI Sections
        current_row = 17
        sections = [
            ('PROFITABILITY METRICS', 'profitability'),
            ('LIQUIDITY METRICS', 'liquidity'),
            ('EFFICIENCY METRICS', 'efficiency'),
            ('LEVERAGE METRICS', 'leverage'),
            ('CASH FLOW METRICS', 'cashflow'),
        ]

        for section_title, category in sections:
            current_row = self._create_kpi_section(ws, section_title, category, current_row)
            current_row += 2

        # Trend data section
        current_row = self._create_trend_data_section(ws, current_row)

        # Column widths - auto-fit KPI names column for better visibility
        ws.column_dimensions['A'].width = 3    # Margin
        self._auto_fit_column(ws, 'B', min_width=20, max_width=35)  # KPI names - auto-fit
        ws.column_dimensions['C'].width = 12   # Current values
        ws.column_dimensions['D'].width = 10   # Targets
        ws.column_dimensions['E'].width = 6    # Status
        ws.column_dimensions['F'].width = 2    # Spacer
        ws.column_dimensions['G'].width = 12   # YTD values
        ws.column_dimensions['H'].width = 10   # YTD Targets
        ws.column_dimensions['I'].width = 6    # YTD Status
        ws.column_dimensions['J'].width = 2    # Spacer
        ws.column_dimensions['K'].width = 12   # All-Time
        ws.column_dimensions['L'].width = 6    # Trend
        ws.column_dimensions['M'].width = 10   # Extra

        # Set vertical centering for all cells
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.alignment:
                    cell.alignment = Alignment(
                        horizontal=cell.alignment.horizontal or 'center',
                        vertical='center',
                        wrap_text=cell.alignment.wrap_text
                    )
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        # Print setup
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1

    def _create_summary_cards(self, ws, row_start: int):
        """Create summary metric cards at top of dashboard"""
        cards = [
            {'title': 'Revenue', 'metric': 'Total for Income', 'source': 'PL', 'format': '"$"#,##0', 'col': 2},
            {'title': 'Net Income', 'metric': 'Net Income', 'source': 'PL', 'format': '"$"#,##0', 'col': 4},
            {'title': 'Gross Margin', 'type': 'ratio', 'numerator': 'Gross Profit', 'denominator': 'Total for Income', 'format': '0.0%', 'col': 6},
            {'title': 'Cash Balance', 'metric': 'Total for Bank Accounts', 'source': 'BS', 'format': '"$"#,##0', 'col': 8},
            {'title': 'Current Ratio', 'type': 'bs_ratio', 'numerator': 'Total for Current Assets', 'denominator': 'Total for Current Liabilities', 'format': '0.00', 'col': 10},
            {'title': 'AR Days', 'type': 'days', 'balance': 'Total for Accounts Receivable', 'flow': 'Total for Income', 'format': '0', 'col': 12},
        ]

        for card in cards:
            col = card['col']
            current_month_col = len(self.months) + 1 if self.months else 2

            # Header
            ws.cell(row_start, col).value = card['title']
            ws.cell(row_start, col).font = Font(name='Calibri Light', size=10, bold=True, color='7F8C8D')
            ws.cell(row_start, col).alignment = self.center_align
            ws.merge_cells(start_row=row_start, start_column=col, end_row=row_start, end_column=col+1)

            # Value
            value_cell = ws.cell(row_start + 1, col)
            value_cell.font = Font(name='Calibri Light', size=20, bold=True, color='16213E')
            value_cell.alignment = self.center_align
            ws.merge_cells(start_row=row_start+1, start_column=col, end_row=row_start+1, end_column=col+1)

            if 'metric' in card:
                if card['source'] == 'PL':
                    formula = f"=SUMIF('Source P&L'!$A:$A,\"{card['metric']}\",'Source P&L'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                else:
                    formula = f"=SUMIF('Source BS'!$A:$A,\"{card['metric']}\",'Source BS'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
            elif card.get('type') == 'ratio':
                num = f"SUMIF('Source P&L'!$A:$A,\"{card['numerator']}\",'Source P&L'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                den = f"SUMIF('Source P&L'!$A:$A,\"{card['denominator']}\",'Source P&L'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                formula = f"=IFERROR({num}/{den},0)"
            elif card.get('type') == 'bs_ratio':
                num = f"SUMIF('Source BS'!$A:$A,\"{card['numerator']}\",'Source BS'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                den = f"SUMIF('Source BS'!$A:$A,\"{card['denominator']}\",'Source BS'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                formula = f"=IFERROR({num}/{den},0)"
            elif card.get('type') == 'days':
                bal = f"SUMIF('Source BS'!$A:$A,\"{card['balance']}\",'Source BS'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                flow = f"SUMIF('Source P&L'!$A:$A,\"{card['flow']}\",'Source P&L'!{get_column_letter(current_month_col)}:{get_column_letter(current_month_col)})"
                formula = f"=IFERROR({bal}/{flow}*30,0)"
            else:
                formula = "=0"

            value_cell.value = formula
            value_cell.number_format = card['format']

            # Period label
            ws.cell(row_start + 2, col).value = "Current Month"
            ws.cell(row_start + 2, col).font = self.small_font
            ws.cell(row_start + 2, col).alignment = self.center_align
            ws.merge_cells(start_row=row_start+2, start_column=col, end_row=row_start+2, end_column=col+1)

    def _create_kpi_section(self, ws, title: str, category: str, row_start: int) -> int:
        """Create a KPI section with current, YTD, and all-time views"""
        kpis = self.KPI_DEFINITIONS.get(category, [])
        if not kpis:
            return row_start

        current_month_col = len(self.months) + 1 if self.months else 2

        # Section header
        ws.cell(row_start, 2).value = title
        ws.cell(row_start, 2).font = self.section_font
        ws.cell(row_start, 2).fill = self.section_fill
        for col in range(2, 14):
            ws.cell(row_start, col).fill = self.section_fill

        # Column headers
        row_start += 1
        headers = [('KPI', 2), ('Current', 3), ('Target', 4), ('Status', 5), ('', 6),
                   ('YTD', 7), ('Target', 8), ('Status', 9), ('', 10),
                   ('All-Time', 11), ('Trend', 12)]

        for header, col in headers:
            cell = ws.cell(row_start, col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.kpi_header_fill
            cell.alignment = self.center_align

        # KPI rows
        data_row = row_start + 1
        control_row = 17  # Starting row in Dashboard_Control

        for kpi in kpis:
            # KPI Name
            ws.cell(data_row, 2).value = kpi['name']
            ws.cell(data_row, 2).font = self.kpi_name_font
            ws.cell(data_row, 2).alignment = self.left_align

            # Current value formula
            current_formula = self._build_kpi_formula(kpi, 'current', current_month_col)
            ws.cell(data_row, 3).value = current_formula
            ws.cell(data_row, 3).number_format = kpi['format']
            ws.cell(data_row, 3).font = self.kpi_value_font
            ws.cell(data_row, 3).alignment = self.right_align

            # Target (from control sheet)
            ws.cell(data_row, 4).value = f"=Dashboard_Control!E{control_row}"
            ws.cell(data_row, 4).number_format = kpi['format']
            ws.cell(data_row, 4).font = self.normal_font
            ws.cell(data_row, 4).alignment = self.right_align

            # Status
            status_formula = self._build_status_formula(data_row, 3, 4, kpi['higher_is_better'])
            ws.cell(data_row, 5).value = status_formula
            ws.cell(data_row, 5).font = Font(name='Calibri Light', size=12, bold=True)
            ws.cell(data_row, 5).alignment = self.center_align

            # YTD value
            ytd_formula = self._build_kpi_formula(kpi, 'ytd', current_month_col)
            ws.cell(data_row, 7).value = ytd_formula
            ws.cell(data_row, 7).number_format = kpi['format']
            ws.cell(data_row, 7).font = self.bold_font
            ws.cell(data_row, 7).alignment = self.right_align

            # YTD Target
            if '%' in kpi['format'] or 'ratio' in kpi['formula_type']:
                ws.cell(data_row, 8).value = f"=D{data_row}"
            else:
                months_count = len(self.months) if self.months else 1
                ws.cell(data_row, 8).value = f"=D{data_row}*{months_count}"
            ws.cell(data_row, 8).number_format = kpi['format']
            ws.cell(data_row, 8).font = self.normal_font
            ws.cell(data_row, 8).alignment = self.right_align

            # YTD Status
            ytd_status_formula = self._build_status_formula(data_row, 7, 8, kpi['higher_is_better'])
            ws.cell(data_row, 9).value = ytd_status_formula
            ws.cell(data_row, 9).font = Font(name='Calibri Light', size=12, bold=True)
            ws.cell(data_row, 9).alignment = self.center_align

            # All-Time
            alltime_formula = self._build_kpi_formula(kpi, 'alltime', current_month_col)
            ws.cell(data_row, 11).value = alltime_formula
            ws.cell(data_row, 11).number_format = kpi['format']
            ws.cell(data_row, 11).font = self.normal_font
            ws.cell(data_row, 11).alignment = self.right_align

            # Trend
            trend_formula = f'=IF(C{data_row}>G{data_row}*1.05,"UP",IF(C{data_row}<G{data_row}*0.95,"DOWN","FLAT"))'
            ws.cell(data_row, 12).value = trend_formula
            ws.cell(data_row, 12).font = Font(name='Calibri Light', size=10)
            ws.cell(data_row, 12).alignment = self.center_align

            # Alternate row fill
            if data_row % 2 == 0:
                for col in [2, 3, 4, 5, 7, 8, 9, 11, 12]:
                    ws.cell(data_row, col).fill = self.alt_row_fill

            data_row += 1
            control_row += 1

        # Add conditional formatting for status columns
        self._add_status_formatting(ws, row_start + 2, data_row - 1, 5)
        self._add_status_formatting(ws, row_start + 2, data_row - 1, 9)
        self._add_trend_formatting(ws, row_start + 2, data_row - 1, 12)

        return data_row

    def _build_kpi_formula(self, kpi: dict, period: str, current_col: int) -> str:
        """Build formula for a KPI based on its type and period"""
        formula_type = kpi['formula_type']

        if formula_type == 'direct':
            source = kpi['source']
            account = kpi['account']
            sheet = "'Source P&L'" if source == 'PL' else "'Source BS'"
            if period == 'current':
                return f"=SUMIF({sheet}!$A:$A,\"{account}\",{sheet}!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:  # ytd or alltime
                return f"=SUMIF({sheet}!$A:$A,\"{account}\",{sheet}!$B:{get_column_letter(current_col)})"

        elif formula_type == 'ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            if period == 'current':
                num_f = f"SUMIF('Source P&L'!$A:$A,\"{num}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                den_f = f"SUMIF('Source P&L'!$A:$A,\"{den}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:
                num_f = f"SUMIF('Source P&L'!$A:$A,\"{num}\",'Source P&L'!$B:{get_column_letter(current_col)})"
                den_f = f"SUMIF('Source P&L'!$A:$A,\"{den}\",'Source P&L'!$B:{get_column_letter(current_col)})"
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            num_f = f"SUMIF('Source BS'!$A:$A,\"{num}\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            den_f = f"SUMIF('Source BS'!$A:$A,\"{den}\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_difference':
            min_acct = kpi['minuend']
            sub_acct = kpi['subtrahend']
            min_f = f"SUMIF('Source BS'!$A:$A,\"{min_acct}\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            sub_f = f"SUMIF('Source BS'!$A:$A,\"{sub_acct}\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            return f"={min_f}-{sub_f}"

        elif formula_type == 'days_ratio':
            balance = kpi['balance']
            flow = kpi['flow']
            bal_f = f"SUMIF('Source BS'!$A:$A,\"{balance}\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            if period == 'current':
                flow_f = f"SUMIF('Source P&L'!$A:$A,\"{flow}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                return f"=IFERROR({bal_f}/{flow_f}*30,0)"
            else:
                flow_f = f"SUMIF('Source P&L'!$A:$A,\"{flow}\",'Source P&L'!$B:{get_column_letter(current_col)})"
                months = len(self.months) if self.months else 1
                return f"=IFERROR({bal_f}/(({flow_f})/{months})*30,0)"

        elif formula_type == 'turnover':
            flow = kpi['flow']
            balance = kpi['balance']
            if period == 'current':
                flow_f = f"SUMIF('Source P&L'!$A:$A,\"{flow}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:
                flow_f = f"SUMIF('Source P&L'!$A:$A,\"{flow}\",'Source P&L'!$B:{get_column_letter(current_col)})"
            bal_f = f"SUMIF('Source BS'!$A:$A,\"{balance}\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            return f"=IFERROR({flow_f}/{bal_f},0)"

        elif formula_type == 'coverage':
            earnings = kpi['earnings']
            interest = kpi['interest']
            if period == 'current':
                earn_f = f"SUMIF('Source P&L'!$A:$A,\"{earnings}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                int_f = f"SUMIF('Source P&L'!$A:$A,\"{interest}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:
                earn_f = f"SUMIF('Source P&L'!$A:$A,\"{earnings}\",'Source P&L'!$B:{get_column_letter(current_col)})"
                int_f = f"SUMIF('Source P&L'!$A:$A,\"{interest}\",'Source P&L'!$B:{get_column_letter(current_col)})"
            return f"=IFERROR({earn_f}/ABS({int_f}),99)"

        elif formula_type == 'growth':
            metric = kpi['metric']
            if len(self.months) >= 2:
                curr_f = f"SUMIF('Source P&L'!$A:$A,\"{metric}\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                prev_f = f"SUMIF('Source P&L'!$A:$A,\"{metric}\",'Source P&L'!{get_column_letter(current_col-1)}:{get_column_letter(current_col-1)})"
                return f"=IFERROR(({curr_f}-{prev_f})/{prev_f},0)"
            return "=0"

        elif formula_type == 'calculated':
            calc_type = kpi['calc_type']
            return self._build_calculated_kpi_formula(calc_type, period, current_col)

        return "=0"

    def _build_calculated_kpi_formula(self, calc_type: str, period: str, current_col: int) -> str:
        """Build formula for complex calculated KPIs"""
        if calc_type == 'ebitda':
            if period == 'current':
                ni = f"SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                int_e = f"SUMIF('Source P&L'!$A:$A,\"8000 Interest Expense\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                dep = f"SUMIF('Source P&L'!$A:$A,\"6090 Depreciation Expense\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:
                ni = f"SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!$B:{get_column_letter(current_col)})"
                int_e = f"SUMIF('Source P&L'!$A:$A,\"8000 Interest Expense\",'Source P&L'!$B:{get_column_letter(current_col)})"
                dep = f"SUMIF('Source P&L'!$A:$A,\"6090 Depreciation Expense\",'Source P&L'!$B:{get_column_letter(current_col)})"
            return f"={ni}+ABS({int_e})+ABS({dep})"

        elif calc_type == 'ccc':
            ar = f"SUMIF('Source BS'!$A:$A,\"Total for Accounts Receivable\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            ap = f"SUMIF('Source BS'!$A:$A,\"Total for Credit Cards\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            rev = f"SUMIF('Source P&L'!$A:$A,\"Total for Income\",'Source P&L'!$B:{get_column_letter(current_col)})"
            cogs = f"SUMIF('Source P&L'!$A:$A,\"Total for Cost of Sales\",'Source P&L'!$B:{get_column_letter(current_col)})"
            months = len(self.months) if self.months else 1
            ar_days = f"({ar}/(({rev})/{months})*30)"
            ap_days = f"({ap}/(({cogs})/{months})*30)"
            return f"=IFERROR({ar_days}-{ap_days},0)"

        elif calc_type == 'ocf':
            if period == 'current':
                ni = f"SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                dep = f"SUMIF('Source P&L'!$A:$A,\"6090 Depreciation Expense\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:
                ni = f"SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!$B:{get_column_letter(current_col)})"
                dep = f"SUMIF('Source P&L'!$A:$A,\"6090 Depreciation Expense\",'Source P&L'!$B:{get_column_letter(current_col)})"
            return f"={ni}+ABS({dep})"

        elif calc_type == 'fcf':
            if period == 'current':
                ni = f"SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                dep = f"SUMIF('Source P&L'!$A:$A,\"6090 Depreciation Expense\",'Source P&L'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            else:
                ni = f"SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!$B:{get_column_letter(current_col)})"
                dep = f"SUMIF('Source P&L'!$A:$A,\"6090 Depreciation Expense\",'Source P&L'!$B:{get_column_letter(current_col)})"
            return f"={ni}+ABS({dep})"

        elif calc_type == 'burn_rate':
            if len(self.months) >= 2:
                curr = f"SUMIF('Source BS'!$A:$A,\"Total for Bank Accounts\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                prev = f"SUMIF('Source BS'!$A:$A,\"Total for Bank Accounts\",'Source BS'!{get_column_letter(current_col-1)}:{get_column_letter(current_col-1)})"
                return f"={prev}-{curr}"
            return "=0"

        elif calc_type == 'runway':
            cash = f"SUMIF('Source BS'!$A:$A,\"Total for Bank Accounts\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
            if len(self.months) >= 2:
                curr = f"SUMIF('Source BS'!$A:$A,\"Total for Bank Accounts\",'Source BS'!{get_column_letter(current_col)}:{get_column_letter(current_col)})"
                prev = f"SUMIF('Source BS'!$A:$A,\"Total for Bank Accounts\",'Source BS'!{get_column_letter(current_col-1)}:{get_column_letter(current_col-1)})"
                burn = f"({prev}-{curr})"
                return f"=IFERROR({cash}/{burn},999)"
            return "=999"

        return "=0"

    def _build_status_formula(self, row: int, value_col: int, target_col: int, higher_is_better: bool) -> str:
        """Build formula for status indicator"""
        val = f"{get_column_letter(value_col)}{row}"
        tgt = f"{get_column_letter(target_col)}{row}"
        if higher_is_better:
            return f'=IF({val}>={tgt},"G",IF({val}>={tgt}*0.8,"Y","R"))'
        else:
            return f'=IF({val}<={tgt},"G",IF({val}<={tgt}*1.2,"Y","R"))'

    def _add_status_formatting(self, ws, start_row: int, end_row: int, col: int):
        """Add conditional formatting for status column"""
        col_letter = get_column_letter(col)
        range_str = f"{col_letter}{start_row}:{col_letter}{end_row}"

        green_rule = FormulaRule(formula=[f'{col_letter}{start_row}="G"'], fill=self.green_fill,
                                  font=Font(color='FFFFFF', bold=True))
        ws.conditional_formatting.add(range_str, green_rule)

        yellow_rule = FormulaRule(formula=[f'{col_letter}{start_row}="Y"'], fill=self.yellow_fill,
                                   font=Font(color='000000', bold=True))
        ws.conditional_formatting.add(range_str, yellow_rule)

        red_rule = FormulaRule(formula=[f'{col_letter}{start_row}="R"'], fill=self.red_fill,
                                font=Font(color='FFFFFF', bold=True))
        ws.conditional_formatting.add(range_str, red_rule)

    def _add_trend_formatting(self, ws, start_row: int, end_row: int, col: int):
        """Add conditional formatting for trend column"""
        col_letter = get_column_letter(col)
        range_str = f"{col_letter}{start_row}:{col_letter}{end_row}"

        up_rule = FormulaRule(formula=[f'{col_letter}{start_row}="UP"'],
                               font=Font(color='27AE60', bold=True))
        ws.conditional_formatting.add(range_str, up_rule)

        down_rule = FormulaRule(formula=[f'{col_letter}{start_row}="DOWN"'],
                                 font=Font(color='E74C3C', bold=True))
        ws.conditional_formatting.add(range_str, down_rule)

    def _create_trend_data_section(self, ws, row_start: int) -> int:
        """Create trend data section with charts"""
        ws.cell(row_start, 2).value = "TREND ANALYSIS"
        ws.cell(row_start, 2).font = self.section_font
        ws.cell(row_start, 2).fill = self.section_fill
        for col in range(2, 14):
            ws.cell(row_start, col).fill = self.section_fill

        row_start += 2

        # Data table headers
        ws.cell(row_start, 2).value = "Month"
        ws.cell(row_start, 3).value = "Revenue"
        ws.cell(row_start, 4).value = "Net Income"
        ws.cell(row_start, 5).value = "Gross Margin"
        ws.cell(row_start, 6).value = "Cash"

        for col in range(2, 7):
            ws.cell(row_start, col).font = self.header_font
            ws.cell(row_start, col).fill = self.kpi_header_fill
            ws.cell(row_start, col).alignment = self.center_align

        # Monthly data
        data_start = row_start + 1
        display_months = self.months[-12:] if len(self.months) > 12 else self.months

        for i, (m, y) in enumerate(display_months):
            row = data_start + i
            col_idx = 2 + self.months.index((m, y))

            ws.cell(row, 2).value = f"{self.MONTHS[m-1][:3]} {y}"
            ws.cell(row, 2).font = self.normal_font

            ws.cell(row, 3).value = f"=SUMIF('Source P&L'!$A:$A,\"Total for Income\",'Source P&L'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 3).number_format = '#,##0'

            ws.cell(row, 4).value = f"=SUMIF('Source P&L'!$A:$A,\"Net Income\",'Source P&L'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 4).number_format = '#,##0'

            gp = f"SUMIF('Source P&L'!$A:$A,\"Gross Profit\",'Source P&L'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            rev = f"SUMIF('Source P&L'!$A:$A,\"Total for Income\",'Source P&L'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 5).value = f"=IFERROR({gp}/{rev},0)"
            ws.cell(row, 5).number_format = '0.0%'

            ws.cell(row, 6).value = f"=SUMIF('Source BS'!$A:$A,\"Total for Bank Accounts\",'Source BS'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 6).number_format = '#,##0'

        chart_data_end = data_start + len(display_months) - 1

        # Create chart
        try:
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"
            chart.title = "Revenue & Net Income Trend"
            chart.style = 10

            revenue_data = Reference(ws, min_col=3, min_row=row_start, max_row=chart_data_end)
            chart.add_data(revenue_data, titles_from_data=True)

            ni_data = Reference(ws, min_col=4, min_row=row_start, max_row=chart_data_end)
            chart.add_data(ni_data, titles_from_data=True)

            cats = Reference(ws, min_col=2, min_row=data_start, max_row=chart_data_end)
            chart.set_categories(cats)

            chart.width = 15
            chart.height = 8

            ws.add_chart(chart, f"H{row_start}")
        except Exception:
            ws.cell(row_start + 2, 8).value = "[Chart: Revenue & Net Income Trend]"
            ws.cell(row_start + 2, 8).font = self.small_font

        return chart_data_end + 3
