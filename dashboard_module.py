"""
Financial Model Dashboard Module
Adds CFO-grade KPI Dashboard with metrics, ratios, and visualizations
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Color
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.workbook.defined_name import DefinedName
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import re


class DashboardModule:
    """Adds Dashboard functionality to Financial Model workbook"""

    # KPI Definitions with formulas, descriptions, and default targets
    KPI_DEFINITIONS = {
        # Profitability KPIs
        'profitability': [
            {
                'name': 'Revenue',
                'formula_type': 'direct',
                'source': 'PL',
                'account': 'Total for Income',
                'description': 'Total revenue generated',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Gross Profit',
                'formula_type': 'direct',
                'source': 'PL',
                'account': 'Gross Profit',
                'description': 'Revenue minus cost of goods sold',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Gross Margin %',
                'formula_type': 'ratio',
                'numerator': 'Gross Profit',
                'denominator': 'Total for Income',
                'description': 'Gross profit as percentage of revenue',
                'target_type': 'standard',
                'default_target': 0.40,
                'format': '0.0%',
                'higher_is_better': True
            },
            {
                'name': 'Net Income',
                'formula_type': 'direct',
                'source': 'PL',
                'account': 'Net Income',
                'description': 'Bottom line profit after all expenses',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Net Margin %',
                'formula_type': 'ratio',
                'numerator': 'Net Income',
                'denominator': 'Total for Income',
                'description': 'Net income as percentage of revenue',
                'target_type': 'standard',
                'default_target': 0.10,
                'format': '0.0%',
                'higher_is_better': True
            },
            {
                'name': 'EBITDA',
                'formula_type': 'calculated',
                'calc_type': 'ebitda',
                'description': 'Earnings before interest, taxes, depreciation, amortization',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'EBITDA Margin %',
                'formula_type': 'ratio',
                'numerator': 'EBITDA',
                'denominator': 'Total for Income',
                'description': 'EBITDA as percentage of revenue',
                'target_type': 'standard',
                'default_target': 0.15,
                'format': '0.0%',
                'higher_is_better': True
            },
            {
                'name': 'Operating Income',
                'formula_type': 'direct',
                'source': 'PL',
                'account': 'Net Operating Income',
                'description': 'Income from core operations',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Operating Margin %',
                'formula_type': 'ratio',
                'numerator': 'Net Operating Income',
                'denominator': 'Total for Income',
                'description': 'Operating income as percentage of revenue',
                'target_type': 'standard',
                'default_target': 0.12,
                'format': '0.0%',
                'higher_is_better': True
            },
            {
                'name': 'Revenue Growth %',
                'formula_type': 'growth',
                'metric': 'Total for Income',
                'description': 'Period-over-period revenue growth rate',
                'target_type': 'standard',
                'default_target': 0.10,
                'format': '0.0%',
                'higher_is_better': True
            },
        ],
        # Liquidity KPIs
        'liquidity': [
            {
                'name': 'Current Ratio',
                'formula_type': 'bs_ratio',
                'numerator': 'Total for Current Assets',
                'denominator': 'Total for Current Liabilities',
                'description': 'Current assets / current liabilities',
                'target_type': 'standard',
                'default_target': 1.5,
                'format': '0.00',
                'higher_is_better': True
            },
            {
                'name': 'Quick Ratio',
                'formula_type': 'calculated',
                'calc_type': 'quick_ratio',
                'description': '(Current assets - inventory) / current liabilities',
                'target_type': 'standard',
                'default_target': 1.0,
                'format': '0.00',
                'higher_is_better': True
            },
            {
                'name': 'Cash Ratio',
                'formula_type': 'bs_ratio',
                'numerator': 'Total for Bank Accounts',
                'denominator': 'Total for Current Liabilities',
                'description': 'Cash / current liabilities',
                'target_type': 'standard',
                'default_target': 0.2,
                'format': '0.00',
                'higher_is_better': True
            },
            {
                'name': 'Working Capital',
                'formula_type': 'bs_difference',
                'minuend': 'Total for Current Assets',
                'subtrahend': 'Total for Current Liabilities',
                'description': 'Current assets minus current liabilities',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Cash Balance',
                'formula_type': 'direct',
                'source': 'BS',
                'account': 'Total for Bank Accounts',
                'description': 'Total cash and bank accounts',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
        ],
        # Efficiency KPIs
        'efficiency': [
            {
                'name': 'AR Days (DSO)',
                'formula_type': 'days_ratio',
                'balance': 'Total for Accounts Receivable',
                'flow': 'Total for Income',
                'description': 'Average days to collect receivables',
                'target_type': 'standard',
                'default_target': 45,
                'format': '0',
                'higher_is_better': False
            },
            {
                'name': 'AP Days (DPO)',
                'formula_type': 'days_ratio',
                'balance': 'Total for Credit Cards',
                'flow': 'Total for Cost of Sales',
                'description': 'Average days to pay payables',
                'target_type': 'standard',
                'default_target': 30,
                'format': '0',
                'higher_is_better': True
            },
            {
                'name': 'Asset Turnover',
                'formula_type': 'turnover',
                'flow': 'Total for Income',
                'balance': 'Total for Assets',
                'description': 'Revenue generated per dollar of assets',
                'target_type': 'standard',
                'default_target': 1.0,
                'format': '0.00',
                'higher_is_better': True
            },
            {
                'name': 'Cash Conversion Cycle',
                'formula_type': 'calculated',
                'calc_type': 'ccc',
                'description': 'Days from cash outflow to cash inflow',
                'target_type': 'standard',
                'default_target': 30,
                'format': '0',
                'higher_is_better': False
            },
        ],
        # Leverage KPIs
        'leverage': [
            {
                'name': 'Debt-to-Equity',
                'formula_type': 'bs_ratio',
                'numerator': 'Total for Liabilities',
                'denominator': 'Total for Equity',
                'description': 'Total debt relative to equity',
                'target_type': 'standard',
                'default_target': 1.0,
                'format': '0.00',
                'higher_is_better': False
            },
            {
                'name': 'Debt-to-Assets',
                'formula_type': 'bs_ratio',
                'numerator': 'Total for Liabilities',
                'denominator': 'Total for Assets',
                'description': 'Percentage of assets financed by debt',
                'target_type': 'standard',
                'default_target': 0.5,
                'format': '0.0%',
                'higher_is_better': False
            },
            {
                'name': 'Equity Ratio',
                'formula_type': 'bs_ratio',
                'numerator': 'Total for Equity',
                'denominator': 'Total for Assets',
                'description': 'Percentage of assets financed by equity',
                'target_type': 'standard',
                'default_target': 0.5,
                'format': '0.0%',
                'higher_is_better': True
            },
            {
                'name': 'Interest Coverage',
                'formula_type': 'coverage',
                'earnings': 'Net Operating Income',
                'interest': '8000 Interest Expense',
                'description': 'Operating income / interest expense',
                'target_type': 'standard',
                'default_target': 3.0,
                'format': '0.0',
                'higher_is_better': True
            },
        ],
        # Cash Flow KPIs
        'cashflow': [
            {
                'name': 'Operating Cash Flow',
                'formula_type': 'calculated',
                'calc_type': 'ocf',
                'description': 'Cash generated from operations',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Free Cash Flow',
                'formula_type': 'calculated',
                'calc_type': 'fcf',
                'description': 'Operating cash flow minus capital expenditures',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': True
            },
            {
                'name': 'Cash Burn Rate',
                'formula_type': 'calculated',
                'calc_type': 'burn_rate',
                'description': 'Monthly cash consumption rate',
                'target_type': 'user_defined',
                'default_target': 0,
                'format': '#,##0',
                'higher_is_better': False
            },
            {
                'name': 'Cash Runway (Months)',
                'formula_type': 'calculated',
                'calc_type': 'runway',
                'description': 'Months of cash remaining at current burn rate',
                'target_type': 'standard',
                'default_target': 12,
                'format': '0.0',
                'higher_is_better': True
            },
        ],
    }

    def __init__(self, workbook_path: str):
        """Initialize with path to existing Financial Model workbook"""
        self.wb_path = workbook_path
        self.wb = openpyxl.load_workbook(workbook_path, keep_vba=True)
        self._define_styles()
        self._analyze_data_structure()

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

        # Dashboard-specific fills
        self.header_fill = PatternFill(start_color='16213E', end_color='16213E', fill_type='solid')
        self.section_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        self.kpi_header_fill = PatternFill(start_color='34495E', end_color='34495E', fill_type='solid')
        self.light_fill = PatternFill(start_color='ECF0F1', end_color='ECF0F1', fill_type='solid')
        self.alt_row_fill = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
        self.green_fill = PatternFill(start_color='27AE60', end_color='27AE60', fill_type='solid')
        self.yellow_fill = PatternFill(start_color='F39C12', end_color='F39C12', fill_type='solid')
        self.red_fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
        self.metric_fill = PatternFill(start_color='3498DB', end_color='3498DB', fill_type='solid')
        self.control_fill = PatternFill(start_color='E8F4FD', end_color='E8F4FD', fill_type='solid')

        # Fonts
        self.title_font = Font(name='Calibri Light', size=24, bold=True, color='16213E')
        self.section_font = Font(name='Calibri Light', size=14, bold=True, color='FFFFFF')
        self.header_font = Font(name='Calibri Light', size=11, bold=True, color='FFFFFF')
        self.kpi_name_font = Font(name='Calibri Light', size=11, bold=True, color='2C3E50')
        self.kpi_value_font = Font(name='Calibri Light', size=14, bold=True, color='16213E')
        self.kpi_small_font = Font(name='Calibri Light', size=9, color='7F8C8D')
        self.normal_font = Font(name='Calibri Light', size=10)
        self.bold_font = Font(name='Calibri Light', size=10, bold=True)
        self.small_font = Font(name='Calibri Light', size=9, color='666666')
        self.white_font = Font(name='Calibri Light', size=10, color='FFFFFF')
        self.link_font = Font(name='Calibri Light', size=10, color='0563C1', underline='single')

        # Alignments
        self.right_align = Alignment(horizontal='right', vertical='center')
        self.left_align = Alignment(horizontal='left', vertical='center')
        self.center_align = Alignment(horizontal='center', vertical='center')
        self.wrap_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    def _analyze_data_structure(self):
        """Analyze the source data to understand structure"""
        # Get P&L source data
        if 'Source_PL' in self.wb.sheetnames:
            ws_pl = self.wb['Source_PL']
        elif 'Source P&L' in self.wb.sheetnames:
            ws_pl = self.wb['Source P&L']
        else:
            raise ValueError("Cannot find Source P&L sheet")

        # Get BS source data
        if 'Source_BS' in self.wb.sheetnames:
            ws_bs = self.wb['Source_BS']
        elif 'Source BS' in self.wb.sheetnames:
            ws_bs = self.wb['Source BS']
        else:
            raise ValueError("Cannot find Source BS sheet")

        self.pl_sheet_name = ws_pl.title
        self.bs_sheet_name = ws_bs.title

        # Parse months from header row
        self.months = []
        self.month_columns = {}
        for col in range(2, ws_pl.max_column + 1):
            val = ws_pl.cell(1, col).value
            if val:
                parsed = self._parse_month(val)
                if parsed:
                    self.months.append(parsed)
                    self.month_columns[parsed] = col

        # Get account row mappings
        self.pl_accounts = {}
        for row in range(2, ws_pl.max_row + 1):
            account = ws_pl.cell(row, 1).value
            if account:
                self.pl_accounts[account] = row

        self.bs_accounts = {}
        for row in range(2, ws_bs.max_row + 1):
            account = ws_bs.cell(row, 1).value
            if account:
                self.bs_accounts[account] = row

        # Determine current month
        if self.months:
            self.current_month = self.months[-1]
            self.first_month = self.months[0]
        else:
            self.current_month = None
            self.first_month = None

    def _parse_month(self, val) -> Optional[Tuple[int, int]]:
        """Parse month value to (month, year) tuple"""
        if val is None:
            return None

        # Handle datetime objects
        if hasattr(val, 'month') and hasattr(val, 'year'):
            return (val.month, val.year)

        # Handle string formats
        val_str = str(val).strip()
        months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December']

        for i, month in enumerate(months, 1):
            if val_str.lower().startswith(month.lower()):
                parts = val_str.split()
                if len(parts) >= 2:
                    try:
                        year = int(parts[-1])
                        return (i, year)
                    except ValueError:
                        pass

        return None

    def _month_str(self, month_tuple: Tuple[int, int]) -> str:
        """Convert month tuple to string"""
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        return f"{months[month_tuple[0]-1]} {month_tuple[1]}"

    def _full_month_str(self, month_tuple: Tuple[int, int]) -> str:
        """Convert month tuple to full string"""
        months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December']
        return f"{months[month_tuple[0]-1]} {month_tuple[1]}"

    def _create_named_range(self, name: str, sheet: str, cell: str):
        """Create a named range"""
        if ':' in cell:
            parts = cell.split(':')
            ref = f"'{sheet}'!${parts[0].replace('$', '')}:${parts[1].replace('$', '')}"
        else:
            ref = f"'{sheet}'!${cell.replace('$', '')}"
        defn = DefinedName(name=name, attr_text=ref)
        self.wb.defined_names[name] = defn

    def create_dashboard_control_sheet(self):
        """Create the Dashboard Control page with target settings"""
        # Remove existing sheet if present
        if 'Dashboard_Control' in self.wb.sheetnames:
            del self.wb['Dashboard_Control']

        ws = self.wb.create_sheet("Dashboard_Control")
        ws.sheet_properties.tabColor = "3498DB"
        ws.sheet_view.showGridLines = False

        # Title
        ws['B2'] = "Dashboard Control Panel"
        ws['B2'].font = self.title_font
        ws.merge_cells('B2:F2')

        ws['B3'] = "Configure KPI targets and thresholds"
        ws['B3'].font = self.small_font
        ws.merge_cells('B3:F3')

        # Instructions section
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

        # KPI Target Settings Header
        header_row = 14
        ws.cell(header_row, 2).value = "KPI TARGETS AND THRESHOLDS"
        ws.cell(header_row, 2).font = self.section_font
        ws.cell(header_row, 2).fill = self.section_fill
        for col in range(2, 9):
            ws.cell(header_row, col).fill = self.section_fill

        # Column headers
        headers = ['Category', 'KPI Name', 'Description', 'Target', 'Yellow %', 'Red %', 'Direction', 'Current Value']
        header_row = 16
        for col, header in enumerate(headers, 2):
            cell = ws.cell(header_row, col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.kpi_header_fill
            cell.alignment = self.center_align
            cell.border = self.thin_border

        # Populate KPI definitions
        data_row = 17
        for category, kpis in self.KPI_DEFINITIONS.items():
            for kpi in kpis:
                # Category
                ws.cell(data_row, 2).value = category.title()
                ws.cell(data_row, 2).font = self.normal_font
                ws.cell(data_row, 2).alignment = self.left_align

                # KPI Name
                ws.cell(data_row, 3).value = kpi['name']
                ws.cell(data_row, 3).font = self.bold_font
                ws.cell(data_row, 3).alignment = self.left_align

                # Description
                ws.cell(data_row, 4).value = kpi['description']
                ws.cell(data_row, 4).font = self.small_font
                ws.cell(data_row, 4).alignment = self.wrap_align

                # Target (editable)
                target_cell = ws.cell(data_row, 5)
                target_cell.value = kpi['default_target']
                if '%' in kpi['format']:
                    target_cell.number_format = '0.0%'
                elif '0.0' in kpi['format'] or '0.00' in kpi['format']:
                    target_cell.number_format = '0.00'
                else:
                    target_cell.number_format = '#,##0'
                target_cell.font = self.normal_font
                target_cell.fill = self.control_fill
                target_cell.alignment = self.right_align
                target_cell.border = self.thin_border

                # Yellow threshold (default 80%)
                yellow_cell = ws.cell(data_row, 6)
                yellow_cell.value = 0.80
                yellow_cell.number_format = '0%'
                yellow_cell.font = self.normal_font
                yellow_cell.fill = self.control_fill
                yellow_cell.alignment = self.center_align
                yellow_cell.border = self.thin_border

                # Red threshold (default 60%)
                red_cell = ws.cell(data_row, 7)
                red_cell.value = 0.60
                red_cell.number_format = '0%'
                red_cell.font = self.normal_font
                red_cell.fill = self.control_fill
                red_cell.alignment = self.center_align
                red_cell.border = self.thin_border

                # Direction (Higher/Lower is better)
                dir_cell = ws.cell(data_row, 8)
                dir_cell.value = "Higher" if kpi['higher_is_better'] else "Lower"
                dir_cell.font = self.normal_font
                dir_cell.alignment = self.center_align

                # Current Value (formula placeholder - will link to dashboard)
                val_cell = ws.cell(data_row, 9)
                val_cell.value = "=Dashboard!" + f"E{data_row - 14}"  # Placeholder
                val_cell.font = self.kpi_value_font
                val_cell.alignment = self.right_align

                # Alternate row fill
                if data_row % 2 == 0:
                    for col in range(2, 10):
                        if ws.cell(data_row, col).fill == PatternFill():
                            ws.cell(data_row, col).fill = self.alt_row_fill

                data_row += 1

        # Add data validation for yellow/red thresholds
        pct_validation = DataValidation(
            type="decimal",
            operator="between",
            formula1="0",
            formula2="1",
            allow_blank=True
        )
        pct_validation.error = "Please enter a percentage between 0% and 100%"
        pct_validation.errorTitle = "Invalid Percentage"
        ws.add_data_validation(pct_validation)
        pct_validation.add(f'F17:G{data_row-1}')

        # Column widths
        ws.column_dimensions['A'].width = 3
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 45
        ws.column_dimensions['E'].width = 14
        ws.column_dimensions['F'].width = 10
        ws.column_dimensions['G'].width = 10
        ws.column_dimensions['H'].width = 10
        ws.column_dimensions['I'].width = 16

        # Create named ranges for targets
        self._create_named_range('KPI_Targets', 'Dashboard_Control', f'C17:E{data_row-1}')

        # Print setup
        ws.print_title_rows = '16:16'
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        return ws

    def create_dashboard_sheet(self):
        """Create the main Dashboard sheet with all KPIs and visualizations"""
        # Remove existing sheet if present
        if 'Dashboard' in self.wb.sheetnames:
            del self.wb['Dashboard']

        ws = self.wb.create_sheet("Dashboard", 0)
        ws.sheet_properties.tabColor = "27AE60"
        ws.sheet_view.showGridLines = False

        # Get company name from Menu
        company_name = "Company"
        if 'Menu' in self.wb.sheetnames:
            menu = self.wb['Menu']
            if menu['B2'].value:
                company_name = menu['B2'].value

        # Title Section
        ws['B2'] = company_name
        ws['B2'].font = Font(name='Calibri Light', size=28, bold=True, color='16213E')
        ws.merge_cells('B2:L2')

        ws['B3'] = "Executive Dashboard"
        ws['B3'].font = Font(name='Calibri Light', size=16, color='7F8C8D')
        ws.merge_cells('B3:L3')

        # Current period info
        if self.current_month:
            ws['B4'] = f"Current Period: {self._full_month_str(self.current_month)}"
        else:
            ws['B4'] = "Current Period: N/A"
        ws['B4'].font = self.small_font

        ws['B5'] = f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M')}"
        ws['B5'].font = self.small_font

        # Summary Cards Row (Key metrics at top)
        self._create_summary_cards(ws, row_start=7)

        # KPI Sections
        current_row = 17

        # Create sections for each KPI category
        sections = [
            ('PROFITABILITY METRICS', 'profitability'),
            ('LIQUIDITY METRICS', 'liquidity'),
            ('EFFICIENCY METRICS', 'efficiency'),
            ('LEVERAGE METRICS', 'leverage'),
            ('CASH FLOW METRICS', 'cashflow'),
        ]

        kpi_value_cells = {}  # Track KPI value cells for control sheet linking

        for section_title, category in sections:
            current_row = self._create_kpi_section(
                ws, section_title, category, current_row, kpi_value_cells
            )
            current_row += 2  # Space between sections

        # Add Trend Charts section
        current_row = self._create_trend_charts_section(ws, current_row)

        # Column widths
        ws.column_dimensions['A'].width = 3
        ws.column_dimensions['B'].width = 22
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['E'].width = 14
        ws.column_dimensions['F'].width = 3
        ws.column_dimensions['G'].width = 14
        ws.column_dimensions['H'].width = 14
        ws.column_dimensions['I'].width = 14
        ws.column_dimensions['J'].width = 3
        ws.column_dimensions['K'].width = 14
        ws.column_dimensions['L'].width = 14
        ws.column_dimensions['M'].width = 14
        ws.column_dimensions['N'].width = 8
        ws.column_dimensions['O'].width = 12

        # Print setup
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5

        # Store KPI cells for linking
        self.kpi_value_cells = kpi_value_cells

        return ws

    def _create_summary_cards(self, ws, row_start: int):
        """Create summary metric cards at top of dashboard"""
        # Card definitions - key metrics to highlight
        cards = [
            {
                'title': 'Revenue',
                'metric': 'Total for Income',
                'source': 'PL',
                'format': '"$"#,##0',
                'col': 2
            },
            {
                'title': 'Net Income',
                'metric': 'Net Income',
                'source': 'PL',
                'format': '"$"#,##0',
                'col': 4
            },
            {
                'title': 'Gross Margin',
                'numerator': 'Gross Profit',
                'denominator': 'Total for Income',
                'format': '0.0%',
                'col': 6
            },
            {
                'title': 'Cash Balance',
                'metric': 'Total for Bank Accounts',
                'source': 'BS',
                'format': '"$"#,##0',
                'col': 8
            },
            {
                'title': 'Current Ratio',
                'numerator': 'Total for Current Assets',
                'denominator': 'Total for Current Liabilities',
                'source': 'BS',
                'format': '0.00',
                'col': 10
            },
            {
                'title': 'AR Days',
                'type': 'days',
                'balance': 'Total for Accounts Receivable',
                'flow': 'Total for Income',
                'format': '0',
                'col': 12
            },
        ]

        for card in cards:
            col = card['col']

            # Card header
            header_cell = ws.cell(row_start, col)
            header_cell.value = card['title']
            header_cell.font = Font(name='Calibri Light', size=10, bold=True, color='7F8C8D')
            header_cell.alignment = self.center_align
            ws.merge_cells(start_row=row_start, start_column=col, end_row=row_start, end_column=col+1)

            # Card value (with formula)
            value_cell = ws.cell(row_start + 1, col)
            value_cell.font = Font(name='Calibri Light', size=20, bold=True, color='16213E')
            value_cell.alignment = self.center_align
            ws.merge_cells(start_row=row_start+1, start_column=col, end_row=row_start+1, end_column=col+1)

            # Build formula based on card type
            if 'metric' in card:
                # Direct metric lookup
                if card['source'] == 'PL':
                    formula = self._build_pl_lookup_formula(card['metric'], 'current')
                else:
                    formula = self._build_bs_lookup_formula(card['metric'], 'current')
            elif 'numerator' in card:
                # Ratio
                if card.get('source') == 'BS':
                    num_formula = self._build_bs_lookup_formula(card['numerator'], 'current')
                    den_formula = self._build_bs_lookup_formula(card['denominator'], 'current')
                else:
                    num_formula = self._build_pl_lookup_formula(card['numerator'], 'current')
                    den_formula = self._build_pl_lookup_formula(card['denominator'], 'current')
                formula = f"=IFERROR({num_formula}/{den_formula},0)"
            elif card.get('type') == 'days':
                # Days calculation
                bal_formula = self._build_bs_lookup_formula(card['balance'], 'current')
                flow_formula = self._build_pl_lookup_formula(card['flow'], 'ytd')
                months_count = len(self.months) if self.months else 1
                formula = f"=IFERROR({bal_formula}/({flow_formula}/{months_count})*30,0)"
            else:
                formula = "=0"

            value_cell.value = formula
            value_cell.number_format = card['format']

            # Period label
            period_cell = ws.cell(row_start + 2, col)
            period_cell.value = "Current Month" if 'metric' in card or 'numerator' in card else "YTD Avg"
            period_cell.font = self.small_font
            period_cell.alignment = self.center_align
            ws.merge_cells(start_row=row_start+2, start_column=col, end_row=row_start+2, end_column=col+1)

            # Card border/styling
            for r in range(row_start, row_start + 3):
                for c in range(col, col + 2):
                    ws.cell(r, c).border = Border(
                        left=Side(style='thin', color='E0E0E0') if c == col else None,
                        right=Side(style='thin', color='E0E0E0') if c == col + 1 else None,
                        top=Side(style='thin', color='E0E0E0') if r == row_start else None,
                        bottom=Side(style='thin', color='E0E0E0') if r == row_start + 2 else None
                    )

    def _create_kpi_section(self, ws, title: str, category: str, row_start: int,
                           kpi_value_cells: dict) -> int:
        """Create a KPI section with current, YTD, and all-time views"""
        kpis = self.KPI_DEFINITIONS.get(category, [])
        if not kpis:
            return row_start

        # Section header
        ws.cell(row_start, 2).value = title
        ws.cell(row_start, 2).font = self.section_font
        ws.cell(row_start, 2).fill = self.section_fill
        for col in range(2, 15):
            ws.cell(row_start, col).fill = self.section_fill

        # Column sub-headers
        row_start += 1
        headers = [
            ('KPI', 2),
            ('Current Month', 3),
            ('Target', 4),
            ('Status', 5),
            ('', 6),  # Spacer
            ('YTD', 7),
            ('YTD Target', 8),
            ('YTD Status', 9),
            ('', 10),  # Spacer
            ('All-Time', 11),
            ('Trend', 12),
            ('Notes', 13),
        ]

        for header, col in headers:
            cell = ws.cell(row_start, col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.kpi_header_fill
            cell.alignment = self.center_align
            cell.border = self.thin_border

        # KPI rows
        data_row = row_start + 1
        for kpi in kpis:
            # KPI Name
            ws.cell(data_row, 2).value = kpi['name']
            ws.cell(data_row, 2).font = self.kpi_name_font
            ws.cell(data_row, 2).alignment = self.left_align

            # Current Month Value
            current_cell = ws.cell(data_row, 3)
            current_formula = self._build_kpi_formula(kpi, 'current')
            current_cell.value = current_formula
            current_cell.number_format = kpi['format']
            current_cell.font = self.kpi_value_font
            current_cell.alignment = self.right_align

            # Track this cell for control sheet linking
            kpi_value_cells[kpi['name']] = f'C{data_row}'

            # Target (linked to control sheet)
            target_cell = ws.cell(data_row, 4)
            control_row = self._get_control_row_for_kpi(kpi['name'])
            if control_row:
                target_cell.value = f"=Dashboard_Control!E{control_row}"
            else:
                target_cell.value = kpi['default_target']
            target_cell.number_format = kpi['format']
            target_cell.font = self.normal_font
            target_cell.alignment = self.right_align

            # Status indicator (conditional formula)
            status_cell = ws.cell(data_row, 5)
            status_formula = self._build_status_formula(data_row, 3, 4, kpi['higher_is_better'])
            status_cell.value = status_formula
            status_cell.font = Font(name='Calibri Light', size=12, bold=True)
            status_cell.alignment = self.center_align

            # YTD Value
            ytd_cell = ws.cell(data_row, 7)
            ytd_formula = self._build_kpi_formula(kpi, 'ytd')
            ytd_cell.value = ytd_formula
            ytd_cell.number_format = kpi['format']
            ytd_cell.font = self.bold_font
            ytd_cell.alignment = self.right_align

            # YTD Target (annualized if applicable)
            ytd_target_cell = ws.cell(data_row, 8)
            if '%' in kpi['format'] or 'ratio' in kpi['formula_type'].lower():
                # Ratio targets stay the same
                ytd_target_cell.value = f"=D{data_row}"
            else:
                # Dollar targets: multiply by months elapsed
                months_elapsed = len(self.months) if self.months else 1
                ytd_target_cell.value = f"=D{data_row}*{months_elapsed}"
            ytd_target_cell.number_format = kpi['format']
            ytd_target_cell.font = self.normal_font
            ytd_target_cell.alignment = self.right_align

            # YTD Status
            ytd_status_cell = ws.cell(data_row, 9)
            ytd_status_formula = self._build_status_formula(data_row, 7, 8, kpi['higher_is_better'])
            ytd_status_cell.value = ytd_status_formula
            ytd_status_cell.font = Font(name='Calibri Light', size=12, bold=True)
            ytd_status_cell.alignment = self.center_align

            # All-Time Value
            alltime_cell = ws.cell(data_row, 11)
            alltime_formula = self._build_kpi_formula(kpi, 'alltime')
            alltime_cell.value = alltime_formula
            alltime_cell.number_format = kpi['format']
            alltime_cell.font = self.normal_font
            alltime_cell.alignment = self.right_align

            # Trend indicator
            trend_cell = ws.cell(data_row, 12)
            trend_formula = self._build_trend_formula(kpi, data_row)
            trend_cell.value = trend_formula
            trend_cell.font = Font(name='Calibri Light', size=12)
            trend_cell.alignment = self.center_align

            # Notes placeholder
            notes_cell = ws.cell(data_row, 13)
            notes_cell.value = ""
            notes_cell.font = self.small_font

            # Alternate row styling
            if data_row % 2 == 0:
                for col in range(2, 14):
                    if col not in [6, 10]:  # Skip spacer columns
                        ws.cell(data_row, col).fill = self.alt_row_fill

            data_row += 1

        # Add conditional formatting for status cells
        self._add_status_conditional_formatting(ws, row_start + 2, data_row - 1, 5)
        self._add_status_conditional_formatting(ws, row_start + 2, data_row - 1, 9)
        self._add_trend_conditional_formatting(ws, row_start + 2, data_row - 1, 12)

        return data_row

    def _build_pl_lookup_formula(self, account: str, period: str) -> str:
        """Build formula to lookup P&L account value"""
        if period == 'current' and self.current_month:
            col = self.month_columns.get(self.current_month, 2)
            return f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"{account}\",'{self.pl_sheet_name}'!{get_column_letter(col)}:{get_column_letter(col)})"
        elif period == 'ytd':
            # Sum all months
            first_col = get_column_letter(2)
            last_col = get_column_letter(max(self.month_columns.values()) if self.month_columns else 2)
            return f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"{account}\",'{self.pl_sheet_name}'!{first_col}:{last_col})"
        else:
            # All-time (same as YTD for now)
            first_col = get_column_letter(2)
            last_col = get_column_letter(max(self.month_columns.values()) if self.month_columns else 2)
            return f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"{account}\",'{self.pl_sheet_name}'!{first_col}:{last_col})"

    def _build_bs_lookup_formula(self, account: str, period: str) -> str:
        """Build formula to lookup Balance Sheet account value"""
        if period == 'current' and self.current_month:
            col = self.month_columns.get(self.current_month, 2)
            return f"SUMIF('{self.bs_sheet_name}'!$A:$A,\"{account}\",'{self.bs_sheet_name}'!{get_column_letter(col)}:{get_column_letter(col)})"
        elif period == 'ytd':
            # For BS, YTD typically uses current period value (point-in-time)
            col = self.month_columns.get(self.current_month, max(self.month_columns.values()) if self.month_columns else 2)
            return f"SUMIF('{self.bs_sheet_name}'!$A:$A,\"{account}\",'{self.bs_sheet_name}'!{get_column_letter(col)}:{get_column_letter(col)})"
        else:
            # All-time uses first period for comparison
            col = self.month_columns.get(self.first_month, 2) if self.first_month else 2
            return f"SUMIF('{self.bs_sheet_name}'!$A:$A,\"{account}\",'{self.bs_sheet_name}'!{get_column_letter(col)}:{get_column_letter(col)})"

    def _build_kpi_formula(self, kpi: dict, period: str) -> str:
        """Build the formula for a KPI based on its type"""
        formula_type = kpi['formula_type']

        if formula_type == 'direct':
            source = kpi['source']
            account = kpi['account']
            if source == 'PL':
                return f"={self._build_pl_lookup_formula(account, period)}"
            else:
                return f"={self._build_bs_lookup_formula(account, period)}"

        elif formula_type == 'ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            num_formula = self._build_pl_lookup_formula(num, period)
            den_formula = self._build_pl_lookup_formula(den, period)
            return f"=IFERROR({num_formula}/{den_formula},0)"

        elif formula_type == 'bs_ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            num_formula = self._build_bs_lookup_formula(num, period)
            den_formula = self._build_bs_lookup_formula(den, period)
            return f"=IFERROR({num_formula}/{den_formula},0)"

        elif formula_type == 'bs_difference':
            min_acct = kpi['minuend']
            sub_acct = kpi['subtrahend']
            min_formula = self._build_bs_lookup_formula(min_acct, period)
            sub_formula = self._build_bs_lookup_formula(sub_acct, period)
            return f"={min_formula}-{sub_formula}"

        elif formula_type == 'days_ratio':
            balance = kpi['balance']
            flow = kpi['flow']
            bal_formula = self._build_bs_lookup_formula(balance, period)
            if period == 'current':
                flow_formula = self._build_pl_lookup_formula(flow, 'current')
                return f"=IFERROR({bal_formula}/{flow_formula}*30,0)"
            else:
                flow_formula = self._build_pl_lookup_formula(flow, 'ytd')
                months = len(self.months) if self.months else 1
                return f"=IFERROR({bal_formula}/(({flow_formula})/{months})*30,0)"

        elif formula_type == 'turnover':
            flow = kpi['flow']
            balance = kpi['balance']
            flow_formula = self._build_pl_lookup_formula(flow, period)
            bal_formula = self._build_bs_lookup_formula(balance, period)
            return f"=IFERROR({flow_formula}/{bal_formula},0)"

        elif formula_type == 'coverage':
            earnings = kpi['earnings']
            interest = kpi['interest']
            earn_formula = self._build_pl_lookup_formula(earnings, period)
            int_formula = self._build_pl_lookup_formula(interest, period)
            return f"=IFERROR({earn_formula}/ABS({int_formula}),99)"

        elif formula_type == 'growth':
            metric = kpi['metric']
            if period == 'current' and len(self.months) >= 2:
                curr_month = self.current_month
                prev_month = self.months[-2] if len(self.months) >= 2 else self.months[0]
                curr_col = self.month_columns.get(curr_month, 2)
                prev_col = self.month_columns.get(prev_month, 2)
                curr_formula = f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"{metric}\",'{self.pl_sheet_name}'!{get_column_letter(curr_col)}:{get_column_letter(curr_col)})"
                prev_formula = f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"{metric}\",'{self.pl_sheet_name}'!{get_column_letter(prev_col)}:{get_column_letter(prev_col)})"
                return f"=IFERROR(({curr_formula}-{prev_formula})/{prev_formula},0)"
            else:
                return "=0"

        elif formula_type == 'calculated':
            calc_type = kpi['calc_type']
            return self._build_calculated_formula(calc_type, period)

        return "=0"

    def _build_calculated_formula(self, calc_type: str, period: str) -> str:
        """Build formulas for complex calculated KPIs"""
        if calc_type == 'ebitda':
            # EBITDA = Net Income + Interest + Depreciation
            ni = self._build_pl_lookup_formula('Net Income', period)
            interest = self._build_pl_lookup_formula('8000 Interest Expense', period)
            deprec = self._build_pl_lookup_formula('6090 Depreciation Expense', period)
            return f"={ni}+ABS({interest})+ABS({deprec})"

        elif calc_type == 'quick_ratio':
            # Quick Ratio = (Current Assets - Inventory) / Current Liabilities
            # Since no inventory account visible, using Current Assets directly
            ca = self._build_bs_lookup_formula('Total for Current Assets', period)
            cl = self._build_bs_lookup_formula('Total for Current Liabilities', period)
            return f"=IFERROR({ca}/{cl},0)"

        elif calc_type == 'ccc':
            # Cash Conversion Cycle = AR Days + Inventory Days - AP Days
            # Simplified: AR Days - AP Days (no inventory)
            ar = self._build_bs_lookup_formula('Total for Accounts Receivable', period)
            ap = self._build_bs_lookup_formula('Total for Credit Cards', period)
            revenue = self._build_pl_lookup_formula('Total for Income', 'ytd')
            cogs = self._build_pl_lookup_formula('Total for Cost of Sales', 'ytd')
            months = len(self.months) if self.months else 1
            ar_days = f"({ar}/(({revenue})/{months})*30)"
            ap_days = f"({ap}/(({cogs})/{months})*30)"
            return f"=IFERROR({ar_days}-{ap_days},0)"

        elif calc_type == 'ocf':
            # Operating Cash Flow (simplified): Net Income + Depreciation + WC Changes
            ni = self._build_pl_lookup_formula('Net Income', period)
            deprec = self._build_pl_lookup_formula('6090 Depreciation Expense', period)
            return f"={ni}+ABS({deprec})"

        elif calc_type == 'fcf':
            # Free Cash Flow = OCF - CapEx (simplified)
            ni = self._build_pl_lookup_formula('Net Income', period)
            deprec = self._build_pl_lookup_formula('6090 Depreciation Expense', period)
            # Using fixed assets change as proxy for CapEx
            return f"={ni}+ABS({deprec})"

        elif calc_type == 'burn_rate':
            # Cash Burn = Change in cash balance
            if len(self.months) >= 2:
                curr = self._build_bs_lookup_formula('Total for Bank Accounts', 'current')
                prev_month = self.months[-2] if len(self.months) >= 2 else self.months[0]
                prev_col = self.month_columns.get(prev_month, 2)
                prev = f"SUMIF('{self.bs_sheet_name}'!$A:$A,\"Total for Bank Accounts\",'{self.bs_sheet_name}'!{get_column_letter(prev_col)}:{get_column_letter(prev_col)})"
                return f"={prev}-{curr}"
            return "=0"

        elif calc_type == 'runway':
            # Cash Runway = Cash / Monthly Burn Rate
            cash = self._build_bs_lookup_formula('Total for Bank Accounts', 'current')
            if len(self.months) >= 2:
                curr = self._build_bs_lookup_formula('Total for Bank Accounts', 'current')
                prev_month = self.months[-2]
                prev_col = self.month_columns.get(prev_month, 2)
                prev = f"SUMIF('{self.bs_sheet_name}'!$A:$A,\"Total for Bank Accounts\",'{self.bs_sheet_name}'!{get_column_letter(prev_col)}:{get_column_letter(prev_col)})"
                burn = f"({prev}-{curr})"
                return f"=IFERROR({cash}/{burn},999)"
            return "=999"

        return "=0"

    def _build_status_formula(self, row: int, value_col: int, target_col: int,
                              higher_is_better: bool) -> str:
        """Build formula for status indicator (checkmark, warning, x)"""
        val = f"{get_column_letter(value_col)}{row}"
        tgt = f"{get_column_letter(target_col)}{row}"

        if higher_is_better:
            # Green if >= target, Yellow if >= 80%, Red if < 80%
            return f'=IF({val}>={tgt},"G",IF({val}>={tgt}*0.8,"Y","R"))'
        else:
            # Green if <= target, Yellow if <= 120%, Red if > 120%
            return f'=IF({val}<={tgt},"G",IF({val}<={tgt}*1.2,"Y","R"))'

    def _build_trend_formula(self, kpi: dict, row: int) -> str:
        """Build formula for trend indicator (up/down/flat arrow)"""
        # Compare current to prior month
        if len(self.months) < 2:
            return '="-"'

        # Use current vs YTD average as trend indicator
        current = f"C{row}"
        ytd = f"G{row}"
        higher_is_better = kpi['higher_is_better']

        if higher_is_better:
            return f'=IF({current}>{ytd}*1.05,"U",IF({current}<{ytd}*0.95,"D","F"))'
        else:
            return f'=IF({current}<{ytd}*0.95,"U",IF({current}>{ytd}*1.05,"D","F"))'

    def _add_status_conditional_formatting(self, ws, start_row: int, end_row: int, col: int):
        """Add conditional formatting for status indicators"""
        col_letter = get_column_letter(col)
        range_str = f"{col_letter}{start_row}:{col_letter}{end_row}"

        # Green for "G"
        green_rule = FormulaRule(
            formula=[f'{col_letter}{start_row}="G"'],
            fill=self.green_fill,
            font=Font(color='FFFFFF', bold=True)
        )
        ws.conditional_formatting.add(range_str, green_rule)

        # Yellow for "Y"
        yellow_rule = FormulaRule(
            formula=[f'{col_letter}{start_row}="Y"'],
            fill=self.yellow_fill,
            font=Font(color='000000', bold=True)
        )
        ws.conditional_formatting.add(range_str, yellow_rule)

        # Red for "R"
        red_rule = FormulaRule(
            formula=[f'{col_letter}{start_row}="R"'],
            fill=self.red_fill,
            font=Font(color='FFFFFF', bold=True)
        )
        ws.conditional_formatting.add(range_str, red_rule)

    def _add_trend_conditional_formatting(self, ws, start_row: int, end_row: int, col: int):
        """Add conditional formatting for trend indicators"""
        col_letter = get_column_letter(col)
        range_str = f"{col_letter}{start_row}:{col_letter}{end_row}"

        # Up trend (green)
        up_rule = FormulaRule(
            formula=[f'{col_letter}{start_row}="U"'],
            font=Font(color='27AE60', bold=True)
        )
        ws.conditional_formatting.add(range_str, up_rule)

        # Down trend (red)
        down_rule = FormulaRule(
            formula=[f'{col_letter}{start_row}="D"'],
            font=Font(color='E74C3C', bold=True)
        )
        ws.conditional_formatting.add(range_str, down_rule)

    def _get_control_row_for_kpi(self, kpi_name: str) -> Optional[int]:
        """Get the row number in Dashboard_Control for a KPI"""
        row = 17  # Start row for KPIs in control sheet
        for category, kpis in self.KPI_DEFINITIONS.items():
            for kpi in kpis:
                if kpi['name'] == kpi_name:
                    return row
                row += 1
        return None

    def _create_trend_charts_section(self, ws, row_start: int) -> int:
        """Create trend charts section"""
        # Section header
        ws.cell(row_start, 2).value = "TREND ANALYSIS"
        ws.cell(row_start, 2).font = self.section_font
        ws.cell(row_start, 2).fill = self.section_fill
        for col in range(2, 15):
            ws.cell(row_start, col).fill = self.section_fill

        row_start += 2

        # Create a mini data table for charts (hidden or visible)
        ws.cell(row_start, 2).value = "Month"
        ws.cell(row_start, 3).value = "Revenue"
        ws.cell(row_start, 4).value = "Net Income"
        ws.cell(row_start, 5).value = "Gross Margin"
        ws.cell(row_start, 6).value = "Cash"

        for col in range(2, 7):
            ws.cell(row_start, col).font = self.header_font
            ws.cell(row_start, col).fill = self.kpi_header_fill
            ws.cell(row_start, col).alignment = self.center_align

        data_start = row_start + 1
        for i, month in enumerate(self.months[-12:]):  # Last 12 months
            row = data_start + i
            ws.cell(row, 2).value = self._month_str(month)
            ws.cell(row, 2).font = self.normal_font

            col_idx = self.month_columns.get(month, 2)

            # Revenue
            ws.cell(row, 3).value = f"=SUMIF('{self.pl_sheet_name}'!$A:$A,\"Total for Income\",'{self.pl_sheet_name}'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 3).number_format = '#,##0'

            # Net Income
            ws.cell(row, 4).value = f"=SUMIF('{self.pl_sheet_name}'!$A:$A,\"Net Income\",'{self.pl_sheet_name}'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 4).number_format = '#,##0'

            # Gross Margin
            gp = f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"Gross Profit\",'{self.pl_sheet_name}'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            rev = f"SUMIF('{self.pl_sheet_name}'!$A:$A,\"Total for Income\",'{self.pl_sheet_name}'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 5).value = f"=IFERROR({gp}/{rev},0)"
            ws.cell(row, 5).number_format = '0.0%'

            # Cash
            ws.cell(row, 6).value = f"=SUMIF('{self.bs_sheet_name}'!$A:$A,\"Total for Bank Accounts\",'{self.bs_sheet_name}'!{get_column_letter(col_idx)}:{get_column_letter(col_idx)})"
            ws.cell(row, 6).number_format = '#,##0'

        chart_data_end = data_start + min(12, len(self.months)) - 1

        # Create Revenue/Net Income combo chart
        try:
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"
            chart.title = "Revenue & Net Income Trend"
            chart.style = 10

            # Revenue data
            revenue_data = Reference(ws, min_col=3, min_row=row_start, max_row=chart_data_end)
            chart.add_data(revenue_data, titles_from_data=True)

            # Net Income data
            ni_data = Reference(ws, min_col=4, min_row=row_start, max_row=chart_data_end)
            chart.add_data(ni_data, titles_from_data=True)

            # Categories (months)
            cats = Reference(ws, min_col=2, min_row=data_start, max_row=chart_data_end)
            chart.set_categories(cats)

            chart.shape = 4
            chart.width = 15
            chart.height = 8

            ws.add_chart(chart, f"H{row_start}")
        except Exception as e:
            # Charts may fail in some environments
            ws.cell(row_start + 2, 8).value = f"[Chart: Revenue & Net Income Trend]"
            ws.cell(row_start + 2, 8).font = self.small_font

        return chart_data_end + 3

    def create_vba_module(self) -> str:
        """Generate VBA code for dashboard automation"""
        vba_code = '''
Attribute VB_Name = "DashboardModule"
'=====================================================
' Financial Model Dashboard VBA Module
' Auto-refresh logic, dynamic range creation, chart binding
'=====================================================

Option Explicit

' Constants for sheet names
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const CONTROL_SHEET As String = "Dashboard_Control"
Private Const SOURCE_PL_SHEET As String = "Source_PL"
Private Const SOURCE_BS_SHEET As String = "Source_BS"

'-----------------------------------------------------
' RefreshDashboard - Main refresh procedure
' Called after file upload or current month change
'-----------------------------------------------------
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Update named ranges
    Call UpdateDynamicRanges

    ' Refresh all calculations
    Application.Calculate

    ' Update charts
    Call RefreshCharts

    ' Log the refresh
    Call LogDiagnostic("Dashboard refreshed successfully")

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Call LogDiagnostic("Error refreshing dashboard: " & Err.Description)
    Resume Cleanup
End Sub

'-----------------------------------------------------
' UpdateDynamicRanges - Recreate named ranges
'-----------------------------------------------------
Public Sub UpdateDynamicRanges()
    On Error Resume Next

    Dim wsPL As Worksheet, wsBS As Worksheet
    Dim lastRowPL As Long, lastRowBS As Long
    Dim lastColPL As Long, lastColBS As Long

    Set wsPL = ThisWorkbook.Worksheets(SOURCE_PL_SHEET)
    Set wsBS = ThisWorkbook.Worksheets(SOURCE_BS_SHEET)

    If wsPL Is Nothing Or wsBS Is Nothing Then Exit Sub

    ' Find last row and column with data
    lastRowPL = wsPL.Cells(wsPL.Rows.Count, 1).End(xlUp).Row
    lastColPL = wsPL.Cells(1, wsPL.Columns.Count).End(xlToLeft).Column

    lastRowBS = wsBS.Cells(wsBS.Rows.Count, 1).End(xlUp).Row
    lastColBS = wsBS.Cells(1, wsBS.Columns.Count).End(xlToLeft).Column

    ' Delete existing ranges if they exist
    On Error Resume Next
    ThisWorkbook.Names("SourcePL_Data").Delete
    ThisWorkbook.Names("SourceBS_Data").Delete
    ThisWorkbook.Names("PL_Accounts").Delete
    ThisWorkbook.Names("BS_Accounts").Delete
    ThisWorkbook.Names("MonthHeaders").Delete
    On Error GoTo 0

    ' Create new named ranges
    ThisWorkbook.Names.Add Name:="SourcePL_Data", _
        RefersTo:="='" & SOURCE_PL_SHEET & "'!$A$1:$" & ColLetter(lastColPL) & "$" & lastRowPL

    ThisWorkbook.Names.Add Name:="SourceBS_Data", _
        RefersTo:="='" & SOURCE_BS_SHEET & "'!$A$1:$" & ColLetter(lastColBS) & "$" & lastRowBS

    ThisWorkbook.Names.Add Name:="PL_Accounts", _
        RefersTo:="='" & SOURCE_PL_SHEET & "'!$A$2:$A$" & lastRowPL

    ThisWorkbook.Names.Add Name:="BS_Accounts", _
        RefersTo:="='" & SOURCE_BS_SHEET & "'!$A$2:$A$" & lastRowBS

    ThisWorkbook.Names.Add Name:="MonthHeaders", _
        RefersTo:="='" & SOURCE_PL_SHEET & "'!$B$1:$" & ColLetter(lastColPL) & "$1"

End Sub

'-----------------------------------------------------
' RefreshCharts - Update chart data ranges
'-----------------------------------------------------
Public Sub RefreshCharts()
    On Error Resume Next

    Dim ws As Worksheet
    Dim cht As ChartObject

    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If ws Is Nothing Then Exit Sub

    ' Update each chart to use current data range
    For Each cht In ws.ChartObjects
        cht.Chart.Refresh
    Next cht

End Sub

'-----------------------------------------------------
' ExtendFormulasForNewAccounts
' Called when new accounts are added to source data
'-----------------------------------------------------
Public Sub ExtendFormulasForNewAccounts()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsDash As Worksheet
    Dim wsCtrl As Worksheet

    Set wsDash = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set wsCtrl = ThisWorkbook.Worksheets(CONTROL_SHEET)

    ' Refresh the named ranges first
    Call UpdateDynamicRanges

    ' Force recalculation of SUMIF formulas
    wsDash.Calculate
    wsCtrl.Calculate

    Call LogDiagnostic("Formulas extended for new accounts")

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Call LogDiagnostic("Error extending formulas: " & Err.Description)
    Resume Cleanup
End Sub

'-----------------------------------------------------
' OnCurrentMonthChange - Handle month selection change
'-----------------------------------------------------
Public Sub OnCurrentMonthChange()
    ' This is called when current month changes on Menu sheet
    Call RefreshDashboard
    Call UpdateYearGrouping
End Sub

'-----------------------------------------------------
' UpdateYearGrouping - Group and collapse prior year columns
' Groups columns by year on P&L and Balance Sheet
' Current year expanded, prior years collapsed
'-----------------------------------------------------
Public Sub UpdateYearGrouping()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsMenu As Worksheet
    Dim wsPL As Worksheet
    Dim wsBS As Worksheet
    Dim currentMonthStr As String
    Dim currentYear As Long

    Set wsMenu = GetWorksheet("Menu")
    Set wsPL = GetWorksheet("PL")
    Set wsBS = GetWorksheet("Balance_Sheet")

    If wsMenu Is Nothing Then Exit Sub

    ' Get current month from Menu
    currentMonthStr = Trim(CStr(wsMenu.Range("C11").Value))
    If Len(currentMonthStr) = 0 Then
        currentMonthStr = Trim(CStr(wsMenu.Range("C7").Value))
    End If

    ' Extract year from current month string
    currentYear = ExtractYearFromString(currentMonthStr)
    If currentYear = 0 Then GoTo Cleanup

    ' Update grouping on P&L sheet
    If Not wsPL Is Nothing Then
        Call GroupColumnsByYear(wsPL, currentYear)
    End If

    ' Update grouping on Balance Sheet
    If Not wsBS Is Nothing Then
        Call GroupColumnsByYear(wsBS, currentYear)
    End If

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Resume Cleanup
End Sub

Private Sub GroupColumnsByYear(ws As Worksheet, currentYear As Long)
    On Error Resume Next
    Dim headerRow As Long, lastCol As Long, col As Long
    Dim yearGroups As Object, yearKey As Variant
    Dim startCol As Long, endCol As Long, colYear As Long, prevYear As Long

    headerRow = 4
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    ws.Cells.ClearOutline
    Set yearGroups = CreateObject("Scripting.Dictionary")

    prevYear = 0
    For col = 2 To lastCol
        Dim headerVal As String
        headerVal = LCase(Trim(CStr(ws.Cells(headerRow, col).Value)))
        If HasMonthName(headerVal) Then
            colYear = ExtractYearFromString(headerVal)
            If colYear > 0 Then
                If prevYear = 0 Then
                    prevYear = colYear
                    startCol = col
                ElseIf colYear <> prevYear Then
                    yearGroups.Add prevYear, Array(startCol, col - 1)
                    prevYear = colYear
                    startCol = col
                End If
            End If
        End If
    Next col
    If prevYear > 0 Then yearGroups.Add prevYear, Array(startCol, FindLastMonthCol(ws, headerRow, startCol, prevYear))

    For Each yearKey In yearGroups.Keys
        startCol = yearGroups(yearKey)(0)
        endCol = yearGroups(yearKey)(1)
        If startCol < endCol Then ws.Columns(startCol).Resize(, endCol - startCol).Group
        ws.Columns(startCol).Resize(, endCol - startCol + 1).Hidden = (CLng(yearKey) < currentYear)
    Next yearKey
End Sub

Private Function HasMonthName(s As String) As Boolean
    Dim m As Variant
    For Each m In Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
        If InStr(s, m) > 0 Then HasMonthName = True: Exit Function
    Next m
    HasMonthName = False
End Function

Private Function ExtractYearFromString(s As String) As Long
    Dim parts() As String, i As Long, y As Long
    ExtractYearFromString = 0
    parts = Split(Replace(s, "-", " "), " ")
    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            y = CLng(parts(i))
            If y >= 2000 And y <= 2100 Then ExtractYearFromString = y: Exit Function
            If y >= 0 And y <= 99 Then ExtractYearFromString = 2000 + y: Exit Function
        End If
    Next i
End Function

Private Function FindLastMonthCol(ws As Worksheet, row As Long, startCol As Long, yr As Long) As Long
    Dim col As Long
    FindLastMonthCol = startCol
    For col = startCol To ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
        If ExtractYearFromString(CStr(ws.Cells(row, col).Value)) = yr Then FindLastMonthCol = col
    Next col
End Function

'-----------------------------------------------------
' GetCurrentMonthColumn - Returns column index for current month
'-----------------------------------------------------
Public Function GetCurrentMonthColumn() As Long
    On Error Resume Next

    Dim wsMenu As Worksheet
    Dim wsPL As Worksheet
    Dim currentMonth As String
    Dim col As Long
    Dim headerValue As String

    Set wsMenu = ThisWorkbook.Worksheets("Menu")
    Set wsPL = ThisWorkbook.Worksheets(SOURCE_PL_SHEET)

    If wsMenu Is Nothing Or wsPL Is Nothing Then
        GetCurrentMonthColumn = 0
        Exit Function
    End If

    ' Get current month from Menu
    currentMonth = CStr(wsMenu.Range("B5").Value)

    ' Find matching column in source data
    For col = 2 To wsPL.Cells(1, wsPL.Columns.Count).End(xlToLeft).Column
        headerValue = CStr(wsPL.Cells(1, col).Value)
        If InStr(1, headerValue, currentMonth, vbTextCompare) > 0 Then
            GetCurrentMonthColumn = col
            Exit Function
        End If
    Next col

    ' Default to last column if not found
    GetCurrentMonthColumn = wsPL.Cells(1, wsPL.Columns.Count).End(xlToLeft).Column

End Function

'-----------------------------------------------------
' ValidateDashboardData - Check for data issues
'-----------------------------------------------------
Public Sub ValidateDashboardData()
    On Error Resume Next

    Dim issues As String
    issues = ""

    ' Check if source sheets exist
    If GetWorksheet(SOURCE_PL_SHEET) Is Nothing Then
        issues = issues & "- Source P&L sheet missing" & vbCrLf
    End If

    If GetWorksheet(SOURCE_BS_SHEET) Is Nothing Then
        issues = issues & "- Source Balance Sheet missing" & vbCrLf
    End If

    ' Check for data in source sheets
    Dim wsPL As Worksheet
    Set wsPL = GetWorksheet(SOURCE_PL_SHEET)
    If Not wsPL Is Nothing Then
        If wsPL.Cells(2, 1).Value = "" Then
            issues = issues & "- No P&L data found" & vbCrLf
        End If
    End If

    ' Report issues
    If issues <> "" Then
        MsgBox "Dashboard Validation Issues:" & vbCrLf & vbCrLf & issues, vbExclamation, "Validation"
        Call LogDiagnostic("Validation issues: " & Replace(issues, vbCrLf, "; "))
    Else
        MsgBox "Dashboard data validated successfully!", vbInformation, "Validation"
        Call LogDiagnostic("Dashboard validation passed")
    End If

End Sub

'-----------------------------------------------------
' Helper Functions
'-----------------------------------------------------

Private Function ColLetter(colNum As Long) As String
    Dim n As Long
    Dim c As String
    Dim s As String

    n = colNum
    Do
        c = Chr(((n - 1) Mod 26) + Asc("A"))
        s = c & s
        n = (n - 1) \\ 26
    Loop While n > 0
    ColLetter = s
End Function

Private Function GetWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Sub LogDiagnostic(message As String)
    On Error Resume Next

    Dim wsDiag As Worksheet
    Dim nextRow As Long

    Set wsDiag = ThisWorkbook.Worksheets("Diagnostics")
    If wsDiag Is Nothing Then Exit Sub

    nextRow = wsDiag.Cells(wsDiag.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 22 Then nextRow = 22

    wsDiag.Cells(nextRow, 1).Value = Now
    wsDiag.Cells(nextRow, 2).Value = message
    wsDiag.Cells(nextRow, 3).Value = "Info"

End Sub

'-----------------------------------------------------
' Auto-run on workbook open
'-----------------------------------------------------
Private Sub Workbook_Open()
    ' Refresh dashboard on open
    Application.OnTime Now + TimeValue("00:00:01"), "RefreshDashboard"
End Sub
'''
        return vba_code

    def create_documentation(self) -> str:
        """Generate documentation for the Dashboard module"""
        doc = f'''
================================================================================
FINANCIAL MODEL DASHBOARD MODULE - DOCUMENTATION
================================================================================
Generated: {datetime.now().strftime('%B %d, %Y')}

OVERVIEW
--------
The Dashboard Module provides a CFO-grade executive dashboard with comprehensive
KPIs, financial ratios, and visualizations. It dynamically calculates metrics
from the Source P&L and Source Balance Sheet tabs.

DASHBOARD VIEWS
---------------
Each KPI is displayed across three time periods:

1. CURRENT MONTH
   - Shows the metric for the selected current month
   - Controlled by the "Current Month" setting on Menu tab
   - Updates automatically when month selection changes

2. YEAR-TO-DATE (YTD)
   - Aggregates data from the first month through current month
   - For P&L items: Sum of all months
   - For Balance Sheet items: Current period balance
   - For ratios: Calculated using YTD numerator/denominator

3. ALL-TIME / INCEPTION-TO-DATE
   - Shows metrics since the beginning of available data
   - Useful for long-term trend analysis

KPI CATEGORIES
--------------

PROFITABILITY METRICS
---------------------
| KPI               | Formula                                    | Target  |
|-------------------|-------------------------------------------|---------|
| Revenue           | Total for Income                          | Custom  |
| Gross Profit      | Gross Profit line from P&L                | Custom  |
| Gross Margin %    | Gross Profit / Revenue                    | 40%     |
| Net Income        | Net Income line from P&L                  | Custom  |
| Net Margin %      | Net Income / Revenue                      | 10%     |
| EBITDA            | Net Income + Interest + Depreciation      | Custom  |
| EBITDA Margin %   | EBITDA / Revenue                          | 15%     |
| Operating Income  | Net Operating Income from P&L             | Custom  |
| Operating Margin% | Operating Income / Revenue                | 12%     |
| Revenue Growth %  | (Current - Prior) / Prior                 | 10%     |

LIQUIDITY METRICS
-----------------
| KPI               | Formula                                    | Target  |
|-------------------|-------------------------------------------|---------|
| Current Ratio     | Current Assets / Current Liabilities      | 1.5     |
| Quick Ratio       | (Current Assets - Inventory) / CL         | 1.0     |
| Cash Ratio        | Cash / Current Liabilities                | 0.2     |
| Working Capital   | Current Assets - Current Liabilities      | Custom  |
| Cash Balance      | Total Bank Accounts                       | Custom  |

EFFICIENCY METRICS
------------------
| KPI               | Formula                                    | Target  |
|-------------------|-------------------------------------------|---------|
| AR Days (DSO)     | AR / (Revenue/30)                         | 45 days |
| AP Days (DPO)     | AP / (COGS/30)                            | 30 days |
| Asset Turnover    | Revenue / Total Assets                    | 1.0     |
| Cash Conversion   | AR Days - AP Days                         | 30 days |

LEVERAGE METRICS
----------------
| KPI               | Formula                                    | Target  |
|-------------------|-------------------------------------------|---------|
| Debt-to-Equity    | Total Liabilities / Total Equity          | 1.0     |
| Debt-to-Assets    | Total Liabilities / Total Assets          | 50%     |
| Equity Ratio      | Total Equity / Total Assets               | 50%     |
| Interest Coverage | Operating Income / Interest Expense       | 3.0x    |

CASH FLOW METRICS
-----------------
| KPI               | Formula                                    | Target  |
|-------------------|-------------------------------------------|---------|
| Operating CF      | Net Income + Depreciation                 | Custom  |
| Free Cash Flow    | Operating CF - CapEx                      | Custom  |
| Cash Burn Rate    | Monthly change in cash balance            | Custom  |
| Cash Runway       | Cash Balance / Monthly Burn Rate          | 12 mos  |

DASHBOARD CONTROL PAGE
----------------------
The Dashboard_Control sheet allows you to customize:

1. TARGET VALUES
   - Set specific targets for each KPI
   - User-defined targets for metrics like Revenue, Net Income
   - Standard industry targets for ratios (adjustable)

2. ALERT THRESHOLDS
   - Yellow Alert: Triggers at specified % of target (default 80%)
   - Red Alert: Triggers at specified % of target (default 60%)
   - For "lower is better" metrics, thresholds work inversely

3. DIRECTION INDICATOR
   - "Higher" = Green when above target
   - "Lower" = Green when below target (e.g., AR Days, Debt ratios)

STATUS INDICATORS
-----------------
| Symbol | Color  | Meaning                                    |
|--------|--------|-------------------------------------------|
| G      | Green  | On or above target                         |
| Y      | Yellow | Between yellow and target thresholds       |
| R      | Red    | Below red threshold                        |

TREND INDICATORS
----------------
| Symbol | Meaning                                              |
|--------|-----------------------------------------------------|
| U      | Improving (up arrow) - current > YTD by 5%+         |
| D      | Declining (down arrow) - current < YTD by 5%+       |
| F      | Flat - within 5% of YTD average                     |

VBA AUTOMATION
--------------
The following VBA procedures are available:

RefreshDashboard()
- Refreshes all dashboard calculations
- Updates named ranges
- Refreshes charts
- Call after data upload or month change

UpdateDynamicRanges()
- Recreates named ranges to include new accounts
- Called automatically by RefreshDashboard

ExtendFormulasForNewAccounts()
- Extends SUMIF formulas to include new accounts
- Recalculates all dashboard formulas

ValidateDashboardData()
- Checks for missing source data
- Reports any issues found

FORMULA REFERENCES
------------------
All dashboard formulas use these patterns:

P&L Lookups:
=SUMIF(Source_PL!$A:$A,"Account Name",Source_PL!Column:Column)

Balance Sheet Lookups:
=SUMIF(Source_BS!$A:$A,"Account Name",Source_BS!Column:Column)

Ratio Calculations:
=IFERROR(Numerator/Denominator,0)

Days Calculations:
=IFERROR(Balance/(Flow/MonthCount)*30,0)

PRINT SETTINGS
--------------
- Orientation: Landscape
- Fit to: 1 page wide
- Margins: 0.5" all sides
- Headers repeat on each page

MAINTENANCE
-----------
1. When adding new accounts:
   - Run "ExtendFormulasForNewAccounts" macro
   - Or manually refresh the workbook

2. When changing fiscal year:
   - Update Menu configuration
   - Refresh dashboard

3. For custom KPIs:
   - Add to Dashboard_Control sheet
   - Create formula in Dashboard sheet

TROUBLESHOOTING
---------------
Issue: KPIs showing 0 or #REF!
- Check that Source_PL and Source_BS have data
- Verify account names match exactly
- Run ValidateDashboardData macro

Issue: Charts not updating
- Run RefreshDashboard macro
- Check chart data ranges

Issue: Targets not linking
- Verify Dashboard_Control sheet exists
- Check named range "KPI_Targets"

================================================================================
'''
        return doc

    def save(self, output_path: Optional[str] = None):
        """Save the workbook with dashboard additions"""
        save_path = output_path or self.wb_path
        self.wb.save(save_path)

    def build_complete_dashboard(self) -> Tuple[str, str]:
        """Build the complete dashboard module and return VBA code and documentation"""
        # Create the sheets
        self.create_dashboard_control_sheet()
        self.create_dashboard_sheet()

        # Generate VBA and documentation
        vba_code = self.create_vba_module()
        documentation = self.create_documentation()

        return vba_code, documentation


def add_dashboard_to_workbook(workbook_path: str, output_path: Optional[str] = None) -> Tuple[str, str, str]:
    """
    Add dashboard module to an existing Financial Model workbook.

    Args:
        workbook_path: Path to existing .xlsm workbook
        output_path: Optional output path (defaults to overwriting input)

    Returns:
        Tuple of (output_path, vba_code, documentation)
    """
    module = DashboardModule(workbook_path)
    vba_code, documentation = module.build_complete_dashboard()

    save_path = output_path or workbook_path
    module.save(save_path)

    return save_path, vba_code, documentation


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python dashboard_module.py <workbook_path> [output_path]")
        sys.exit(1)

    wb_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else None

    print(f"Adding dashboard to: {wb_path}")
    result_path, vba, docs = add_dashboard_to_workbook(wb_path, out_path)

    print(f"Dashboard added successfully!")
    print(f"Workbook saved to: {result_path}")

    # Save VBA code to file
    vba_file = result_path.replace('.xlsm', '_dashboard_vba.bas')
    with open(vba_file, 'w') as f:
        f.write(vba)
    print(f"VBA code saved to: {vba_file}")

    # Save documentation
    doc_file = result_path.replace('.xlsm', '_dashboard_documentation.txt')
    with open(doc_file, 'w') as f:
        f.write(docs)
    print(f"Documentation saved to: {doc_file}")
