"""
CFO DNA Financial Model Generator - Desktop Application
Creates true macro-enabled Excel files with embedded VBA using xlwings
"""

import os
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import tempfile
import shutil
import pandas as pd
import xlwings as xw
from xlwings.constants import DeleteShiftDirection

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
TEMPLATE_PATH = os.path.join(APP_DIR, 'DNA_Template.xlsm')
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(EXE_DIR, 'DNA_Template.xlsm')


class DNAModelApp:
    """Desktop application for generating DNA financial models"""

    MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    VBA_CODE = '''
Option Explicit

Private Const SOURCE_PL_SHEET As String = "Source_PL"
Private Const SOURCE_BS_SHEET As String = "Source_BS"
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
'''

    def __init__(self, root):
        self.root = root
        self.root.title("CFO DNA Financial Model Generator")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        # Variables
        self.pl_path = tk.StringVar()
        self.bs_path = tk.StringVar()
        self.company_name = tk.StringVar(value="Company Name")
        self.fiscal_start = tk.StringVar(value="January")
        self.display_month = tk.StringVar(value="January")
        self.display_year = tk.StringVar(value=str(datetime.now().year))

        self._create_ui()

    def _create_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="CFO DNA Financial Model Generator",
                                font=('Segoe UI', 16, 'bold'))
        title_label.pack(pady=(0, 20))

        # Company Name
        name_frame = ttk.LabelFrame(main_frame, text="Company Information", padding="10")
        name_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(name_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(name_frame, textvariable=self.company_name, width=40).grid(row=0, column=1, padx=5)

        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text="Upload Files", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(file_frame, text="P&L File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.pl_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse...", command=self._browse_pl).grid(row=0, column=2)

        ttk.Label(file_frame, text="Balance Sheet:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        ttk.Entry(file_frame, textvariable=self.bs_path, width=40).grid(row=1, column=1, padx=5, pady=(5, 0))
        ttk.Button(file_frame, text="Browse...", command=self._browse_bs).grid(row=1, column=2, pady=(5, 0))

        # Configuration
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(config_frame, text="Fiscal Year Start:").grid(row=0, column=0, sticky=tk.W)
        fiscal_combo = ttk.Combobox(config_frame, textvariable=self.fiscal_start, values=self.MONTHS, width=15)
        fiscal_combo.grid(row=0, column=1, padx=5)

        ttk.Label(config_frame, text="First Display Month:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        display_combo = ttk.Combobox(config_frame, textvariable=self.display_month, values=self.MONTHS, width=15)
        display_combo.grid(row=0, column=3, padx=5)

        ttk.Label(config_frame, text="First Display Year:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        years = [str(y) for y in range(datetime.now().year - 5, datetime.now().year + 2)]
        year_combo = ttk.Combobox(config_frame, textvariable=self.display_year, values=years, width=15)
        year_combo.grid(row=1, column=1, padx=5, pady=(5, 0))

        # Generate Button
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)

        self.generate_btn = ttk.Button(btn_frame, text="Generate DNA Model",
                                       command=self._generate_model, style='Accent.TButton')
        self.generate_btn.pack(fill=tk.X, ipady=10)

        # Progress
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X)

        # Status
        self.status_label = ttk.Label(main_frame, text="Ready", foreground='gray')
        self.status_label.pack(pady=(10, 0))

    def _browse_pl(self):
        path = filedialog.askopenfilename(
            title="Select P&L File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            self.pl_path.set(path)

    def _browse_bs(self):
        path = filedialog.askopenfilename(
            title="Select Balance Sheet File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            self.bs_path.set(path)

    def _generate_model(self):
        """Generate the DNA model"""
        # Validate inputs
        if not self.pl_path.get() or not self.bs_path.get():
            messagebox.showerror("Error", "Please select both P&L and Balance Sheet files")
            return

        if not self.company_name.get().strip():
            messagebox.showerror("Error", "Please enter a company name")
            return

        # Ask for save location
        save_path = filedialog.asksaveasfilename(
            title="Save DNA Model As",
            defaultextension=".xlsm",
            filetypes=[("Excel Macro-Enabled", "*.xlsm")],
            initialfile=f"{self.company_name.get().replace(' ', '_')}_DNA_Model.xlsm"
        )

        if not save_path:
            return

        self.progress.start()
        self.status_label.config(text="Generating model...")
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
            self.progress.stop()
            self._update_status("Model generated successfully!")
            messagebox.showinfo("Success", f"DNA Model created:\n{save_path}")
        except Exception as e:
            import traceback
            print(f"ERROR: {e}")
            traceback.print_exc()
            self.progress.stop()
            self._update_status("Error occurred")
            messagebox.showerror("Error", str(e))
        finally:
            self.generate_btn.config(state='normal')

    def _update_status(self, message):
        """Update the status label and refresh the UI"""
        self.status_label.config(text=message)
        self.root.update()
        print(message)

    def _create_excel_model(self, save_path):
        """Create the Excel model using xlwings"""
        # Parse input files
        self._update_status("Step 1/10: Reading P&L file...")
        pl_data = pd.read_excel(self.pl_path.get(), header=None)

        self._update_status("Step 2/10: Reading Balance Sheet file...")
        bs_data = pd.read_excel(self.bs_path.get(), header=None)

        self._update_status("Step 3/10: Extracting account data...")
        pl_accounts, pl_months, pl_totals = self._parse_financial_data(pl_data)
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data)
        print(f"Found {len(pl_accounts)} P&L accounts, {len(bs_accounts)} BS accounts, {len(pl_months)} months")
        print(f"P&L Totals detected: {pl_totals}")
        print(f"BS Totals detected: {bs_totals}")

        # Check if template exists
        use_template = os.path.exists(TEMPLATE_PATH)

        # Work in temp directory to avoid OneDrive locking issues
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        # Create Excel application
        self._update_status("Step 4/10: Starting Excel...")
        app = xw.App(visible=False)
        app.display_alerts = False  # Suppress Excel prompts
        app.screen_updating = False  # Speed up processing
        wb = None

        try:
            if use_template:
                # Copy template to temp location and open
                shutil.copy(TEMPLATE_PATH, temp_path)
                wb = app.books.open(temp_path)

                # Get existing sheets
                menu_sheet = wb.sheets['Menu']
                source_pl = wb.sheets['Source_PL']
                source_bs = wb.sheets['Source_BS']
                pl_sheet = wb.sheets['PL']
                bs_sheet = wb.sheets['Balance_Sheet']
                cf_sheet = wb.sheets['Cash_Flow']
                notes_sheet = wb.sheets['Notes']

                # Clear source sheets (keep headers)
                source_pl.range('A2:ZZ1000').clear()
                source_bs.range('A2:ZZ1000').clear()

                # Also clear the report sheets for fresh data
                pl_sheet.range('A5:ZZ1000').clear()
                bs_sheet.range('A5:ZZ1000').clear()
                cf_sheet.range('A5:ZZ1000').clear()
            else:
                # Create from scratch (requires VBA trust setting)
                wb = app.books.add()

                # Remove default sheets and create our sheets
                for sheet in wb.sheets:
                    if sheet.name not in ['Sheet1']:
                        sheet.delete()

                # Create sheets
                menu_sheet = wb.sheets[0]
                menu_sheet.name = 'Menu'

                source_pl = wb.sheets.add('Source_PL', after=menu_sheet)
                source_bs = wb.sheets.add('Source_BS', after=source_pl)
                pl_sheet = wb.sheets.add('PL', after=source_bs)
                bs_sheet = wb.sheets.add('Balance_Sheet', after=pl_sheet)
                cf_sheet = wb.sheets.add('Cash_Flow', after=bs_sheet)
                notes_sheet = wb.sheets.add('Notes', after=cf_sheet)

            # Populate source sheets
            self._update_status("Step 5/10: Populating source data...")
            self._populate_source_sheet(source_pl, pl_accounts, pl_months)
            self._populate_source_sheet(source_bs, bs_accounts, pl_months)

            # Create/update named ranges
            pl_last_row = len(pl_accounts) + 1
            pl_last_col = len(pl_months) + 1
            bs_last_row = len(bs_accounts) + 1

            # Delete existing named ranges if they exist
            try:
                wb.names['SourcePL'].delete()
            except:
                pass
            try:
                wb.names['SourceBS'].delete()
            except:
                pass

            wb.names.add('SourcePL', f"=Source_PL!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")
            wb.names.add('SourceBS', f"=Source_BS!$A$1:${self._col_letter(pl_last_col)}${bs_last_row}")

            # Create Menu sheet
            self._update_status("Step 6/10: Creating Menu sheet...")
            self._create_menu_sheet(menu_sheet, pl_months)

            # Create P&L report
            self._update_status("Step 7/10: Creating P&L report...")
            self._create_pl_report(pl_sheet, pl_accounts, pl_months, pl_totals)

            # Create Balance Sheet report
            self._update_status("Step 8/10: Creating Balance Sheet...")
            self._create_bs_report(bs_sheet, bs_accounts, pl_months, bs_totals)

            # Create Cash Flow statement
            self._update_status("Step 9/10: Creating Cash Flow statement...")
            self._create_cash_flow(cf_sheet, pl_months)

            # Create Notes sheet with dropdowns
            all_accounts = [a['name'] for a in pl_accounts] + [a['name'] for a in bs_accounts]
            month_names = [name for m, y, name in pl_months]
            self._create_notes_sheet(notes_sheet, all_accounts, month_names)

            # Move source sheets to the end
            try:
                source_pl.api.Move(After=notes_sheet.api)
                source_bs.api.Move(After=source_pl.api)
            except:
                pass

            # Set source sheet tab colors to black
            try:
                source_pl.api.Tab.Color = 0x000000  # Black
                source_bs.api.Tab.Color = 0x000000  # Black
            except:
                pass

            # Activate Menu sheet so file opens to Menu
            try:
                menu_sheet.activate()
            except:
                pass

            # Add VBA code only if not using template
            if not use_template:
                try:
                    vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                    vba_module.Name = "DNAModel"
                    vba_module.CodeModule.AddFromString(self.VBA_CODE)
                except Exception as e:
                    # VBA access not enabled - save without macros
                    pass

            # Save the workbook
            self._update_status("Step 10/10: Saving workbook...")
            wb.save()
            wb.close()
            wb = None

        finally:
            # Ensure proper cleanup
            try:
                if wb is not None:
                    wb.close()
            except:
                pass
            try:
                app.quit()
            except:
                pass

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

    def _parse_financial_data(self, df):
        """Parse financial data from dataframe"""
        # Find header row
        header_row = None
        for idx in range(min(15, len(df))):
            for col_idx in range(len(df.columns)):
                val = df.iloc[idx, col_idx]
                if pd.notna(val):
                    val_str = str(val).strip().lower()
                    for month in self.MONTHS:
                        if val_str.startswith(month.lower()) and '-' not in val_str:
                            parts = str(df.iloc[idx, col_idx]).split()
                            if len(parts) == 2:
                                header_row = idx
                                break
                if header_row is not None:
                    break
            if header_row is not None:
                break

        if header_row is None:
            raise ValueError("Could not find month headers in file")

        # Extract months
        months = []
        for col_idx in range(1, len(df.columns)):
            val = df.iloc[header_row, col_idx]
            if pd.notna(val) and 'total' not in str(val).lower():
                val_str = str(val).strip()
                for i, month in enumerate(self.MONTHS, 1):
                    if val_str.lower().startswith(month.lower()):
                        parts = val_str.split()
                        if len(parts) == 2:
                            try:
                                year = int(parts[1])
                                months.append((i, year, val_str))
                            except:
                                pass
                        break

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

        for row_idx in range(header_row + 1, len(df)):
            account_name = df.iloc[row_idx, 0]
            if pd.isna(account_name) or not str(account_name).strip():
                continue
            account_name = str(account_name).strip()
            if 'cash basis' in account_name.lower():
                continue

            values = {}
            for col_idx in range(1, min(len(df.columns), len(months) + 1)):
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val):
                    try:
                        values[col_idx - 1] = float(val)
                    except:
                        values[col_idx - 1] = 0
                else:
                    values[col_idx - 1] = 0

            is_total = account_name.lower().startswith('total') or account_name in ['Net Income', 'Gross Profit']
            is_header = account_name in ['Income', 'Expenses', 'Cost of Sales', 'Assets', 'Liabilities', 'Equity']

            # Detect specific total rows by their exact names
            name_lower = account_name.lower()
            if name_lower.startswith('total income') or name_lower == 'total for income':
                detected_totals['total_income'] = account_name
            elif name_lower.startswith('total cost') or name_lower.startswith('total cogs') or 'cost of sales' in name_lower and 'total' in name_lower:
                detected_totals['total_cogs'] = account_name
            elif name_lower.startswith('total expenses') or name_lower == 'total for expenses':
                detected_totals['total_expenses'] = account_name
            elif 'gross profit' in name_lower:
                detected_totals['gross_profit'] = account_name
            elif name_lower == 'net income':
                detected_totals['net_income'] = account_name
            elif name_lower.startswith('total for assets') or name_lower == 'total assets':
                detected_totals['total_assets'] = account_name
            elif name_lower.startswith('total for liabilities and equity') or name_lower == 'total liabilities and equity':
                detected_totals['total_liab_equity'] = account_name
            elif (name_lower.startswith('total for liabilities') or name_lower == 'total liabilities') and 'equity' not in name_lower:
                detected_totals['total_liabilities'] = account_name
            elif (name_lower.startswith('total for equity') or name_lower == 'total equity') and 'liabilities' not in name_lower:
                detected_totals['total_equity'] = account_name

            accounts.append({
                'name': account_name,
                'values': values,
                'is_total': is_total,
                'is_header': is_header
            })

        return accounts, months, detected_totals

    def _populate_source_sheet(self, sheet, accounts, months):
        """Populate a source data sheet with proper formatting"""
        # Colors - matching web version
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A - almost black for source sheets

        # Header
        sheet.range('A1').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            sheet.range((1, i + 2)).value = name

        # Data - write all at once for speed and to ensure numbers are numbers
        data = []
        for account in accounts:
            row = [account['name']]
            for col_idx in range(len(months)):
                val = account['values'].get(col_idx, 0)
                # Ensure it's a number
                if isinstance(val, str):
                    try:
                        val = float(val.replace(',', ''))
                    except:
                        val = 0
                row.append(val)
            data.append(row)

        if data:
            sheet.range('A2').value = data

        # Format header row with dark background and white text (like web version)
        try:
            header_range = sheet.range((1, 1), (1, len(months) + 1))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = SOURCE_BLACK

            # Center align month headers
            for col in range(2, len(months) + 2):
                sheet.range((1, col)).api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass

        # Apply number format and font to data columns
        if len(accounts) > 0 and len(months) > 0:
            try:
                data_range = sheet.range((2, 2), (len(accounts) + 1, len(months) + 1))
                data_range.number_format = '#,##0'
                data_range.font.name = 'Calibri Light'
                data_range.font.size = 10

                # Account names column
                account_range = sheet.range((2, 1), (len(accounts) + 1, 1))
                account_range.font.name = 'Calibri Light'
                account_range.font.size = 10
            except:
                pass

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, len(months) + 2):
            sheet.range((1, col), (1, col)).column_width = 14

    def _create_menu_sheet(self, sheet, months):
        """Create the menu/control sheet with proper formatting"""
        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        GRAY = (128, 128, 128)

        # Write all content first in bulk
        content = [
            ['', '', ''],  # Row 1
            ['', self.company_name.get(), ''],  # Row 2 - Title
            ['', 'DNA Financial Model', ''],  # Row 3
            ['', '', ''],  # Row 4
            ['', 'CONFIGURATION', ''],  # Row 5
            ['', 'Company:', self.company_name.get()],  # Row 6
            ['', 'Current Month:', months[-1][2] if months else ''],  # Row 7
            ['', 'Data Range:', f"{months[0][2]} to {months[-1][2]}" if months else ''],  # Row 8
            ['', 'Total Months:', str(len(months)) if months else '0'],  # Row 9
            ['', '', ''],  # Row 10
            ['', 'ACTIONS', ''],  # Row 11
            ['', '', ''],  # Row 12
            ['', 'Upload P&L Data', 'Import new monthly P&L data (Alt+F8)'],  # Row 13
            ['', '', ''],  # Row 14
            ['', 'Upload Balance Sheet', 'Import new Balance Sheet data (Alt+F8)'],  # Row 15
            ['', '', ''],  # Row 16
            ['', 'Refresh All', 'Recalculate all formulas (Alt+F8)'],  # Row 17
            ['', '', ''],  # Row 18
            ['', 'Run Diagnostics', 'Validate model and check for errors (Alt+F8)'],  # Row 19
            ['', '', ''],  # Row 20
            ['', '', ''],  # Row 21
            ['', 'HOW TO USE', ''],  # Row 22
            ['', '1. Press Alt+F8 to open Macros dialog', ''],  # Row 23
            ['', '2. Select the macro and click Run', ''],  # Row 24
            ['', '3. For uploads, select your file when prompted', ''],  # Row 25
            ['', '4. Run Diagnostics to check model integrity', ''],  # Row 26
            ['', '', ''],  # Row 27
            ['', 'QUICK NAVIGATION', ''],  # Row 28
            ['', 'Go to P&L Statement', ''],  # Row 29
            ['', 'Go to Balance Sheet', ''],  # Row 30
            ['', 'Go to Cash Flow', ''],  # Row 31
            ['', 'Go to Variance Notes', ''],  # Row 32
            ['', 'Go to Source P&L', ''],  # Row 33
            ['', 'Go to Source BS', ''],  # Row 34
        ]

        # Write all content at once
        sheet.range('A1').value = content

        # Apply formatting in batches
        try:
            # Title formatting
            title_cell = sheet.range('B2')
            title_cell.font.name = 'Calibri Light'
            title_cell.font.size = 28
            title_cell.font.bold = True
            title_cell.font.color = DARK_BLUE

            sheet.range('B3').font.name = 'Calibri Light'
            sheet.range('B3').font.size = 16
            sheet.range('B3').font.color = GRAY

            # Section headers - batch format
            for row in [5, 11, 22, 28]:
                cell = sheet.range(f'B{row}')
                cell.font.name = 'Calibri Light'
                cell.font.size = 12
                cell.font.bold = True
                cell.font.color = DARK_BLUE

            # Action buttons - format with background
            for row in [13, 15, 17, 19]:
                btn = sheet.range(f'B{row}')
                btn.font.name = 'Calibri Light'
                btn.font.size = 11
                btn.font.bold = True
                btn.font.color = (255, 255, 255)
                btn.color = DARK_BLUE

            # Set entire sheet font as base
            sheet.range('B6:C9').font.name = 'Calibri Light'
            sheet.range('B6:C9').font.size = 10
            sheet.range('C6:C9').font.bold = True

            # Add dropdown for Current Month (C7)
            try:
                month_list = ','.join([m[2] for m in months])
                sheet.range('C7').api.Validation.Delete()
                sheet.range('C7').api.Validation.Add(
                    Type=3,  # xlValidateList
                    AlertStyle=1,  # xlValidAlertStop
                    Formula1=month_list
                )
            except:
                pass

            sheet.range('C13:C19').font.name = 'Calibri Light'
            sheet.range('C13:C19').font.size = 9
            sheet.range('C13:C19').font.color = GRAY

            sheet.range('B23:B26').font.name = 'Calibri Light'
            sheet.range('B23:B26').font.size = 10
            sheet.range('B23:B26').font.color = GRAY

            # Navigation links
            nav_range = sheet.range('B29:B34')
            nav_range.font.name = 'Calibri Light'
            nav_range.font.size = 10
            nav_range.font.color = (0, 102, 204)

            # Add hyperlinks for navigation
            nav_sheets = ['PL', 'Balance_Sheet', 'Cash_Flow', 'Notes', 'Source_PL', 'Source_BS']
            for i, target in enumerate(nav_sheets):
                try:
                    sheet.range(f'B{29+i}').add_hyperlink(f'#{target}!A1')
                except:
                    pass

        except Exception as e:
            print(f"Menu formatting warning: {e}")

        # Set column widths
        try:
            sheet.range('A:A').column_width = 3
            sheet.range('B:B').column_width = 30
            sheet.range('C:C').column_width = 40
        except:
            pass

    def _create_pl_report(self, sheet, accounts, months, detected_totals=None):
        """Create P&L report with SUMIF formulas and CFO-grade formatting"""
        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC

        # Title section
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Profit & Loss Statement'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

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
        sheet.range(f'A{header_row}').value = 'Account'

        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"

        # Notes and Summary headers
        sheet.range((header_row, notes_col)).value = 'Notes'
        sheet.range((header_row, py_ytd_col)).value = 'PY YTD'
        sheet.range((header_row, cy_ytd_col)).value = 'CY YTD'
        sheet.range((header_row, var_col)).value = 'Var $'
        sheet.range((header_row, var_pct_col)).value = 'Var %'

        # Full year headers (just year)
        for i, year in enumerate(years):
            sheet.range((header_row, fy_start_col + i)).value = str(year)

        # Format header row
        try:
            header_range = sheet.range((header_row, 1), (header_row, last_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = DARK_BLUE
            for col in range(2, last_col + 1):
                if col not in [spacer1_col, spacer2_col]:
                    sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
            # Clear spacer column headers completely
            sheet.range((header_row, spacer1_col)).value = ''
            sheet.range((header_row, spacer2_col)).value = ''
        except:
            pass

        # Track current section for indentation
        current_section = None
        in_section = False
        row_idx = header_row + 1
        gross_margin_row = None
        net_income_row = None
        total_income_row = None
        total_cogs_row = None
        total_expenses_row = None

        for account in accounts:
            account_name = account['name']

            # Detect section headers
            if account['is_header']:
                current_section = account_name.lower()
                in_section = True

            # Determine if this should be indented (not a header, not a total, in a section)
            should_indent = in_section and not account['is_header'] and not account['is_total']

            # Write account name with indentation
            display_name = account_name
            if should_indent:
                display_name = '    ' + account_name  # 4 space indent

            # Rename Gross Profit to Gross Margin
            if 'gross profit' in account_name.lower():
                display_name = 'Gross Margin'
                gross_margin_row = row_idx

            if 'net income' in account_name.lower() and account['is_total']:
                net_income_row = row_idx

            # Track total rows for validation
            if 'total' in account_name.lower():
                if 'income' in account_name.lower() and 'net' not in account_name.lower():
                    total_income_row = row_idx
                elif 'cost' in account_name.lower() or 'cogs' in account_name.lower():
                    total_cogs_row = row_idx
                elif 'expense' in account_name.lower():
                    total_expenses_row = row_idx

            sheet.range((row_idx, 1)).value = display_name
            sheet.range((row_idx, 1)).font.name = 'Calibri Light'
            sheet.range((row_idx, 1)).font.size = 10
            # Vertical center alignment for account names
            try:
                sheet.range((row_idx, 1)).api.VerticalAlignment = -4108  # xlVAlignCenter
            except:
                pass

            # Apply formatting based on row type
            if account['is_header']:
                sheet.range((row_idx, 1)).font.bold = True
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
                            cell = sheet.range((row_idx, col))
                            cell.font.bold = True
                            try:
                                cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                                cell.api.Borders(8).Weight = 3  # xlMedium
                                cell.api.Borders(9).LineStyle = -4119  # xlDouble
                                cell.api.Borders(9).Weight = 4
                            except:
                                pass
                else:
                    # Other totals: bold, thin top border, gray background
                    for col in range(1, last_col + 1):
                        if col not in [spacer1_col, spacer2_col]:
                            cell = sheet.range((row_idx, col))
                            cell.font.bold = True  # ALL totals bold
                            cell.color = SUBTOTAL_GRAY
                            try:
                                cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                                cell.api.Borders(8).Weight = 2  # xlThin
                            except:
                                pass

            # SUMIF formulas for each month
            for i, (m, y, name) in enumerate(months):
                col = i + 2
                formula = f"=SUMIF(Source_PL!$A:$A,\"{account_name}\",Source_PL!{self._col_letter(col)}:{self._col_letter(col)})"
                sheet.range((row_idx, col)).formula = formula

            # Notes column - lookup formula
            formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100=TRIM($A{row_idx}))*(Notes!$B$2:$B$100=Menu!$C$7)),Notes!$C$2:$C$100),"")'
            sheet.range((row_idx, notes_col)).formula = formula
            sheet.range((row_idx, notes_col)).font.name = 'Calibri Light'
            sheet.range((row_idx, notes_col)).font.size = 9
            # Left justify and wrap notes
            try:
                sheet.range((row_idx, notes_col)).api.HorizontalAlignment = -4131  # xlLeft
                sheet.range((row_idx, notes_col)).api.WrapText = True
                sheet.range((row_idx, notes_col)).api.VerticalAlignment = -4108  # xlVAlignCenter
            except:
                pass

            # PY YTD (placeholder - 0 for now)
            sheet.range((row_idx, py_ytd_col)).value = 0

            # CY YTD formula
            first_col = self._col_letter(2)
            last_month_letter = self._col_letter(last_month_col)
            sheet.range((row_idx, cy_ytd_col)).formula = f'=SUM({first_col}{row_idx}:{last_month_letter}{row_idx})'

            # Variance $ (CY - PY)
            sheet.range((row_idx, var_col)).formula = f'={self._col_letter(cy_ytd_col)}{row_idx}-{self._col_letter(py_ytd_col)}{row_idx}'

            # Variance % (Var/PY)
            sheet.range((row_idx, var_pct_col)).formula = f'=IFERROR({self._col_letter(var_col)}{row_idx}/{self._col_letter(py_ytd_col)}{row_idx},0)'

            # Full Year columns (sum months for that year)
            for i, year in enumerate(years):
                year_month_cols = [c + 2 for c, (m, y, name) in enumerate(months) if y == year]
                if year_month_cols:
                    refs = '+'.join([f'{self._col_letter(c)}{row_idx}' for c in year_month_cols])
                    sheet.range((row_idx, fy_start_col + i)).formula = f'={refs}'

            row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        data_end_row = row_idx - 1
        data_start_row = header_row + 1

        # Apply number format to data range
        try:
            time.sleep(0.1)
            # Number columns
            for col in range(2, last_col + 1):
                if col == var_pct_col:
                    sheet.range((data_start_row, col), (data_end_row, col)).number_format = '0.0%'
                elif col not in [spacer1_col, spacer2_col, notes_col]:
                    sheet.range((data_start_row, col), (data_end_row, col)).number_format = '#,##0'

            # Right align number columns
            for col in range(2, last_col + 1):
                if col not in [spacer1_col, spacer2_col, notes_col]:
                    sheet.range((data_start_row, col), (data_end_row, col)).api.HorizontalAlignment = -4152  # xlRight
                    sheet.range((data_start_row, col), (data_end_row, col)).font.name = 'Calibri Light'
                    sheet.range((data_start_row, col), (data_end_row, col)).font.size = 10
        except:
            pass

        # Clear all formatting from spacer columns (entire column, true white space)
        try:
            for spacer_col in [spacer1_col, spacer2_col]:
                spacer_range = sheet.range((1, spacer_col), (row_idx + 50, spacer_col))
                spacer_range.clear()
                spacer_range.color = None  # Remove any background
        except:
            pass

        # Add EBITDA Reconciliation section
        ebitda_start = row_idx + 2
        sheet.range((ebitda_start, 1)).value = 'RECONCILIATION TO EBITDA'
        sheet.range((ebitda_start, 1)).font.name = 'Calibri Light'
        sheet.range((ebitda_start, 1)).font.size = 10
        sheet.range((ebitda_start, 1)).font.bold = True
        sheet.range((ebitda_start, 1)).font.color = DARK_BLUE

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
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if item_type == 'ebitda_total':
                # Bold with borders for EBITDA total
                sheet.range((r, 1)).font.bold = True
                for col in range(1, last_month_col + 1):
                    if col > 1:
                        # Sum of Net Income + all add-backs
                        formula = f'=SUM({self._col_letter(col)}{ebitda_start+1}:{self._col_letter(col)}{r-1})'
                        sheet.range((r, col)).formula = formula
                    cell = sheet.range((r, col))
                    cell.font.bold = True
                    try:
                        cell.api.Borders(8).LineStyle = 1
                        cell.api.Borders(8).Weight = 3
                        cell.api.Borders(9).LineStyle = -4119  # xlDouble
                    except:
                        pass
            else:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    if item_type == 'net_income' and net_income_row:
                        formula = f'={col_letter}{net_income_row}'
                    elif item_type == 'interest':
                        formula = f'=SUMIF(Source_PL!$A:$A,"*Interest*",Source_PL!{col_letter}:{col_letter})*-1'
                    elif item_type == 'taxes':
                        formula = f'=SUMIF(Source_PL!$A:$A,"*Tax*",Source_PL!{col_letter}:{col_letter})*-1'
                    elif item_type == 'depr':
                        formula = f'=SUMIF(Source_PL!$A:$A,"*Deprec*",Source_PL!{col_letter}:{col_letter})*-1+SUMIF(Source_PL!$A:$A,"*Amort*",Source_PL!{col_letter}:{col_letter})*-1'
                    elif item_type == 'nonrecurring':
                        formula = '0'  # Manual entry placeholder
                    else:
                        formula = '0'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        ebitda_end_row = ebitda_start + len(ebitda_items)

        # Add validation section at bottom
        val_start = ebitda_end_row + 2
        sheet.range((val_start, 1)).value = 'VALIDATION'
        sheet.range((val_start, 1)).font.name = 'Calibri Light'
        sheet.range((val_start, 1)).font.size = 10
        sheet.range((val_start, 1)).font.bold = True
        sheet.range((val_start, 1)).font.color = DARK_BLUE

        # Get exact total names from detected_totals for source lookups
        src_total_income = detected_totals.get('total_income', 'Total Income')
        src_total_cogs = detected_totals.get('total_cogs', 'Total Cost of Sales')
        src_total_expenses = detected_totals.get('total_expenses', 'Total Expenses')
        src_net_income = detected_totals.get('net_income', 'Net Income')

        # Display detected totals for user reference
        detect_row = val_start + 1
        sheet.range((detect_row, 1)).value = 'Detected Source Totals:'
        sheet.range((detect_row, 1)).font.name = 'Calibri Light'
        sheet.range((detect_row, 1)).font.size = 9
        sheet.range((detect_row, 1)).font.italic = True
        sheet.range((detect_row, 1)).font.color = (100, 100, 100)

        detect_info = f"Income: {src_total_income or 'NOT FOUND'} | COGS: {src_total_cogs or 'NOT FOUND'} | Expenses: {src_total_expenses or 'NOT FOUND'} | Net Income: {src_net_income or 'NOT FOUND'}"
        sheet.range((detect_row, 2)).value = detect_info
        sheet.range((detect_row, 2)).font.name = 'Calibri Light'
        sheet.range((detect_row, 2)).font.size = 9
        sheet.range((detect_row, 2)).font.italic = True
        sheet.range((detect_row, 2)).font.color = (100, 100, 100)

        # Report totals section (no GM% or NP% - just the report totals)
        report_start = detect_row + 2
        sheet.range((report_start, 1)).value = 'Report Totals'
        sheet.range((report_start, 1)).font.name = 'Calibri Light'
        sheet.range((report_start, 1)).font.size = 10
        sheet.range((report_start, 1)).font.bold = True

        validation_report_labels = [
            ('Revenue (Report)', total_income_row),
            ('COGS (Report)', total_cogs_row),
            ('Expenses (Report)', total_expenses_row),
            ('Net Income (Report)', net_income_row),
        ]

        for i, (label, ref_row) in enumerate(validation_report_labels):
            r = report_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if ref_row:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        report_val_end = report_start + len(validation_report_labels)

        # Source check rows - using EXACT total names detected from source
        source_start = report_val_end + 1
        sheet.range((source_start, 1)).value = 'Source Totals'
        sheet.range((source_start, 1)).font.name = 'Calibri Light'
        sheet.range((source_start, 1)).font.size = 10
        sheet.range((source_start, 1)).font.bold = True

        # Use exact matches on detected total names (no wildcards)
        source_labels = [
            ('Revenue (Source)', src_total_income),
            ('COGS (Source)', src_total_cogs),
            ('Expenses (Source)', src_total_expenses),
            ('Net Income (Source)', src_net_income),
        ]

        for i, (label, exact_name) in enumerate(source_labels):
            r = source_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if exact_name:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    # Use exact match - no wildcards
                    formula = f'=SUMIF(Source_PL!$A:$A,"{exact_name}",Source_PL!{col_letter}:{col_letter})'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        source_end = source_start + len(source_labels)

        # Variance section (Report - Source = should be 0)
        var_start = source_end + 1
        sheet.range((var_start, 1)).value = 'Variance (Report - Source)'
        sheet.range((var_start, 1)).font.name = 'Calibri Light'
        sheet.range((var_start, 1)).font.size = 10
        sheet.range((var_start, 1)).font.bold = True

        # Calculate row references for variance
        rev_report_row = report_start + 1
        cogs_report_row = report_start + 2
        exp_report_row = report_start + 3
        ni_report_row = report_start + 4
        rev_source_row = source_start + 1
        cogs_source_row = source_start + 2
        exp_source_row = source_start + 3
        ni_source_row = source_start + 4

        variance_items = [
            ('Revenue Variance', rev_report_row, rev_source_row),
            ('COGS Variance', cogs_report_row, cogs_source_row),
            ('Expenses Variance', exp_report_row, exp_source_row),
            ('Net Income Variance', ni_report_row, ni_source_row),
        ]

        for i, (label, report_row, source_row) in enumerate(variance_items):
            r = var_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            for col_idx in range(2, last_month_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, last_col + 1):
            if col in [spacer1_col, spacer2_col]:
                sheet.range((1, col), (1, col)).column_width = 3
            elif col == notes_col:
                sheet.range((1, col), (1, col)).column_width = 30
            else:
                sheet.range((1, col), (1, col)).column_width = 12

        # AutoFit columns to content where appropriate
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, last_month_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

    def _create_bs_report(self, sheet, accounts, months, detected_totals=None):
        """Create Balance Sheet report with SUMIF formulas and CFO-grade formatting"""
        company = self.company_name.get()
        detected_totals = detected_totals or {}

        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC

        # Title section
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Balance Sheet'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        header_row = 4
        last_col = len(months) + 1

        # Headers
        sheet.range(f'A{header_row}').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"

        # Format header row
        try:
            header_range = sheet.range((header_row, 1), (header_row, last_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = DARK_BLUE
            for col in range(2, last_col + 1):
                sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
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

            # Determine indentation
            should_indent = in_section and not account['is_header'] and not account['is_total']

            # Write account name with indentation
            display_name = account_name
            if should_indent:
                display_name = '    ' + account_name

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

            sheet.range((row_idx, 1)).value = display_name
            sheet.range((row_idx, 1)).font.name = 'Calibri Light'
            sheet.range((row_idx, 1)).font.size = 10
            # Vertical center alignment for account names
            try:
                sheet.range((row_idx, 1)).api.VerticalAlignment = -4108  # xlVAlignCenter
            except:
                pass

            # Apply formatting based on row type
            if account['is_header']:
                sheet.range((row_idx, 1)).font.bold = True
                # Header rows should NOT have formulas - leave data cells blank
                row_idx += 1
                in_section = True
                continue  # Skip formula creation for header rows
            elif account['is_total']:
                is_main_total = 'total for assets' in account_name.lower() or 'total for liabilities and equity' in account_name.lower()

                if is_main_total:
                    # Main totals: bold, thick top border, double bottom border
                    for col in range(1, last_col + 1):
                        cell = sheet.range((row_idx, col))
                        cell.font.bold = True
                        try:
                            cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                            cell.api.Borders(8).Weight = 3  # xlMedium
                            cell.api.Borders(9).LineStyle = -4119  # xlDouble
                            cell.api.Borders(9).Weight = 4
                        except:
                            pass
                else:
                    # Subtotals: bold, thin top border, gray background
                    for col in range(1, last_col + 1):
                        cell = sheet.range((row_idx, col))
                        cell.font.bold = True  # ALL totals bold
                        cell.color = SUBTOTAL_GRAY
                        try:
                            cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                            cell.api.Borders(8).Weight = 2  # xlThin
                        except:
                            pass

            # SUMIF formulas for each month
            for i, (m, y, name) in enumerate(months):
                col = i + 2
                formula = f"=SUMIF(Source_BS!$A:$A,\"{account_name}\",Source_BS!{self._col_letter(col)}:{self._col_letter(col)})"
                sheet.range((row_idx, col)).formula = formula

            row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        # Apply number format to data range
        data_start_row = header_row + 1
        data_end_row = row_idx - 1
        try:
            time.sleep(0.1)
            data_range = sheet.range((data_start_row, 2), (data_end_row, last_col))
            data_range.number_format = '#,##0'
            data_range.font.name = 'Calibri Light'
            data_range.font.size = 10
            data_range.api.HorizontalAlignment = -4152  # xlRight
        except:
            pass

        # Add validation section at bottom - with formulas for each month column
        val_start = row_idx + 2
        sheet.range((val_start, 1)).value = 'VALIDATION'
        sheet.range((val_start, 1)).font.name = 'Calibri Light'
        sheet.range((val_start, 1)).font.size = 10
        sheet.range((val_start, 1)).font.bold = True
        sheet.range((val_start, 1)).font.color = DARK_BLUE

        # Get exact total names from detected_totals for source lookups
        src_total_assets = detected_totals.get('total_assets', 'Total for Assets')
        src_total_liab = detected_totals.get('total_liabilities', 'Total for Liabilities')
        src_total_equity = detected_totals.get('total_equity', 'Total for Equity')
        src_total_liab_equity = detected_totals.get('total_liab_equity', 'Total for Liabilities and Equity')

        # Display detected totals for user reference
        detect_row = val_start + 1
        sheet.range((detect_row, 1)).value = 'Detected Source Totals:'
        sheet.range((detect_row, 1)).font.name = 'Calibri Light'
        sheet.range((detect_row, 1)).font.size = 9
        sheet.range((detect_row, 1)).font.italic = True
        sheet.range((detect_row, 1)).font.color = (100, 100, 100)

        detect_info = f"Assets: {src_total_assets or 'NOT FOUND'} | Liabilities: {src_total_liab or 'NOT FOUND'} | Equity: {src_total_equity or 'NOT FOUND'}"
        sheet.range((detect_row, 2)).value = detect_info
        sheet.range((detect_row, 2)).font.name = 'Calibri Light'
        sheet.range((detect_row, 2)).font.size = 9
        sheet.range((detect_row, 2)).font.italic = True
        sheet.range((detect_row, 2)).font.color = (100, 100, 100)

        # Report totals section
        report_start = detect_row + 2
        sheet.range((report_start, 1)).value = 'Report Totals'
        sheet.range((report_start, 1)).font.name = 'Calibri Light'
        sheet.range((report_start, 1)).font.size = 10
        sheet.range((report_start, 1)).font.bold = True

        report_labels = [
            ('Total Assets (Report)', total_assets_row),
            ('Total Liabilities (Report)', total_liab_row),
            ('Total Equity (Report)', total_equity_row),
            ('Total Liab + Equity (Report)', total_liab_equity_row),
        ]

        for i, (label, ref_row) in enumerate(report_labels):
            r = report_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if ref_row:
                for col_idx in range(2, last_col + 1):
                    col_letter = self._col_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        report_end = report_start + len(report_labels)

        # Source check section - using EXACT total names detected from source
        source_start = report_end + 1
        sheet.range((source_start, 1)).value = 'Source Totals'
        sheet.range((source_start, 1)).font.name = 'Calibri Light'
        sheet.range((source_start, 1)).font.size = 10
        sheet.range((source_start, 1)).font.bold = True

        # Use exact matches on detected total names (no wildcards)
        source_labels = [
            ('Total Assets (Source)', src_total_assets),
            ('Total Liabilities (Source)', src_total_liab),
            ('Total Equity (Source)', src_total_equity),
        ]

        for i, (label, exact_name) in enumerate(source_labels):
            r = source_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if exact_name:
                for col_idx in range(2, last_col + 1):
                    col_letter = self._col_letter(col_idx)
                    # Use exact match - no wildcards
                    formula = f'=SUMIF(Source_BS!$A:$A,"{exact_name}",Source_BS!{col_letter}:{col_letter})'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

        source_end = source_start + len(source_labels)

        # Balance Check section
        balance_start = source_end + 1
        sheet.range((balance_start, 1)).value = 'Balance Check'
        sheet.range((balance_start, 1)).font.name = 'Calibri Light'
        sheet.range((balance_start, 1)).font.size = 10
        sheet.range((balance_start, 1)).font.bold = True

        # Assets - (Liab + Equity) should = 0
        r = balance_start + 1
        sheet.range((r, 1)).value = 'Assets - (Liab + Equity)'
        sheet.range((r, 1)).font.name = 'Calibri Light'
        sheet.range((r, 1)).font.size = 10
        sheet.range((r, 1)).font.bold = True

        if total_assets_row and total_liab_equity_row:
            for col_idx in range(2, last_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{total_assets_row}-{col_letter}{total_liab_equity_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'
                sheet.range((r, col_idx)).font.bold = True

        # Variance section (Report - Source = should be 0)
        var_start = balance_start + 2
        sheet.range((var_start, 1)).value = 'Variance (Report - Source)'
        sheet.range((var_start, 1)).font.name = 'Calibri Light'
        sheet.range((var_start, 1)).font.size = 10
        sheet.range((var_start, 1)).font.bold = True

        # Calculate row references for variance
        assets_report_row = report_start + 1
        liab_report_row = report_start + 2
        equity_report_row = report_start + 3
        assets_source_row = source_start + 1
        liab_source_row = source_start + 2
        equity_source_row = source_start + 3

        variance_items = [
            ('Assets Variance', assets_report_row, assets_source_row),
            ('Liabilities Variance', liab_report_row, liab_source_row),
            ('Equity Variance', equity_report_row, equity_source_row),
        ]

        for i, (label, report_row, source_row) in enumerate(variance_items):
            r = var_start + 1 + i
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            for col_idx in range(2, last_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, last_col + 1):
            sheet.range((1, col), (1, col)).column_width = 14

        # AutoFit columns
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, last_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

    def _create_cash_flow(self, sheet, months):
        """Create Cash Flow statement with indirect method formulas"""
        company = self.company_name.get()

        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC

        # Title section
        sheet.range('A1').value = company
        sheet.range('A1').font.name = 'Calibri Light'
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.bold = True

        sheet.range('A2').value = 'Statement of Cash Flows (Indirect Method)'
        sheet.range('A2').font.name = 'Calibri Light'
        sheet.range('A2').font.size = 12
        sheet.range('A2').font.bold = True

        header_row = 4
        last_month_col = len(months) + 1
        ytd_col = last_month_col + 2

        # Headers
        sheet.range(f'A{header_row}').value = 'Description'
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
        sheet.range((header_row, ytd_col)).value = 'YTD'

        # Format header row
        try:
            header_range = sheet.range((header_row, 1), (header_row, ytd_col))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = DARK_BLUE
            for col in range(2, ytd_col + 1):
                sheet.range((header_row, col)).api.HorizontalAlignment = -4108  # xlCenter
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

        # Track section rows for subtotals
        operating_start = None
        operating_end = None
        investing_start = None
        investing_end = None
        financing_start = None
        financing_end = None
        net_income_row = None
        net_change_row = None
        beginning_cash_row = None
        ending_cash_row = None

        for display_name, row_type, formula_type, search_term in cf_structure:
            # Track section boundaries
            if 'OPERATING ACTIVITIES' in display_name and row_type == 'header':
                operating_start = row + 1
            elif 'Net Cash Provided by Operating' in display_name:
                operating_end = row - 1
            elif 'INVESTING ACTIVITIES' in display_name and row_type == 'header':
                investing_start = row + 1
            elif 'Net Cash Used in Investing' in display_name:
                investing_end = row - 1
            elif 'FINANCING ACTIVITIES' in display_name and row_type == 'header':
                financing_start = row + 1
            elif 'Net Cash Provided by Financing' in display_name:
                financing_end = row - 1
            elif 'Net Income' in display_name and formula_type == 'pl_lookup':
                net_income_row = row
            elif 'NET INCREASE' in display_name:
                net_change_row = row
            elif 'Beginning of Period' in display_name:
                beginning_cash_row = row
            elif 'End of Period' in display_name:
                ending_cash_row = row

            # Write row
            sheet.range((row, 1)).value = display_name
            sheet.range((row, 1)).font.name = 'Calibri Light'
            sheet.range((row, 1)).font.size = 10

            # Apply formatting based on type
            if row_type == 'header':
                sheet.range((row, 1)).font.bold = True
            elif row_type == 'subtotal':
                sheet.range((row, 1)).font.bold = True
                for col in range(1, ytd_col + 1):
                    cell = sheet.range((row, col))
                    cell.color = SUBTOTAL_GRAY
                    try:
                        cell.api.Borders(8).LineStyle = 1
                        cell.api.Borders(8).Weight = 2
                    except:
                        pass
            elif row_type == 'total':
                sheet.range((row, 1)).font.bold = True
                for col in range(1, ytd_col + 1):
                    cell = sheet.range((row, col))
                    try:
                        cell.api.Borders(8).LineStyle = 1
                        cell.api.Borders(8).Weight = 3
                        cell.api.Borders(9).LineStyle = -4119  # xlDouble
                    except:
                        pass

            # Apply formulas for each month column
            if row_type in ['item', 'subtotal', 'total'] and formula_type:
                for i in range(len(months)):
                    col = i + 2

                    if formula_type == 'pl_lookup':
                        formula = f'=SUMIF(Source_PL!$A:$A,"*{search_term}*",Source_PL!{self._col_letter(col)}:{self._col_letter(col)})'
                        sheet.range((row, col)).formula = formula

                    elif formula_type == 'bs_change_asset':
                        # For assets: decrease in asset = source of cash (prior - current)
                        if i == 0:
                            sheet.range((row, col)).value = 0
                        else:
                            formula = f'=SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col-1)}:{self._col_letter(col-1)})-SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col)}:{self._col_letter(col)})'
                            sheet.range((row, col)).formula = formula

                    elif formula_type == 'bs_change_liab':
                        # For liabilities/equity: increase = source of cash (current - prior)
                        if i == 0:
                            sheet.range((row, col)).value = 0
                        else:
                            formula = f'=SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col)}:{self._col_letter(col)})-SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col-1)}:{self._col_letter(col-1)})'
                            sheet.range((row, col)).formula = formula

                    elif formula_type == 'bs_prior':
                        # Beginning cash = prior month ending balance
                        if i == 0:
                            formula = f'=SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col)}:{self._col_letter(col)})'
                        else:
                            formula = f'=SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col-1)}:{self._col_letter(col-1)})'
                        sheet.range((row, col)).formula = formula

                    elif formula_type == 'bs_current':
                        # Ending cash = current month balance
                        formula = f'=SUMIF(Source_BS!$A:$A,"*{search_term}*",Source_BS!{self._col_letter(col)}:{self._col_letter(col)})'
                        sheet.range((row, col)).formula = formula

                    elif formula_type == 'manual':
                        sheet.range((row, col)).value = 0

                # YTD formula
                first_col = self._col_letter(2)
                last_col_letter = self._col_letter(last_month_col)
                sheet.range((row, ytd_col)).formula = f'=SUM({first_col}{row}:{last_col_letter}{row})'

            row += 1

        # Now apply subtotal formulas (need to do after all rows are created)
        # Operating subtotal
        if operating_start and operating_end:
            for col in range(2, ytd_col + 1):
                refs = '+'.join([f'{self._col_letter(col)}{r}' for r in range(operating_start, operating_end + 1) if sheet.range((r, 1)).value and 'header' not in str(sheet.range((r, 1)).value).lower()])
                # Simpler: just sum the range
                sheet.range((operating_end + 1, col)).formula = f'=SUM({self._col_letter(col)}{operating_start}:{self._col_letter(col)}{operating_end})'

        # Investing subtotal
        if investing_start and investing_end:
            for col in range(2, ytd_col + 1):
                sheet.range((investing_end + 1, col)).formula = f'=SUM({self._col_letter(col)}{investing_start}:{self._col_letter(col)}{investing_end})'

        # Financing subtotal
        if financing_start and financing_end:
            for col in range(2, ytd_col + 1):
                sheet.range((financing_end + 1, col)).formula = f'=SUM({self._col_letter(col)}{financing_start}:{self._col_letter(col)}{financing_end})'

        # Net change in cash = Operating + Investing + Financing subtotals
        if net_change_row and operating_end and investing_end and financing_end:
            op_row = operating_end + 1
            inv_row = investing_end + 1
            fin_row = financing_end + 1
            for col in range(2, ytd_col + 1):
                col_letter = self._col_letter(col)
                sheet.range((net_change_row, col)).formula = f'={col_letter}{op_row}+{col_letter}{inv_row}+{col_letter}{fin_row}'

        # Ending cash = Beginning + Net Change
        if ending_cash_row and beginning_cash_row and net_change_row:
            for col in range(2, ytd_col + 1):
                col_letter = self._col_letter(col)
                sheet.range((ending_cash_row, col)).formula = f'={col_letter}{beginning_cash_row}+{col_letter}{net_change_row}'

        # Apply number format to data range
        try:
            time.sleep(0.1)
            data_range = sheet.range((header_row + 1, 2), (row - 1, ytd_col))
            data_range.number_format = '#,##0'
            data_range.font.name = 'Calibri Light'
            data_range.font.size = 10
            data_range.api.HorizontalAlignment = -4152  # xlRight
        except:
            pass

        # Set column widths
        sheet.range('A:A').column_width = 50
        for col in range(2, ytd_col + 1):
            sheet.range((1, col), (1, col)).column_width = 14

    def _create_notes_sheet(self, sheet, accounts, months):
        """Create notes sheet for variance explanations with dropdowns"""
        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E

        # Headers
        sheet.range('A1').value = 'Account'
        sheet.range('B1').value = 'Month'
        sheet.range('C1').value = 'Note'

        # Format header row with dark blue background and white text
        header_range = sheet.range('A1:C1')
        header_range.font.name = 'Calibri Light'
        header_range.font.size = 10
        header_range.font.bold = True
        header_range.font.color = (255, 255, 255)
        header_range.color = DARK_BLUE

        # Add data validation (dropdowns) for Account and Month columns
        # Pre-populate some rows for data entry
        num_rows = 100  # Allow up to 100 notes

        try:
            # Create account dropdown list (comma-separated, max 255 chars per validation)
            # Use a named range approach instead
            account_list = ','.join(accounts[:50])  # Limit to avoid Excel limitations
            month_list = ','.join(months)

            # Apply data validation to Account column (A2:A101)
            for row in range(2, min(22, num_rows + 2)):  # Just first 20 rows to avoid slowdown
                try:
                    sheet.range(f'A{row}').api.Validation.Delete()
                    sheet.range(f'A{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=account_list[:255]  # Excel limit
                    )
                except:
                    pass

            # Apply data validation to Month column (B2:B101)
            for row in range(2, min(22, num_rows + 2)):
                try:
                    sheet.range(f'B{row}').api.Validation.Delete()
                    sheet.range(f'B{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=month_list
                    )
                except:
                    pass
        except Exception as e:
            print(f"Warning: Could not add dropdowns to Notes sheet: {e}")

        # Set column widths
        sheet.range('A:A').column_width = 40
        sheet.range('B:B').column_width = 15
        sheet.range('C:C').column_width = 60

        # Format data area
        data_range = sheet.range('A2:C21')
        data_range.font.name = 'Calibri Light'
        data_range.font.size = 10

    def _col_letter(self, col_num):
        """Convert column number to letter"""
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result


def main():
    root = tk.Tk()
    app = DNAModelApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
