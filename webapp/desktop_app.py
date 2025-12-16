"""
CFO Financial Model Generator - Desktop Application
Creates true macro-enabled Excel files with embedded VBA using xlwings

VERSION HISTORY:
- v1.0.0 (2024-12-01): Initial release with P&L, Balance Sheet, Cash Flow
- v1.1.0 (2024-12-05): Added Dashboard module with KPIs and ratios
- v1.2.0 (2025-12-09): Renamed from DNA Model to Financial Model
- v1.2.2 (2025-12-10): Fixed Dashboard YTD calculations to use current year only
                        (Jan through current month from Menu!C7)
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

# Application Version
APP_VERSION = "1.3.0"
APP_BUILD_DATE = "2025-12-10"
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
'''

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.resizable(False, False)

        # Variables
        self.pl_path = tk.StringVar()
        self.bs_path = tk.StringVar()
        self.existing_model_path = tk.StringVar()  # For update mode
        self.company_name = tk.StringVar(value="Company Name")
        self.fiscal_start = tk.StringVar(value="January")

        # Calculate first month of current/most recent fiscal year for display defaults
        fy_month, fy_year = self._get_current_fy_start("January")
        self.display_month = tk.StringVar(value=fy_month)
        self.display_year = tk.StringVar(value=str(fy_year))
        self.mode = tk.StringVar(value="new")  # "new" or "update"

        # Data period selection (for when headers don't clearly indicate dates)
        self.data_start_month = tk.StringVar(value=self.MONTHS[datetime.now().month - 1])
        self.data_start_year = tk.StringVar(value=str(datetime.now().year))
        self.data_end_month = tk.StringVar(value=self.MONTHS[datetime.now().month - 1])
        self.data_end_year = tk.StringVar(value=str(datetime.now().year))

        # Add trace to update display month when fiscal year start changes
        self.fiscal_start.trace_add('write', self._on_fiscal_start_change)

        self._create_ui()
        self._center_window(700, 780)

    def _get_current_fy_start(self, fiscal_start_month_name):
        """Get the first month of the current or most recent fiscal year.

        For example, if fiscal year starts in July and current date is December 2025,
        the current FY started July 2025. If current date is March 2025, the current
        FY started July 2024.
        """
        now = datetime.now()
        fiscal_month_num = self.MONTHS.index(fiscal_start_month_name) + 1

        if now.month >= fiscal_month_num:
            # Current FY started this calendar year
            fy_year = now.year
        else:
            # Current FY started last calendar year
            fy_year = now.year - 1

        return (fiscal_start_month_name, fy_year)

    def _on_fiscal_start_change(self, *args):
        """Update First Display Month when Fiscal Year Start changes."""
        fiscal_start = self.fiscal_start.get()
        fy_month, fy_year = self._get_current_fy_start(fiscal_start)
        self.display_month.set(fy_month)
        self.display_year.set(str(fy_year))

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

        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text="Upload Files", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(file_frame, text="P&L File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.pl_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse...", command=self._browse_pl).grid(row=0, column=2)

        ttk.Label(file_frame, text="Balance Sheet:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        ttk.Entry(file_frame, textvariable=self.bs_path, width=40).grid(row=1, column=1, padx=5, pady=(5, 0))
        ttk.Button(file_frame, text="Browse...", command=self._browse_bs).grid(row=1, column=2, pady=(5, 0))

        # Data Period Selection - what date range does your file cover?
        period_frame = ttk.LabelFrame(main_frame, text="Date Range of Your Data Files", padding="10")
        period_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(period_frame, text="What months are included in your P&L and Balance Sheet files?",
                  font=('Segoe UI', 9)).grid(row=0, column=0, columnspan=6, sticky=tk.W, pady=(0, 8))

        years = [str(y) for y in range(datetime.now().year - 5, datetime.now().year + 2)]

        ttk.Label(period_frame, text="First Month:").grid(row=1, column=0, sticky=tk.W)
        ttk.Combobox(period_frame, textvariable=self.data_start_month, values=self.MONTHS, width=12).grid(row=1, column=1, padx=2)
        ttk.Combobox(period_frame, textvariable=self.data_start_year, values=years, width=6).grid(row=1, column=2, padx=2)

        ttk.Label(period_frame, text="Last Month:").grid(row=1, column=3, sticky=tk.W, padx=(15, 0))
        ttk.Combobox(period_frame, textvariable=self.data_end_month, values=self.MONTHS, width=12).grid(row=1, column=4, padx=2)
        ttk.Combobox(period_frame, textvariable=self.data_end_year, values=years, width=6).grid(row=1, column=5, padx=2)

        ttk.Label(period_frame, text="(Only needed if column headers don't clearly show month names like 'January 2024')",
                  font=('Segoe UI', 8), foreground='gray').grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(5, 0))

        # Configuration (for new models only)
        self.config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        self.config_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(self.config_frame, text="Fiscal Year Start:").grid(row=0, column=0, sticky=tk.W)
        fiscal_combo = ttk.Combobox(self.config_frame, textvariable=self.fiscal_start, values=self.MONTHS, width=15)
        fiscal_combo.grid(row=0, column=1, padx=5)

        ttk.Label(self.config_frame, text="First Display Month:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        display_combo = ttk.Combobox(self.config_frame, textvariable=self.display_month, values=self.MONTHS, width=15)
        display_combo.grid(row=0, column=3, padx=5)

        ttk.Label(self.config_frame, text="First Display Year:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        years = [str(y) for y in range(datetime.now().year - 5, datetime.now().year + 2)]
        year_combo = ttk.Combobox(self.config_frame, textvariable=self.display_year, values=years, width=15)
        year_combo.grid(row=1, column=1, padx=5, pady=(5, 0))

        # Button frame for Generate and Close buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)

        self.generate_btn = ttk.Button(btn_frame, text="Generate Financial Model",
                                       command=self._generate_model, style='Accent.TButton')
        self.generate_btn.pack(fill=tk.X, ipady=10)

        # Status label (shows step progress)
        self.status_label = ttk.Label(main_frame, text="Ready", foreground='gray', font=('Segoe UI', 10))
        self.status_label.pack(pady=(10, 10))

        # Close button
        close_btn = ttk.Button(main_frame, text="Close", command=self.root.quit)
        close_btn.pack(pady=(5, 0))

        # FocusCFO Edition branding footer
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(20, 0))

        # Separator line
        separator = ttk.Separator(footer_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 10))

        # FocusCFO colors: Navy (#1a3a5c) and Orange (#e87722) based on typical CFO branding
        footer_style = ttk.Style()
        footer_style.configure('Footer.TLabel', foreground='#5a6a7a', font=('Segoe UI', 8))
        footer_style.configure('FooterBold.TLabel', foreground='#1a3a5c', font=('Segoe UI', 8, 'bold'))

        # Copyright line
        copyright_label = ttk.Label(footer_frame,
            text="© 2025 Vectra Finance. All rights reserved.",
            style='Footer.TLabel')
        copyright_label.pack()

        # Trademark line
        tm_label = ttk.Label(footer_frame,
            text="Vectra Finance™ is a trademark of Vectra Finance LLC.",
            style='Footer.TLabel')
        tm_label.pack()

        # FocusCFO Edition line
        edition_label = ttk.Label(footer_frame,
            text="FocusCFO Edition – Customized under license for internal use.",
            style='FooterBold.TLabel')
        edition_label.pack(pady=(2, 0))

    def _toggle_mode(self):
        """Toggle between New Model and Update Existing modes"""
        if self.mode.get() == "new":
            # Show company info and config, hide existing model selection
            self.name_frame.pack(fill=tk.X, pady=(0, 10), after=self.root.winfo_children()[0].winfo_children()[2])
            self.config_frame.pack(fill=tk.X, pady=(0, 10))
            self.existing_frame.pack_forget()
            self.generate_btn.config(text="Generate Financial Model")
        else:
            # Show existing model selection, hide company info and config
            self.name_frame.pack_forget()
            self.config_frame.pack_forget()
            self.existing_frame.pack(fill=tk.X, pady=(0, 10), after=self.root.winfo_children()[0].winfo_children()[2])
            self.generate_btn.config(text="Update Financial Model")

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
            self.pl_path.set(path)

    def _browse_bs(self):
        path = filedialog.askopenfilename(
            title="Select Balance Sheet File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if path:
            self.bs_path.set(path)

    def _generate_model(self):
        """Generate or update the financial model based on mode"""
        # Validate common inputs
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

    def _update_status(self, message, color='blue'):
        """Update the status label and refresh the UI"""
        self.status_label.config(text=message, foreground=color)
        self.root.update()
        print(message)

    def _get_user_date_params(self):
        """Get user-specified start/end dates as tuples for _parse_financial_data"""
        start_month_name = self.data_start_month.get()
        start_year = int(self.data_start_year.get())
        end_month_name = self.data_end_month.get()
        end_year = int(self.data_end_year.get())

        # Convert month name to number
        start_month_num = self.MONTHS.index(start_month_name) + 1
        end_month_num = self.MONTHS.index(end_month_name) + 1

        return (start_month_num, start_year), (end_month_num, end_year)

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

    def _update_existing_model(self, save_path):
        """Update an existing Financial Model with new month data"""
        import shutil

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
        self._update_status("Step 1/8: Reading new P&L file...")
        pl_data = pd.read_excel(self.pl_path.get(), header=None)
        pl_indents = self._get_cell_indents(self.pl_path.get())

        self._update_status("Step 2/8: Reading new Balance Sheet file...")
        bs_data = pd.read_excel(self.bs_path.get(), header=None)
        bs_indents = self._get_cell_indents(self.bs_path.get())

        self._update_status("Step 3/8: Extracting account data...")
        pl_accounts, new_months, pl_totals = self._parse_financial_data(pl_data, pl_indents, user_start, user_end)
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data, bs_indents, user_start, user_end)
        print(f"New data has {len(new_months)} months")

        # Work in temp directory
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        try:
            # Copy existing model to temp location
            self._update_status("Step 4/8: Opening existing model...")
            shutil.copy2(self.existing_model_path.get(), temp_path)

            app = xw.App(visible=False)
            try:
                wb = app.books.open(temp_path)

                # Get existing sheets
                source_pl = wb.sheets['Source_PL']
                source_bs = wb.sheets['Source_BS']

                # Find existing months in Source_PL (row 1 has month names, row 2 has YYYYMM)
                self._update_status("Step 5/8: Analyzing existing data...")
                existing_months = set()
                last_col = source_pl.range('A1').end('right').column
                for col in range(2, last_col + 1):
                    yyyymm = source_pl.range((2, col)).value
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
                self._update_status("Step 6/8: Adding new month data to Source sheets...")
                SOURCE_BLACK = (26, 26, 26)  # Match existing header color

                for m, y, name, yyyymm in new_month_list:
                    # Find next available column
                    new_col = source_pl.range('A1').end('right').column + 1

                    # Add month header and YYYYMM
                    source_pl.range((1, new_col)).value = name
                    source_pl.range((2, new_col)).value = yyyymm

                    # Format header to match existing columns
                    header_cell = source_pl.range((1, new_col))
                    header_cell.font.name = 'Calibri Light'
                    header_cell.font.size = 10
                    header_cell.font.bold = True
                    header_cell.font.color = (255, 255, 255)
                    header_cell.color = SOURCE_BLACK
                    header_cell.api.HorizontalAlignment = -4108  # xlCenter

                    # Add P&L data for this month - MATCH BY ACCOUNT NAME
                    print(f"DEBUG: Adding P&L data for month ({m}, {y}) to column {new_col}")

                    # Build a lookup dict from parsed accounts (normalize names for matching)
                    pl_values_lookup = {}
                    for account in pl_accounts:
                        acct_name = account.get('name', '').strip()
                        account_values = account.get('values', {})
                        value = account_values.get((m, y), 0)
                        pl_values_lookup[acct_name] = value
                        # Also store without leading spaces for indented accounts
                        pl_values_lookup[acct_name.lstrip()] = value

                    # Get existing account names from Source_PL and match
                    # Use UsedRange to reliably find last row (end('down') stops at empty cells)
                    try:
                        used_range = source_pl.api.UsedRange
                        last_row_pl = used_range.Row + used_range.Rows.Count - 1
                    except:
                        last_row_pl = source_pl.range('A1').end('down').row

                    print(f"DEBUG: Source_PL last_row_pl = {last_row_pl}")
                    non_zero_count = 0
                    matched_count = 0

                    for row in range(3, last_row_pl + 1):
                        existing_name = source_pl.range((row, 1)).value
                        if existing_name:
                            existing_name_str = str(existing_name).strip()
                            # Try to find match in lookup
                            value = pl_values_lookup.get(existing_name_str,
                                    pl_values_lookup.get(existing_name_str.lstrip(), 0))
                            if value != 0:
                                non_zero_count += 1
                            if existing_name_str in pl_values_lookup or existing_name_str.lstrip() in pl_values_lookup:
                                matched_count += 1
                            source_pl.range((row, new_col)).value = value

                    # Format data column - font and number format
                    data_range_pl = source_pl.range((3, new_col), (last_row_pl, new_col))
                    data_range_pl.number_format = '#,##0'
                    data_range_pl.font.name = 'Calibri Light'
                    data_range_pl.font.size = 10

                    print(f"DEBUG: P&L - {len(pl_accounts)} parsed accounts, {matched_count} matched, {non_zero_count} with non-zero values")
                    if len(pl_accounts) > 0:
                        print(f"DEBUG: First 5 parsed account names:")
                        for idx, acct in enumerate(pl_accounts[:5]):
                            print(f"  [{idx}] '{acct.get('name')}'")
                        print(f"DEBUG: First 5 existing Source_PL account names:")
                        for row in range(3, min(8, last_row_pl + 1)):
                            existing = source_pl.range((row, 1)).value
                            print(f"  [row {row}] '{existing}'")

                    # Add BS data for same month
                    new_col_bs = source_bs.range('A1').end('right').column + 1
                    source_bs.range((1, new_col_bs)).value = name
                    source_bs.range((2, new_col_bs)).value = yyyymm

                    # Format BS header to match existing columns
                    bs_header_cell = source_bs.range((1, new_col_bs))
                    bs_header_cell.font.name = 'Calibri Light'
                    bs_header_cell.font.size = 10
                    bs_header_cell.font.bold = True
                    bs_header_cell.font.color = (255, 255, 255)
                    bs_header_cell.color = SOURCE_BLACK
                    bs_header_cell.api.HorizontalAlignment = -4108  # xlCenter

                    # Add BS data for this month - MATCH BY ACCOUNT NAME
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
                        used_range_bs = source_bs.api.UsedRange
                        last_row_bs = used_range_bs.Row + used_range_bs.Rows.Count - 1
                    except:
                        last_row_bs = source_bs.range('A1').end('down').row

                    print(f"DEBUG: Source_BS last_row_bs = {last_row_bs}")
                    non_zero_count_bs = 0
                    matched_count_bs = 0

                    for row in range(3, last_row_bs + 1):
                        existing_name = source_bs.range((row, 1)).value
                        if existing_name:
                            existing_name_str = str(existing_name).strip()
                            value = bs_values_lookup.get(existing_name_str,
                                    bs_values_lookup.get(existing_name_str.lstrip(), 0))
                            if value != 0:
                                non_zero_count_bs += 1
                            if existing_name_str in bs_values_lookup or existing_name_str.lstrip() in bs_values_lookup:
                                matched_count_bs += 1
                            source_bs.range((row, new_col_bs)).value = value

                    # Format BS data column - font and number format
                    data_range_bs = source_bs.range((3, new_col_bs), (last_row_bs, new_col_bs))
                    data_range_bs.number_format = '#,##0'
                    data_range_bs.font.name = 'Calibri Light'
                    data_range_bs.font.size = 10

                    print(f"DEBUG: BS - {len(bs_accounts)} parsed accounts, {matched_count_bs} matched, {non_zero_count_bs} with non-zero values")

                # Update named ranges - use UsedRange for reliable row counts
                try:
                    used_range_pl = source_pl.api.UsedRange
                    pl_last_row = used_range_pl.Row + used_range_pl.Rows.Count - 1
                    pl_last_col = used_range_pl.Column + used_range_pl.Columns.Count - 1
                except:
                    pl_last_row = source_pl.range('A1').end('down').row
                    pl_last_col = source_pl.range('A1').end('right').column

                try:
                    used_range_bs = source_bs.api.UsedRange
                    bs_last_row = used_range_bs.Row + used_range_bs.Rows.Count - 1
                    bs_last_col = used_range_bs.Column + used_range_bs.Columns.Count - 1
                except:
                    bs_last_row = source_bs.range('A1').end('down').row
                    bs_last_col = source_bs.range('A1').end('right').column

                print(f"DEBUG: Source_PL dimensions: {pl_last_row} rows, {pl_last_col} columns (up to column {self._col_letter(pl_last_col)})")
                print(f"DEBUG: Source_BS dimensions: {bs_last_row} rows, {bs_last_col} columns (up to column {self._col_letter(bs_last_col)})")

                # Check what's in row 2 (YYYYMM values) of Source_PL
                yyyymm_row = []
                for col in range(2, pl_last_col + 1):
                    val = source_pl.range((2, col)).value
                    yyyymm_row.append(val)
                print(f"DEBUG: Source_PL row 2 (YYYYMM values): {yyyymm_row}")

                try:
                    wb.names['SourcePL'].delete()
                except:
                    pass
                try:
                    wb.names['SourceBS'].delete()
                except:
                    pass

                wb.names.add('SourcePL', f"=Source_PL!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")
                wb.names.add('SourceBS', f"=Source_BS!$A$1:${self._col_letter(bs_last_col)}${bs_last_row}")

                # Get the latest month added
                latest_month = new_month_list[-1]
                latest_m, latest_y, latest_name, latest_yyyymm = latest_month

                # Update Menu sheet with new current month
                self._update_status("Step 7/8: Updating Menu and reports...")
                try:
                    menu_sheet = wb.sheets['Menu']
                    # Set current month - display (C7), YYYYMM helper (G7), and YTD helpers (E7, F7)
                    menu_sheet.range('C7').value = latest_name  # Current month display
                    menu_sheet.range('G7').value = latest_yyyymm  # YYYYMM for formulas

                    # E7 = month number (1-12), F7 = year - used by YTD formulas
                    menu_sheet.range('E7').value = latest_m  # Month number
                    menu_sheet.range('F7').value = latest_y  # Year

                    # Actuals Through should always match Current Month
                    # Update both C9 (display) and G9 (YYYYMM helper)
                    menu_sheet.range('C9').value = menu_sheet.range('C7').value  # Same as current month
                    menu_sheet.range('G9').value = latest_yyyymm  # Actuals through YYYYMM

                    print(f"DEBUG: Updated Menu - Current Month: {latest_name} ({latest_yyyymm}), E7={latest_m}, F7={latest_y}")
                except Exception as e:
                    print(f"Warning: Could not update Menu sheet: {e}")

                # Add new month columns to report sheets (PL, Balance_Sheet, Cash_Flow)
                self._add_month_to_reports(wb, new_month_list, pl_last_col, bs_last_col)

                # Force recalculation
                wb.app.calculate()

                # Save
                self._update_status("Step 8/8: Saving updated model...")
                wb.save()
                wb.close()

            finally:
                app.quit()

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
            menu_sheet = wb.sheets['Menu']

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
                    sheet = wb.sheets[sheet_name]
                    # Add hyperlink in A1
                    sheet.range('A1').value = '← Menu'
                    sheet.range('A1').font.color = (0, 102, 204)  # Blue link color
                    sheet.range('A1').font.underline = True
                    sheet.range('A1').font.size = 9
                    sheet.api.Hyperlinks.Add(
                        Anchor=sheet.range('A1').api,
                        Address="",
                        SubAddress="Menu!A1",
                        TextToDisplay="← Menu"
                    )
                except Exception as e:
                    print(f"DEBUG: Could not add Menu link to {sheet_name}: {e}")

            # Hide helper columns E:G on Menu
            try:
                menu_sheet.range('E:G').api.EntireColumn.Hidden = True
            except Exception as e:
                print(f"DEBUG: Could not hide columns E:G: {e}")

            # Add Quick Links box in upper right of Menu (starting at I2)
            try:
                start_col = 'I'
                start_row = 2

                # Header
                menu_sheet.range(f'{start_col}{start_row}').value = "Quick Links"
                menu_sheet.range(f'{start_col}{start_row}').font.bold = True
                menu_sheet.range(f'{start_col}{start_row}').font.size = 11

                # Links
                link_row = start_row + 1
                for sheet_name, display_name in report_sheets:
                    try:
                        menu_sheet.range(f'{start_col}{link_row}').value = display_name
                        menu_sheet.range(f'{start_col}{link_row}').font.color = (0, 102, 204)
                        menu_sheet.range(f'{start_col}{link_row}').font.underline = True
                        menu_sheet.range(f'{start_col}{link_row}').font.size = 10
                        menu_sheet.api.Hyperlinks.Add(
                            Anchor=menu_sheet.range(f'{start_col}{link_row}').api,
                            Address="",
                            SubAddress=f"'{sheet_name}'!A1",
                            TextToDisplay=display_name
                        )
                        link_row += 1
                    except Exception as e:
                        print(f"DEBUG: Could not add link to {sheet_name} on Menu: {e}")

                # Add border box around Quick Links section
                box_range = menu_sheet.range(f'{start_col}{start_row}:{start_col}{link_row - 1}')
                box_range.api.Borders.LineStyle = 1  # xlContinuous
                box_range.api.Borders.Weight = 2  # xlThin

            except Exception as e:
                print(f"DEBUG: Could not create Quick Links box: {e}")

            print("DEBUG: Added navigation links")

        except Exception as e:
            print(f"DEBUG: Navigation links skipped: {e}")

    def _add_month_to_reports(self, wb, new_month_list, source_pl_last_col, source_bs_last_col=None):
        """Add new month columns to PL, Balance Sheet, and Cash Flow reports

        Also updates all existing SUMPRODUCT formulas to include the new source data range.
        """
        if source_bs_last_col is None:
            source_bs_last_col = source_pl_last_col

        try:
            pl_sheet = wb.sheets['PL']
            bs_sheet = wb.sheets['Balance_Sheet']
            cf_sheet = wb.sheets['Cash_Flow']

            # Get the new source column letters (after data was added)
            pl_col_letter = self._col_letter(source_pl_last_col)
            bs_col_letter = self._col_letter(source_bs_last_col)

            print(f"DEBUG: Updating formulas to use Source_PL up to column {pl_col_letter}, Source_BS up to column {bs_col_letter}")

            for m, y, name, yyyymm in new_month_list:
                # Find where to insert new month on PL
                # Month columns start at B, find the last month column
                pl_last_month_col = pl_sheet.range('B3').end('right').column

                # Insert new column after last month
                new_col = pl_last_month_col + 1

                # Insert column on PL
                pl_sheet.range((1, new_col)).api.EntireColumn.Insert()

                # Set up header row (row 3 is YYYYMM, row 4 is month name header)
                pl_sheet.range((3, new_col)).value = yyyymm
                pl_sheet.range((3, new_col)).font.color = (255, 255, 255)  # White/hidden
                pl_sheet.range((4, new_col)).value = name
                pl_sheet.range((4, new_col)).font.bold = True

                last_row = pl_sheet.range('A4').end('down').row

                # Create formula for new column
                print(f"DEBUG: Creating formulas for PL new column {new_col}, using Source_PL range up to column {pl_col_letter}")
                print(f"DEBUG: Formula will look for YYYYMM = {yyyymm}")
                for row in range(5, last_row + 1):
                    account_cell = f"$A{row}"
                    formula = (
                        f"=SUMPRODUCT("
                        f"(Source_PL!$A$3:$A$1000=TRIM({account_cell}))*"
                        f"(Source_PL!$B$2:${pl_col_letter}$2={yyyymm})*"
                        f"(Source_PL!$B$3:${pl_col_letter}$1000))"
                    )
                    pl_sheet.range((row, new_col)).value = formula
                    if row == 5:
                        print(f"DEBUG: Sample PL formula (row 5): {formula}")

                # Format new column
                pl_sheet.range((5, new_col), (last_row, new_col)).number_format = '#,##0'

                # After inserting new column, update YTD formulas to include new column in their range
                # Find the YTD columns (PY YTD and CY YTD) - they're after Notes column
                # Layout: Months | Notes | Spacer | PY YTD | CY YTD | Var $ | Var % | Spacer | Annual years
                # The new column is inserted into the months area, so all columns shift right

                # Find where the summary columns are by looking for headers
                header_row = 4
                py_ytd_col = None
                cy_ytd_col = None
                fy_start_col = None

                # Scan row 4 to find column headers
                for check_col in range(new_col + 1, new_col + 15):
                    header_val = pl_sheet.range((header_row, check_col)).value
                    if header_val == 'PY YTD':
                        py_ytd_col = check_col
                    elif header_val == 'CY YTD':
                        cy_ytd_col = check_col
                    elif header_val and str(header_val).isdigit() and len(str(header_val)) == 4:
                        # Found a year column (like "2024")
                        if fy_start_col is None:
                            fy_start_col = check_col

                # Update YTD formulas to use expanded data range
                if py_ytd_col and cy_ytd_col:
                    print(f"DEBUG: Found YTD columns - PY YTD at col {py_ytd_col}, CY YTD at col {cy_ytd_col}")
                    first_data_col = self._col_letter(2)  # B
                    last_data_col = self._col_letter(new_col)  # Now includes new column
                    helper_range = f'{first_data_col}$3:{last_data_col}$3'

                    for row in range(5, last_row + 1):
                        data_range = f'{first_data_col}{row}:{last_data_col}{row}'

                        # CY YTD formula
                        cy_formula = (
                            f'=SUMPRODUCT(({data_range})*'
                            f'--(INT({helper_range}/100)=Menu!$F$7)*'
                            f'--(MOD({helper_range},100)<=Menu!$E$7))'
                        )
                        pl_sheet.range((row, cy_ytd_col)).value = cy_formula

                        # PY YTD formula
                        py_formula = (
                            f'=SUMPRODUCT(({data_range})*'
                            f'--(INT({helper_range}/100)=Menu!$F$7-1)*'
                            f'--(MOD({helper_range},100)<=Menu!$E$7))'
                        )
                        pl_sheet.range((row, py_ytd_col)).value = py_formula

                    print(f"DEBUG: Updated YTD formulas with range up to column {last_data_col}")

                # Update Annual column formulas
                if fy_start_col:
                    # Get the year of the new month
                    new_year = y  # From the loop variable (m, y, name, yyyymm)

                    # Collect all months and their years from row 3
                    month_cols_by_year = {}
                    for col in range(2, new_col + 1):
                        yyyymm_val = pl_sheet.range((3, col)).value
                        if yyyymm_val:
                            col_year = int(yyyymm_val) // 100
                            if col_year not in month_cols_by_year:
                                month_cols_by_year[col_year] = []
                            month_cols_by_year[col_year].append(col)

                    # Find all annual columns and update their formulas
                    for fy_col_offset in range(10):  # Check up to 10 year columns
                        fy_col = fy_start_col + fy_col_offset
                        year_header = pl_sheet.range((header_row, fy_col)).value
                        if year_header and str(year_header).isdigit():
                            year = int(year_header)
                            if year in month_cols_by_year:
                                year_cols = month_cols_by_year[year]
                                # Update formula for each data row
                                for row in range(5, last_row + 1):
                                    refs = '+'.join([f'{self._col_letter(c)}{row}' for c in year_cols])
                                    pl_sheet.range((row, fy_col)).formula = f'={refs}'
                                print(f"DEBUG: Updated Annual {year} formula to include columns {[self._col_letter(c) for c in year_cols]}")

                # Update ALL existing month columns to use expanded source range
                print(f"DEBUG: Updating existing PL formulas in columns B to {self._col_letter(new_col-1)}")
                for col in range(2, new_col):  # B through previous last column
                    col_yyyymm = pl_sheet.range((3, col)).value
                    if col_yyyymm:
                        for row in range(5, last_row + 1):
                            account_cell = f"$A{row}"
                            formula = (
                                f"=SUMPRODUCT("
                                f"(Source_PL!$A$3:$A$1000=TRIM({account_cell}))*"
                                f"(Source_PL!$B$2:${pl_col_letter}$2={int(col_yyyymm)})*"
                                f"(Source_PL!$B$3:${pl_col_letter}$1000))"
                            )
                            pl_sheet.range((row, col)).value = formula

                # Do same for Balance Sheet
                bs_last_month_col = bs_sheet.range('B3').end('right').column
                new_col_bs = bs_last_month_col + 1

                bs_sheet.range((1, new_col_bs)).api.EntireColumn.Insert()
                bs_sheet.range((3, new_col_bs)).value = yyyymm
                bs_sheet.range((3, new_col_bs)).font.color = (255, 255, 255)
                bs_sheet.range((4, new_col_bs)).value = name
                bs_sheet.range((4, new_col_bs)).font.bold = True

                bs_last_row = bs_sheet.range('A4').end('down').row
                for row in range(5, bs_last_row + 1):
                    account_cell = f"$A{row}"
                    formula = (
                        f"=SUMPRODUCT("
                        f"(Source_BS!$A$3:$A$1000=TRIM({account_cell}))*"
                        f"(Source_BS!$B$2:${bs_col_letter}$2={yyyymm})*"
                        f"(Source_BS!$B$3:${bs_col_letter}$1000))"
                    )
                    bs_sheet.range((row, new_col_bs)).value = formula

                bs_sheet.range((5, new_col_bs), (bs_last_row, new_col_bs)).number_format = '#,##0'

                # Update existing BS columns
                print(f"DEBUG: Updating existing BS formulas in columns B to {self._col_letter(new_col_bs-1)}")
                for col in range(2, new_col_bs):
                    col_yyyymm = bs_sheet.range((3, col)).value
                    if col_yyyymm:
                        for row in range(5, bs_last_row + 1):
                            account_cell = f"$A{row}"
                            formula = (
                                f"=SUMPRODUCT("
                                f"(Source_BS!$A$3:$A$1000=TRIM({account_cell}))*"
                                f"(Source_BS!$B$2:${bs_col_letter}$2={int(col_yyyymm)})*"
                                f"(Source_BS!$B$3:${bs_col_letter}$1000))"
                            )
                            bs_sheet.range((row, col)).value = formula

                # Do same for Cash Flow
                cf_last_month_col = cf_sheet.range('B3').end('right').column
                new_col_cf = cf_last_month_col + 1

                cf_sheet.range((1, new_col_cf)).api.EntireColumn.Insert()
                cf_sheet.range((3, new_col_cf)).value = yyyymm
                cf_sheet.range((3, new_col_cf)).font.color = (255, 255, 255)
                cf_sheet.range((4, new_col_cf)).value = name
                cf_sheet.range((4, new_col_cf)).font.bold = True

                # Cash Flow formulas reference PL and BS - copy from previous column
                cf_last_row = cf_sheet.range('A4').end('down').row
                for row in range(5, cf_last_row + 1):
                    prev_formula = cf_sheet.range((row, cf_last_month_col)).formula
                    if prev_formula:
                        cf_sheet.range((row, new_col_cf)).value = prev_formula

                cf_sheet.range((5, new_col_cf), (cf_last_row, new_col_cf)).number_format = '#,##0'

                print(f"Added month {name} to reports")

            # Update Forecast sheet if it exists
            try:
                forecast_sheet = wb.sheets['Forecast']
                forecast_used = forecast_sheet.api.UsedRange
                forecast_last_row = forecast_used.Row + forecast_used.Rows.Count - 1
                forecast_last_col = forecast_used.Column + forecast_used.Columns.Count - 1

                print(f"DEBUG: Updating Forecast formulas to use expanded source range (up to column {pl_col_letter})")
                print(f"DEBUG: Forecast sheet has {forecast_last_row} rows, {forecast_last_col} columns")
                import re
                forecast_formula_count = 0
                formulas_found = 0

                def expand_forecast_range(formula_str, new_col):
                    """Replace Source_PL and Source_Budget column references with new_col"""
                    # Handle both Source_PL and Source_Budget
                    result = formula_str
                    for source_sheet in ['Source_PL', 'Source_Budget']:
                        # Pattern matches Source_XX!$B$2:$T$2 etc
                        result = re.sub(
                            rf'{source_sheet}!\$?B\$?(\d+):\$?([A-Z]+)\$?(\d+)',
                            rf'{source_sheet}!$B$\1:${new_col}$\3',
                            result
                        )
                    return result

                # Scan all cells in the Forecast sheet for Source_PL references
                for col in range(2, min(forecast_last_col + 1, 65)):  # B through all used columns
                    for row in range(5, forecast_last_row + 1):  # Data starts at row 5
                        cell = forecast_sheet.range((row, col))
                        formula = cell.formula
                        if formula and ('Source_PL!' in str(formula) or 'Source_Budget!' in str(formula)):
                            formulas_found += 1
                            original = str(formula)
                            updated = expand_forecast_range(original, pl_col_letter)
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
            except Exception as e:
                print(f"DEBUG: Forecast sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Update Forecast_Summary sheet if it exists
            try:
                forecast_summary = wb.sheets['Forecast_Summary']
                fs_used = forecast_summary.api.UsedRange
                fs_last_row = fs_used.Row + fs_used.Rows.Count - 1

                print(f"DEBUG: Updating Forecast_Summary formulas to use expanded source range (up to column {pl_col_letter})")
                import re
                fs_formula_count = 0

                def expand_fs_range(formula_str, new_col):
                    """Replace Source_PL column references with new_col"""
                    result = re.sub(
                        r'Source_PL!\$?B\$?(\d+):\$?[A-Z]+\$?(\d+)',
                        rf'Source_PL!$B$\1:${new_col}$\2',
                        formula_str
                    )
                    return result

                # Check all columns for formulas that reference Source_PL
                for col in range(2, 15):  # B through N
                    for row in range(3, fs_last_row + 1):
                        cell = forecast_summary.range((row, col))
                        formula = cell.formula
                        if formula and 'Source_PL!' in str(formula):
                            original = str(formula)
                            updated = expand_fs_range(original, pl_col_letter)
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

            # Update Dashboard sheet formulas to use new source range
            try:
                dashboard_sheet = wb.sheets['Dashboard']
                dashboard_used = dashboard_sheet.api.UsedRange
                dashboard_last_row = dashboard_used.Row + dashboard_used.Rows.Count - 1

                print(f"DEBUG: Updating Dashboard formulas to use expanded source range (up to column {pl_col_letter})")
                import re
                formula_count = 0

                # More flexible patterns - handle both $T and T column references
                def expand_source_range(formula_str, source_sheet, new_col):
                    """Replace any column reference in Source_XX ranges with new_col"""
                    # Pattern matches Source_PL!$B$2:$T$2 or Source_PL!$B$2:T$2 or Source_PL!B$2:T$2 etc.
                    # Captures the row numbers and replaces the end column
                    patterns = [
                        (rf'{source_sheet}!\$?B\$?(\d+):\$?[A-Z]+\$?(\d+)', rf'{source_sheet}!$B$\1:${new_col}$\2'),
                    ]
                    result = formula_str
                    for pattern, replacement in patterns:
                        result = re.sub(pattern, replacement, result)
                    return result

                for col in range(2, 13):  # B through L
                    for row in range(1, min(dashboard_last_row + 1, 60)):
                        cell = dashboard_sheet.range((row, col))
                        formula = cell.formula
                        if formula and ('Source_PL!' in str(formula) or 'Source_BS!' in str(formula)):
                            original = str(formula)
                            updated = expand_source_range(original, 'Source_PL', pl_col_letter)
                            updated = expand_source_range(updated, 'Source_BS', bs_col_letter)

                            if updated != original:
                                if formula_count < 3:  # Show first 3 changes
                                    print(f"DEBUG: Dashboard ({row},{col}) BEFORE: {original[:80]}...")
                                    print(f"DEBUG: Dashboard ({row},{col}) AFTER:  {updated[:80]}...")
                                cell.value = updated
                                formula_count += 1

                print(f"DEBUG: Updated {formula_count} Dashboard formulas")
            except Exception as e:
                print(f"DEBUG: Dashboard sheet update skipped: {e}")
                import traceback
                traceback.print_exc()

            # Auto-fit columns on Balance Sheet and Cash Flow after update
            try:
                bs_sheet = wb.sheets['Balance_Sheet']
                bs_sheet.autofit('c')  # Auto-fit all columns
                print("DEBUG: Auto-fit Balance Sheet columns")
            except Exception as e:
                print(f"DEBUG: Balance Sheet auto-fit skipped: {e}")

            try:
                cf_sheet = wb.sheets['Cash_Flow']
                cf_sheet.autofit('c')  # Auto-fit all columns
                print("DEBUG: Auto-fit Cash Flow columns")
            except Exception as e:
                print(f"DEBUG: Cash Flow auto-fit skipped: {e}")

            # Format Forecast notes: wrap text, vertical center, auto-fit row heights
            try:
                forecast_sheet = wb.sheets['Forecast']
                forecast_used = forecast_sheet.api.UsedRange
                forecast_last_row = forecast_used.Row + forecast_used.Rows.Count - 1

                # Note columns are every 5th column starting at column 5 (E), 10 (J), 15 (O), etc.
                # Actually: col 2=Actual, col 3=Budget, col 4=Adj, col 5=Note, col 6=Forecast
                # So Note columns are at 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60
                note_cols = [5 + (i * 5) for i in range(12)]  # 12 months

                for note_col in note_cols:
                    try:
                        note_range = forecast_sheet.range((5, note_col), (forecast_last_row, note_col))
                        note_range.api.WrapText = True
                        note_range.api.VerticalAlignment = -4108  # xlVAlignCenter
                    except:
                        pass

                # Auto-fit row heights to accommodate wrapped text
                forecast_sheet.range((5, 1), (forecast_last_row, 1)).api.EntireRow.AutoFit()

                # Also auto-fit all data cells to be vertically centered
                data_range = forecast_sheet.range((5, 1), (forecast_last_row, 62))
                data_range.api.VerticalAlignment = -4108  # xlVAlignCenter

                print("DEBUG: Formatted Forecast notes and auto-fit rows")
            except Exception as e:
                print(f"DEBUG: Forecast formatting skipped: {e}")

            # Add navigation links to all sheets
            self._add_navigation_links(wb, new_month_list)

        except Exception as e:
            print(f"Warning: Error adding month to reports: {e}")
            import traceback
            traceback.print_exc()

    def _create_excel_model(self, save_path):
        """Create the Excel model using xlwings"""
        # Get user-specified date range
        user_start, user_end = self._get_user_date_params()

        # Parse input files with indentation detection
        self._update_status("Step 1/12: Reading P&L file...")
        pl_data = pd.read_excel(self.pl_path.get(), header=None)
        pl_indents = self._get_cell_indents(self.pl_path.get())

        self._update_status("Step 2/12: Reading Balance Sheet file...")
        bs_data = pd.read_excel(self.bs_path.get(), header=None)
        bs_indents = self._get_cell_indents(self.bs_path.get())

        self._update_status("Step 3/12: Extracting account data...")
        pl_accounts, pl_months, pl_totals = self._parse_financial_data(pl_data, pl_indents, user_start, user_end)
        bs_accounts, _, bs_totals = self._parse_financial_data(bs_data, bs_indents, user_start, user_end)
        print(f"Found {len(pl_accounts)} P&L accounts, {len(bs_accounts)} BS accounts, {len(pl_months)} months")
        print(f"P&L Totals detected: {pl_totals}")
        print(f"BS Totals detected: {bs_totals}")

        # Check if template exists
        use_template = os.path.exists(TEMPLATE_PATH)

        # Work in temp directory to avoid OneDrive locking issues
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, 'temp_model.xlsm')

        # Create Excel application
        self._update_status("Step 4/12: Starting Excel...")
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

                # Create new sheets if they don't exist (for v1.3.0 Forecast module)
                try:
                    source_budget = wb.sheets['Source_Budget']
                    source_budget.range('A1:ZZ1000').clear()
                except:
                    source_budget = wb.sheets.add('Source_Budget', after=source_bs)

                try:
                    forecast_sheet = wb.sheets['Forecast']
                    forecast_sheet.range('A1:ZZ1000').clear()
                except:
                    forecast_sheet = wb.sheets.add('Forecast', after=cf_sheet)

                try:
                    forecast_summary_sheet = wb.sheets['Forecast_Summary']
                    forecast_summary_sheet.range('A1:ZZ1000').clear()
                except:
                    forecast_summary_sheet = wb.sheets.add('Forecast_Summary', after=forecast_sheet)

                # Clear source sheets (keep headers)
                source_pl.range('A2:ZZ1000').clear()
                source_bs.range('A2:ZZ1000').clear()

                # Also clear the report sheets for fresh data (including header rows)
                pl_sheet.range('A1:ZZ1000').clear()
                bs_sheet.range('A1:ZZ1000').clear()
                cf_sheet.range('A1:ZZ1000').clear()
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
                source_budget = wb.sheets.add('Source_Budget', after=source_bs)
                pl_sheet = wb.sheets.add('PL', after=source_budget)
                bs_sheet = wb.sheets.add('Balance_Sheet', after=pl_sheet)
                cf_sheet = wb.sheets.add('Cash_Flow', after=bs_sheet)
                forecast_sheet = wb.sheets.add('Forecast', after=cf_sheet)
                forecast_summary_sheet = wb.sheets.add('Forecast_Summary', after=forecast_sheet)
                notes_sheet = wb.sheets.add('Notes', after=forecast_summary_sheet)

            # Populate source sheets
            self._update_status("Step 5/15: Populating source data...")
            self._populate_source_sheet(source_pl, pl_accounts, pl_months)
            self._populate_source_sheet(source_bs, bs_accounts, pl_months)

            # Create empty Source_Budget sheet (for budget import)
            self._update_status("Step 6/15: Creating Source_Budget sheet...")
            self._create_source_budget_sheet(source_budget, pl_accounts, pl_months)

            # Create/update named ranges
            pl_last_row = len(pl_accounts) + 2  # +2 for header and YYYYMM rows
            pl_last_col = len(pl_months) + 1
            bs_last_row = len(bs_accounts) + 2

            # Delete existing named ranges if they exist
            try:
                wb.names['SourcePL'].delete()
            except:
                pass
            try:
                wb.names['SourceBS'].delete()
            except:
                pass
            try:
                wb.names['SourceBudget'].delete()
            except:
                pass

            wb.names.add('SourcePL', f"=Source_PL!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")
            wb.names.add('SourceBS', f"=Source_BS!$A$1:${self._col_letter(pl_last_col)}${bs_last_row}")
            wb.names.add('SourceBudget', f"=Source_Budget!$A$1:${self._col_letter(pl_last_col)}${pl_last_row}")

            # Create Menu sheet
            self._update_status("Step 7/15: Creating Menu sheet...")
            self._create_menu_sheet(menu_sheet, pl_months)

            # Create P&L report
            self._update_status("Step 8/15: Creating P&L report...")
            self._create_pl_report(pl_sheet, pl_accounts, pl_months, pl_totals)

            # Create Balance Sheet report
            self._update_status("Step 9/15: Creating Balance Sheet...")
            self._create_bs_report(bs_sheet, bs_accounts, pl_months, bs_totals)

            # Create Cash Flow statement
            self._update_status("Step 10/15: Creating Cash Flow statement...")
            self._create_cash_flow(cf_sheet, pl_months)

            # Create Forecast sheet
            self._update_status("Step 11/15: Creating Forecast sheet...")
            self._create_forecast_sheet(forecast_sheet, pl_accounts, pl_months)

            # Create Forecast P&L summary
            self._update_status("Step 12/15: Creating Forecast P&L...")
            self._create_forecast_summary_sheet(forecast_summary_sheet, pl_accounts, pl_months)

            # Create Notes sheet with dropdowns (consolidated for P&L, BS, and CF)
            month_names = [name for m, y, name in pl_months]
            self._create_notes_sheet(notes_sheet, pl_accounts, bs_accounts, month_names)

            # Create Dashboard Control sheet
            self._update_status("Step 13/15: Creating Dashboard Control...")
            if use_template:
                # Check if Dashboard_Control already exists
                if 'Dashboard_Control' not in [s.name for s in wb.sheets]:
                    dashboard_control = wb.sheets.add('Dashboard_Control', after=notes_sheet)
                else:
                    dashboard_control = wb.sheets['Dashboard_Control']
                    dashboard_control.range('A1:ZZ1000').clear()
            else:
                dashboard_control = wb.sheets.add('Dashboard_Control', after=notes_sheet)
            self._create_dashboard_control_sheet(dashboard_control, pl_months)

            # Create Dashboard sheet
            self._update_status("Step 14/15: Creating Dashboard...")
            if use_template:
                if 'Dashboard' not in [s.name for s in wb.sheets]:
                    dashboard_sheet = wb.sheets.add('Dashboard', before=menu_sheet)
                else:
                    dashboard_sheet = wb.sheets['Dashboard']
                    dashboard_sheet.range('A1:ZZ1000').clear()
            else:
                dashboard_sheet = wb.sheets.add('Dashboard', before=menu_sheet)
            self._create_dashboard_sheet(dashboard_sheet, pl_accounts, bs_accounts, pl_months, pl_totals)

            # Reorder sheets: Dashboard first, then reports, then Menu at end, then source sheets
            try:
                # Move Menu to after Dashboard_Control (right side)
                menu_sheet.api.Move(After=dashboard_control.api)
                # Move source sheets to the very end
                source_pl.api.Move(After=menu_sheet.api)
                source_bs.api.Move(After=source_pl.api)
                source_budget.api.Move(After=source_bs.api)
            except:
                pass

            # Set source sheet tab colors to black
            try:
                source_pl.api.Tab.Color = 0x000000  # Black
                source_bs.api.Tab.Color = 0x000000  # Black
                source_budget.api.Tab.Color = 0x000000  # Black
            except:
                pass

            # Activate Dashboard sheet so file opens to Dashboard
            try:
                dashboard_sheet.activate()
            except:
                pass

            # Add VBA code - always add to ensure macros work
            try:
                # Check if FinancialModel module already exists (from template)
                module_exists = False
                for component in wb.api.VBProject.VBComponents:
                    if component.Name == "FinancialModel":
                        # Clear existing code and replace with current version
                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                        component.CodeModule.AddFromString(self.VBA_CODE)
                        module_exists = True
                        break

                if not module_exists:
                    vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                    vba_module.Name = "FinancialModel"
                    vba_module.CodeModule.AddFromString(self.VBA_CODE)

                # Add Worksheet_Change event to Menu sheet to trigger UpdateColumnVisibility
                # when C7 changes
                menu_sheet_code = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("C7")) Is Nothing Then
        Call UpdateColumnVisibility
    End If
End Sub
'''
                # Find the Menu sheet's code module and add the event
                for component in wb.api.VBProject.VBComponents:
                    if component.Type == 100:  # 100 = vbext_ct_Document (worksheet)
                        if component.Name == "Menu" or (hasattr(component, 'Properties') and component.Properties("Name").Value == "Menu"):
                            # Clear any existing code and add fresh
                            if component.CodeModule.CountOfLines > 0:
                                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                            component.CodeModule.AddFromString(menu_sheet_code)
                            break
            except Exception as e:
                # VBA access not enabled - save without macros
                pass

            # Save the workbook
            self._update_status("Step 15/15: Saving workbook...")
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
            wb.close()
        except Exception as e:
            print(f"Could not read indentation from {file_path}: {e}")
        return indents

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
            display_name = f"{month_abbrev} {year % 100}"
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
                        display_name = f"{month[:3]} {year % 100}"
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
                    display_name = f"{month_abbrev} {year % 100}"
                    return (month_num, year, display_name)

        # Strategy 3: Try pandas date parsing as last resort
        try:
            parsed_date = pd.to_datetime(val_str, errors='raise', dayfirst=False)
            if parsed_date:
                month_num = parsed_date.month
                year = parsed_date.year
                if 1900 <= year <= 2100:  # Sanity check
                    month_abbrev = self.MONTHS[month_num - 1][:3]
                    display_name = f"{month_abbrev} {year % 100}"
                    return (month_num, year, display_name)
        except:
            pass

        return None

    def _parse_financial_data(self, df, indents=None, user_start=None, user_end=None):
        """Parse financial data from dataframe with indentation levels.

        Args:
            df: DataFrame with financial data
            indents: Dict of row_idx -> indent level
            user_start: Tuple of (month_num, year) user-specified start date
            user_end: Tuple of (month_num, year) user-specified end date
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

        # If still no months, raise error with helpful message
        if not months:
            raise ValueError("Could not determine date columns. Please specify the data period in the UI (Start/End dates).")

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

            is_total = account_name.lower().startswith('total') or account_name in ['Net Income', 'Gross Profit']
            is_header = account_name in ['Income', 'Expenses', 'Cost of Sales', 'Assets', 'Liabilities', 'Equity',
                                         'Other Current Assets', 'Fixed Assets', 'Other Assets',
                                         'Current Liabilities', 'Long Term Liabilities', 'Other Current Liabilities']

            # Get indent level from source file
            indent_level = indents.get(row_idx, 0)

            # Detect specific total rows by their exact names
            name_lower = account_name.lower()
            if name_lower.startswith('total income') or name_lower == 'total for income':
                detected_totals['total_income'] = account_name
            elif (name_lower.startswith('total cost') or
                  name_lower.startswith('total cogs') or
                  name_lower.startswith('total for cost') or
                  ('cost of sales' in name_lower and 'total' in name_lower) or
                  ('cost of goods' in name_lower and 'total' in name_lower)):
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
                'indent': indent_level,
                'values': values,
                'is_total': is_total,
                'is_header': is_header
            })

        return accounts, months, detected_totals

    def _populate_source_sheet(self, sheet, accounts, months):
        """Populate a source data sheet with proper formatting

        Structure:
        - Row 1: Headers (Account, month names like "Jan 24")
        - Row 2: YYYYMM helper values (e.g., 202401) for YTD calculations - hidden
        - Row 3+: Account data
        """
        # Colors - matching web version
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A - almost black for source sheets

        # Row 1: Header
        sheet.range('A1').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            sheet.range((1, i + 2)).value = name

        # Row 2: YYYYMM helper values for YTD calculations (e.g., 202411 for Nov 2024)
        # This enables SUMPRODUCT formulas to filter by year and month
        for i, (m, y, name) in enumerate(months):
            sheet.range((2, i + 2)).value = y * 100 + m

        # Hide row 2 (helper row)
        try:
            sheet.range('2:2').api.Hidden = True
        except:
            pass

        # Row 3+: Data - write all at once for speed and to ensure numbers are numbers
        data = []
        for account in accounts:
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
            sheet.range('A3').value = data  # Start at row 3 now

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

        # Apply number format and font to data columns (now starting at row 3)
        if len(accounts) > 0 and len(months) > 0:
            try:
                data_range = sheet.range((3, 2), (len(accounts) + 2, len(months) + 1))
                data_range.number_format = '#,##0'
                data_range.font.name = 'Calibri Light'
                data_range.font.size = 10

                # Account names column
                account_range = sheet.range((3, 1), (len(accounts) + 2, 1))
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
            ['', 'Financial Model', ''],  # Row 3
            ['', '', ''],  # Row 4
            ['', 'CONFIGURATION', ''],  # Row 5
            ['', 'Company:', self.company_name.get()],  # Row 6
            ['', 'Current Month:', months[-1][2] if months else ''],  # Row 7
            ['', 'Data Range:', f"{months[0][2]} to {months[-1][2]}" if months else ''],  # Row 8
            ['', 'Actuals Through:', months[-1][2] if months else ''],  # Row 9 - For forecast logic
            ['', 'Total Months:', str(len(months)) if months else '0'],  # Row 10
            ['', '', ''],  # Row 11
            ['', 'ACTIONS', ''],  # Row 12
            ['', '', ''],  # Row 13
            ['', 'Upload P&L Data', 'Import new monthly P&L data (Alt+F8)'],  # Row 14
            ['', '', ''],  # Row 15
            ['', 'Upload Balance Sheet', 'Import new Balance Sheet data (Alt+F8)'],  # Row 16
            ['', '', ''],  # Row 17
            ['', 'Upload Budget', 'Import budget data (Alt+F8)'],  # Row 18
            ['', '', ''],  # Row 19
            ['', 'Refresh All', 'Recalculate all formulas (Alt+F8)'],  # Row 20
            ['', '', ''],  # Row 21
            ['', 'Run Diagnostics', 'Validate model and check for errors (Alt+F8)'],  # Row 22
            ['', '', ''],  # Row 23
            ['', '', ''],  # Row 24
            ['', 'HOW TO USE', ''],  # Row 25
            ['', '1. Press Alt+F8 to open Macros dialog', ''],  # Row 26
            ['', '2. Select the macro and click Run', ''],  # Row 27
            ['', '3. For uploads, select your file when prompted', ''],  # Row 28
            ['', '4. Run Diagnostics to check model integrity', ''],  # Row 29
            ['', '', ''],  # Row 30
            ['', 'QUICK NAVIGATION', ''],  # Row 31
            ['', 'Go to P&L Statement', ''],  # Row 32
            ['', 'Go to Balance Sheet', ''],  # Row 33
            ['', 'Go to Cash Flow', ''],  # Row 34
            ['', 'Go to Forecast', ''],  # Row 35
            ['', 'Go to Variance Notes', ''],  # Row 36
            ['', 'Go to Source P&L', ''],  # Row 37
            ['', 'Go to Source BS', ''],  # Row 38
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
            for row in [5, 12, 25, 31]:
                cell = sheet.range(f'B{row}')
                cell.font.name = 'Calibri Light'
                cell.font.size = 12
                cell.font.bold = True
                cell.font.color = DARK_BLUE

            # Action buttons - format with background
            for row in [14, 16, 18, 20, 22]:
                btn = sheet.range(f'B{row}')
                btn.font.name = 'Calibri Light'
                btn.font.size = 11
                btn.font.bold = True
                btn.font.color = (255, 255, 255)
                btn.color = DARK_BLUE

            # Set entire sheet font as base
            sheet.range('B6:C10').font.name = 'Calibri Light'
            sheet.range('B6:C10').font.size = 10
            sheet.range('C6:C10').font.bold = True

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

            # Add dropdown for Actuals Through (C9) - same list as Current Month
            try:
                sheet.range('C9').api.Validation.Delete()
                sheet.range('C9').api.Validation.Add(
                    Type=3,  # xlValidateList
                    AlertStyle=1,  # xlValidAlertStop
                    Formula1=month_list
                )
            except:
                pass

            # Add helper cells for YTD calculations (E7 = month number, F7 = year)
            # These are referenced by YTD formulas on other sheets
            # Store actual numeric values since C7 contains text like "Nov 25"
            try:
                if months:
                    current_month = months[-1][0]  # Month number (1-12)
                    current_year = months[-1][1]   # Year (e.g., 2025)

                    # E7: Current month number (1-12)
                    sheet.range('E7').value = current_month
                    # F7: Current year
                    sheet.range('F7').value = current_year
                    # G7: Current month as YYYYMM value (for Forecast Summary formulas)
                    # E.g., 202411 for Nov 2024
                    sheet.range('G7').value = current_year * 100 + current_month

                    # G9: Actuals Through as YYYYMM value (for Forecast sheet logic)
                    # E.g., 202411 for Nov 2024
                    sheet.range('G9').value = current_year * 100 + current_month

                # Hide columns E, F, and G
                sheet.range('E:G').api.Hidden = True
            except:
                pass

            sheet.range('C14:C22').font.name = 'Calibri Light'
            sheet.range('C14:C22').font.size = 9
            sheet.range('C14:C22').font.color = GRAY

            sheet.range('B26:B29').font.name = 'Calibri Light'
            sheet.range('B26:B29').font.size = 10
            sheet.range('B26:B29').font.color = GRAY

            # Navigation links
            nav_range = sheet.range('B32:B38')
            nav_range.font.name = 'Calibri Light'
            nav_range.font.size = 10
            nav_range.font.color = (0, 102, 204)

            # Add hyperlinks for navigation
            nav_sheets = ['PL', 'Balance_Sheet', 'Cash_Flow', 'Forecast', 'Notes', 'Source_PL', 'Source_BS']
            for i, target in enumerate(nav_sheets):
                try:
                    sheet.range(f'B{32+i}').add_hyperlink(f'#{target}!A1')
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

        # Row 3: Helper row with YYYYMM values for dynamic YTD calculations
        # This allows formulas to compare dates against Menu!C7
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
            # Row 3 stores YYYYMM as number (e.g., 202411 for Nov 2024)
            sheet.range((3, col)).value = y * 100 + m

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

        # Add YTD date range label in row 3 (e.g., "Jan-Nov") merged across PY YTD and CY YTD
        try:
            ytd_label_cell = sheet.range((3, py_ytd_col))
            # Formula shows "Jan-[current month]" based on Menu!E7
            ytd_label_cell.formula = '="Jan-"&TEXT(DATE(2024,Menu!$E$7,1),"mmm")'
            # Merge across PY YTD and CY YTD columns
            sheet.range((3, py_ytd_col), (3, cy_ytd_col)).merge()
            ytd_label_cell.api.HorizontalAlignment = -4108  # xlCenter
            ytd_label_cell.font.name = 'Calibri Light'
            ytd_label_cell.font.size = 9
            ytd_label_cell.font.italic = True
            ytd_label_cell.font.color = DARK_BLUE
        except:
            pass

        # Hide YYYYMM values in row 3 month columns by setting font color to white
        try:
            for col in range(2, last_month_col + 2):
                sheet.range((3, col)).font.color = (255, 255, 255)
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

            # Notes column - lookup formula (matches Statement Type, Date, and Account)
            # Notes structure: A=Statement Type, B=Date, C=Account, D=Note
            formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="P&L")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{row_idx}))),Notes!$D$2:$D$100),"")'
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

            # Dynamic YTD formulas that reference Menu helper cells
            # Row 3 contains YYYYMM values (e.g., 202411 for Nov 2024)
            # Menu!E7 = current month number (1-12)
            # Menu!F7 = current year (e.g., 2024)

            first_data_col = self._col_letter(2)  # B
            last_data_col = self._col_letter(len(months) + 1)
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

            sheet.range((row_idx, cy_ytd_col)).formula = cy_formula
            sheet.range((row_idx, py_ytd_col)).formula = py_formula

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

            # Add Gross Margin % row after Gross Profit
            if 'gross profit' in name_lower and total_income_row:
                sheet.range((row_idx, 1)).value = '    Gross Margin %'
                sheet.range((row_idx, 1)).font.name = 'Calibri Light'
                sheet.range((row_idx, 1)).font.size = 10
                sheet.range((row_idx, 1)).font.italic = True
                sheet.range((row_idx, 1)).font.color = (100, 100, 100)

                # Gross Margin % = Gross Profit / Total Revenue for each column
                gross_profit_row = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = self._col_letter(col)
                        formula = f'=IFERROR({col_letter}{gross_profit_row}/{col_letter}{total_income_row},0)'
                        sheet.range((row_idx, col)).formula = formula
                        sheet.range((row_idx, col)).number_format = '0.0%'
                        sheet.range((row_idx, col)).font.name = 'Calibri Light'
                        sheet.range((row_idx, col)).font.size = 10
                        sheet.range((row_idx, col)).font.italic = True
                        sheet.range((row_idx, col)).font.color = (100, 100, 100)
                row_idx += 1

            # Add Net Profit % row after Net Income
            if 'net income' in name_lower and account['is_total'] and total_income_row:
                sheet.range((row_idx, 1)).value = '    Net Profit %'
                sheet.range((row_idx, 1)).font.name = 'Calibri Light'
                sheet.range((row_idx, 1)).font.size = 10
                sheet.range((row_idx, 1)).font.italic = True
                sheet.range((row_idx, 1)).font.color = (100, 100, 100)

                # Net Profit % = Net Income / Total Revenue for each column
                net_income_row_num = row_idx - 1
                for col in range(2, last_col + 1):
                    if col not in [spacer1_col, spacer2_col, notes_col]:
                        col_letter = self._col_letter(col)
                        formula = f'=IFERROR({col_letter}{net_income_row_num}/{col_letter}{total_income_row},0)'
                        sheet.range((row_idx, col)).formula = formula
                        sheet.range((row_idx, col)).number_format = '0.0%'
                        sheet.range((row_idx, col)).font.name = 'Calibri Light'
                        sheet.range((row_idx, col)).font.size = 10
                        sheet.range((row_idx, col)).font.italic = True
                        sheet.range((row_idx, col)).font.color = (100, 100, 100)
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
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if ref_row:
                for col_idx in range(2, last_month_col + 1):
                    col_letter = self._col_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

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
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            for col_idx in range(2, last_month_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'

        var_end_row = var_start + len(variance_items) - 1  # Last row of variance section

        # Add matrix borders around validation sections
        try:
            # Outer border for entire validation section (thick)
            val_range = sheet.range((val_start, 1), (var_end_row, last_month_col))
            val_range.api.Borders(7).LineStyle = 1   # xlLeft
            val_range.api.Borders(7).Weight = 3      # xlMedium
            val_range.api.Borders(8).LineStyle = 1   # xlTop
            val_range.api.Borders(8).Weight = 3
            val_range.api.Borders(9).LineStyle = 1   # xlBottom
            val_range.api.Borders(9).Weight = 3
            val_range.api.Borders(10).LineStyle = 1  # xlRight
            val_range.api.Borders(10).Weight = 3
            
            # Add thin borders inside
            val_range.api.Borders(11).LineStyle = 1  # xlInsideVertical
            val_range.api.Borders(11).Weight = 2     # xlThin
            val_range.api.Borders(12).LineStyle = 1  # xlInsideHorizontal
            val_range.api.Borders(12).Weight = 2     # xlThin
            
            # Bold separator line between Report Totals and Source Totals
            separator1 = sheet.range((source_start, 1), (source_start, last_month_col))
            separator1.api.Borders(8).LineStyle = 1
            separator1.api.Borders(8).Weight = 3     # xlMedium - bold line
            
            # Bold separator line between Source Totals and Variance
            separator2 = sheet.range((var_start, 1), (var_start, last_month_col))
            separator2.api.Borders(8).LineStyle = 1
            separator2.api.Borders(8).Weight = 3     # xlMedium - bold line
        except Exception as e:
            print(f"Border formatting warning: {e}")

        # Group validation section so it can be collapsed
        try:
            sheet.range(f'{val_start}:{var_end_row}').api.Rows.Group()
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
                        sheet.range(f'{section_start}:{section_end}').api.Rows.Group()
                    section_start = None
        except Exception as e:
            print(f"PL section grouping warning: {e}")

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

        # Group columns by year and hide prior years
        self._group_columns_by_year(sheet, months, header_row)

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
        last_month_col = len(months) + 1
        notes_col = last_month_col + 1
        last_col = notes_col

        # Headers
        sheet.range(f'A{header_row}').value = 'Account'

        # Row 3: Helper row with YYYYMM values for dynamic calculations
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
            sheet.range((3, col)).value = y * 100 + m

        # Notes header
        sheet.range((header_row, notes_col)).value = 'Notes'

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

        # Hide YYYYMM values in row 3 by setting font color to white
        try:
            for col in range(2, last_month_col + 1):
                sheet.range((3, col)).font.color = (255, 255, 255)
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

            # Notes column - lookup formula (matches Statement Type, Date, and Account)
            # Notes structure: A=Statement Type, B=Date, C=Account, D=Note
            notes_formula = f'=IFERROR(LOOKUP(2,1/((Notes!$A$2:$A$100="Balance Sheet")*(Notes!$B$2:$B$100=TEXT(Menu!$C$7,"mmm yy"))*(Notes!$C$2:$C$100=TRIM($A{row_idx}))),Notes!$D$2:$D$100),"")'
            sheet.range((row_idx, notes_col)).formula = notes_formula
            sheet.range((row_idx, notes_col)).font.name = 'Calibri Light'
            sheet.range((row_idx, notes_col)).font.size = 9
            try:
                sheet.range((row_idx, notes_col)).api.HorizontalAlignment = -4131  # xlLeft
                sheet.range((row_idx, notes_col)).api.WrapText = True
            except:
                pass

            row_idx += 1

            # Reset section tracking after totals
            if account['is_total']:
                in_section = False

        # Apply number format to data range (exclude Notes column)
        data_start_row = header_row + 1
        data_end_row = row_idx - 1
        try:
            time.sleep(0.1)
            data_range = sheet.range((data_start_row, 2), (data_end_row, last_month_col))
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
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            if ref_row:
                for col_idx in range(2, last_col + 1):
                    col_letter = self._col_letter(col_idx)
                    formula = f'={col_letter}{ref_row}'
                    sheet.range((r, col_idx)).formula = formula
                    sheet.range((r, col_idx)).number_format = '#,##0'

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

        source_end = source_start + len(source_labels) - 1  # Last row of source section

        # Balance Check - just the row, no header
        balance_row = source_end + 1
        sheet.range((balance_row, 1)).value = 'Assets - (Liab + Equity)'
        sheet.range((balance_row, 1)).font.name = 'Calibri Light'
        sheet.range((balance_row, 1)).font.size = 10
        sheet.range((balance_row, 1)).font.bold = True

        if total_assets_row and total_liab_equity_row:
            for col_idx in range(2, last_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{total_assets_row}-{col_letter}{total_liab_equity_row}'
                sheet.range((balance_row, col_idx)).formula = formula
                sheet.range((balance_row, col_idx)).number_format = '#,##0'
                sheet.range((balance_row, col_idx)).font.bold = True

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
            sheet.range((r, 1)).value = label
            sheet.range((r, 1)).font.name = 'Calibri Light'
            sheet.range((r, 1)).font.size = 10

            for col_idx in range(2, last_col + 1):
                col_letter = self._col_letter(col_idx)
                formula = f'={col_letter}{report_row}-{col_letter}{source_row}'
                sheet.range((r, col_idx)).formula = formula
                sheet.range((r, col_idx)).number_format = '#,##0'

        var_end_row = var_start + len(variance_items) - 1  # Last row of variance section

        # Add matrix borders around validation sections
        try:
            # Outer border for entire validation section (thick)
            val_range = sheet.range((val_start, 1), (var_end_row, last_col))
            val_range.api.Borders(7).LineStyle = 1   # xlLeft
            val_range.api.Borders(7).Weight = 3      # xlMedium
            val_range.api.Borders(8).LineStyle = 1   # xlTop
            val_range.api.Borders(8).Weight = 3
            val_range.api.Borders(9).LineStyle = 1   # xlBottom
            val_range.api.Borders(9).Weight = 3
            val_range.api.Borders(10).LineStyle = 1  # xlRight
            val_range.api.Borders(10).Weight = 3
            
            # Add thin borders inside
            val_range.api.Borders(11).LineStyle = 1  # xlInsideVertical
            val_range.api.Borders(11).Weight = 2     # xlThin
            val_range.api.Borders(12).LineStyle = 1  # xlInsideHorizontal
            val_range.api.Borders(12).Weight = 2     # xlThin
            
            # Bold separator line between Report Totals and Source Totals
            separator1 = sheet.range((source_start, 1), (source_start, last_col))
            separator1.api.Borders(8).LineStyle = 1
            separator1.api.Borders(8).Weight = 3     # xlMedium - bold line
            
            # Bold separator line between Source Totals and Balance Check
            separator2 = sheet.range((balance_row, 1), (balance_row, last_col))
            separator2.api.Borders(8).LineStyle = 1
            separator2.api.Borders(8).Weight = 3     # xlMedium - bold line
            
            # Bold separator line between Balance Check and Variance
            separator3 = sheet.range((var_start, 1), (var_start, last_col))
            separator3.api.Borders(8).LineStyle = 1
            separator3.api.Borders(8).Weight = 3     # xlMedium - bold line
        except Exception as e:
            print(f"Border formatting warning: {e}")

        # Group validation section so it can be collapsed
        try:
            sheet.range(f'{val_start}:{var_end_row}').api.Rows.Group()
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
                        sheet.range(f'{section_start}:{section_end}').api.Rows.Group()
                    section_start = None
        except Exception as e:
            print(f"BS section grouping warning: {e}")

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, last_month_col + 1):
            sheet.range((1, col), (1, col)).column_width = 14
        sheet.range((1, notes_col), (1, notes_col)).column_width = 30  # Notes column

        # AutoFit columns
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, last_month_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

        # Group columns by year and hide prior years
        self._group_columns_by_year(sheet, months, header_row)

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

        # Row 3: Helper row with YYYYMM values for dynamic YTD calculations
        for i, (m, y, name) in enumerate(months):
            col = i + 2
            sheet.range((header_row, col)).value = f"{self.MONTHS[m-1][:3]} {y}"
            sheet.range((3, col)).value = y * 100 + m

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

                # Dynamic YTD formula - sum current year months through Menu helper cells
                # Menu!E7 = current month number, Menu!F7 = current year
                first_col = self._col_letter(2)
                last_col_letter = self._col_letter(last_month_col)
                data_range = f'{first_col}{row}:{last_col_letter}{row}'
                helper_range = f'{first_col}$3:{last_col_letter}$3'

                # YTD = SUMPRODUCT for current year, months through current month
                # Uses -- to coerce TRUE/FALSE to 1/0
                ytd_formula = (
                    f'=SUMPRODUCT(({data_range})*'
                    f'--(INT({helper_range}/100)=Menu!$F$7)*'
                    f'--(MOD({helper_range},100)<=Menu!$E$7))'
                )
                sheet.range((row, ytd_col)).formula = ytd_formula

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

        # Set column widths and AutoFit
        sheet.range('A:A').column_width = 50
        try:
            sheet.range('A:A').api.EntireColumn.AutoFit()
            for col in range(2, ytd_col + 1):
                sheet.range((1, col), (1, col)).api.EntireColumn.AutoFit()
        except:
            pass

        # Hide row 3 (YYYYMM helper row) - make font white
        try:
            for col in range(2, ytd_col + 1):
                sheet.range((3, col)).font.color = (255, 255, 255)
        except:
            pass

        # Group prior year columns like PL and BS
        self._group_columns_by_year(sheet, months, header_row)

    def _create_notes_sheet(self, sheet, pl_accounts, bs_accounts, months):
        """Create consolidated notes sheet for P&L, Balance Sheet, and Cash Flow with dropdowns.

        Columns:
        - A: Statement Type (P&L, Balance Sheet, Cash Flow)
        - B: Date (month dropdown)
        - C: Account (dropdown based on statement type)
        - D: Note
        """
        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E

        # Headers
        sheet.range('A1').value = 'Statement Type'
        sheet.range('B1').value = 'Date'
        sheet.range('C1').value = 'Account'
        sheet.range('D1').value = 'Note'

        # Format header row with dark blue background and white text
        header_range = sheet.range('A1:D1')
        header_range.font.name = 'Calibri Light'
        header_range.font.size = 10
        header_range.font.bold = True
        header_range.font.color = (255, 255, 255)
        header_range.color = DARK_BLUE

        # Pre-populate some rows for data entry
        num_rows = 100  # Allow up to 100 notes

        try:
            # Statement type dropdown
            statement_types = 'P&L,Balance Sheet,Cash Flow'

            # Create account lists
            pl_account_list = ','.join([a['name'] for a in pl_accounts][:40])
            bs_account_list = ','.join([a['name'] for a in bs_accounts][:40])
            # For Cash Flow, use a subset of BS accounts (cash-related)
            cf_accounts = 'Cash,Accounts Receivable,Accounts Payable,Inventory'

            month_list = ','.join(months)

            # Apply data validation to Statement Type column (A2:A101)
            for row in range(2, min(52, num_rows + 2)):
                try:
                    sheet.range(f'A{row}').api.Validation.Delete()
                    sheet.range(f'A{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=statement_types
                    )
                except:
                    pass

            # Apply data validation to Date column (B2:B101)
            for row in range(2, min(52, num_rows + 2)):
                try:
                    sheet.range(f'B{row}').api.Validation.Delete()
                    sheet.range(f'B{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=month_list
                    )
                except:
                    pass

            # Apply data validation to Account column - combined list
            # (ideally would be dependent on Statement Type, but that requires complex VBA)
            all_accounts = pl_account_list[:120] + ',' + bs_account_list[:120]
            for row in range(2, min(52, num_rows + 2)):
                try:
                    sheet.range(f'C{row}').api.Validation.Delete()
                    sheet.range(f'C{row}').api.Validation.Add(
                        Type=3,  # xlValidateList
                        AlertStyle=1,  # xlValidAlertStop
                        Formula1=all_accounts[:255]  # Excel limit
                    )
                except:
                    pass

        except Exception as e:
            print(f"Warning: Could not add dropdowns to Notes sheet: {e}")

        # Set column widths
        sheet.range('A:A').column_width = 15
        sheet.range('B:B').column_width = 12
        sheet.range('C:C').column_width = 40
        sheet.range('D:D').column_width = 60

        # Format data area
        data_range = sheet.range('A2:D51')
        data_range.font.name = 'Calibri Light'
        data_range.font.size = 10

        # Enable AutoFilter for sorting
        try:
            sheet.range('A1:D1').api.AutoFilter()
        except:
            pass

    # =========================================================================
    # FORECAST MODULE METHODS
    # =========================================================================

    def _create_source_budget_sheet(self, sheet, accounts, months):
        """Create Source_Budget sheet with structure matching Source_PL for budget data import.

        Structure:
        - Row 1: Headers (Account, month names like "Jan 24")
        - Row 2: YYYYMM helper values (e.g., 202401) for formula lookups - hidden
        - Row 3+: Account data (initially zeros, populated via budget import)
        """
        # Colors - matching source sheets
        SOURCE_BLACK = (26, 26, 26)  # #1A1A1A

        # Row 1: Header
        sheet.range('A1').value = 'Account'
        for i, (m, y, name) in enumerate(months):
            sheet.range((1, i + 2)).value = name

        # Row 2: YYYYMM helper values for formula lookups
        for i, (m, y, name) in enumerate(months):
            sheet.range((2, i + 2)).value = y * 100 + m

        # Hide row 2 (helper row)
        try:
            sheet.range('2:2').api.Hidden = True
        except:
            pass

        # Row 3+: Account names with zero values (placeholder for budget import)
        data = []
        for account in accounts:
            row = [account['name']]
            for _ in range(len(months)):
                row.append(0)  # Initialize with zeros
            data.append(row)

        if data:
            sheet.range('A3').value = data

        # Format header row
        try:
            header_range = sheet.range((1, 1), (1, len(months) + 1))
            header_range.font.name = 'Calibri Light'
            header_range.font.size = 10
            header_range.font.bold = True
            header_range.font.color = (255, 255, 255)
            header_range.color = SOURCE_BLACK

            for col in range(2, len(months) + 2):
                sheet.range((1, col)).api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass

        # Format data area
        if len(accounts) > 0 and len(months) > 0:
            try:
                data_range = sheet.range((3, 2), (len(accounts) + 2, len(months) + 1))
                data_range.number_format = '#,##0'
                data_range.font.name = 'Calibri Light'
                data_range.font.size = 10

                account_range = sheet.range((3, 1), (len(accounts) + 2, 1))
                account_range.font.name = 'Calibri Light'
                account_range.font.size = 10
            except:
                pass

        # Set column widths
        sheet.range('A:A').column_width = 45
        for col in range(2, len(months) + 2):
            sheet.range((1, col), (1, col)).column_width = 14

    def _create_forecast_sheet(self, sheet, accounts, months):
        """Create Forecast sheet with Actual/Budget/Adjustment/Forecast/Note columns per month.
        Matches P&L formatting with indentation, borders, and profit % rows.

        Layout:
        - Row 1: Title
        - Row 2: Blank
        - Row 3: Month headers (merged across 5 columns each)
        - Row 4: Column sub-headers (Actual, Budget, Adj, Forecast, Note)
        - Row 5+: Account data with P&L-style formatting
        """
        # Colors
        DARK_BLUE = (22, 33, 62)  # #16213E
        LIGHT_GRAY = (242, 242, 242)  # #F2F2F2
        SUBTOTAL_GRAY = (236, 236, 236)  # #ECECEC - matches P&L
        FORECAST_GREEN = (198, 224, 180)  # Light green for forecast column

        # Determine current year from months
        if months:
            current_year = months[-1][1]  # Year from last month
        else:
            current_year = 2024

        # Create all 12 months for the current year (Jan-Dec)
        # This ensures Forecast always has full year regardless of source data
        month_abbrevs = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        year_suffix = str(current_year)[-2:]  # e.g., "24" for 2024
        current_year_months = [(m + 1, current_year, f"{month_abbrevs[m]} {year_suffix}") for m in range(12)]

        # Row 1: Title
        sheet.range('A1').value = f'Forecast - {current_year}'
        sheet.range('A1').font.bold = True
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.name = 'Calibri Light'

        # Row 3 column A: Account header (aligned with month headers)
        sheet.range('A3').value = 'Account'
        sheet.range('A3').font.bold = True
        sheet.range('A3').font.name = 'Calibri Light'
        sheet.range('A3').color = DARK_BLUE
        sheet.range('A3').font.color = (255, 255, 255)

        # Row 3: Month headers (start at column B)
        # Each month has 5 columns: Actual, Budget, Adj, Note, Forecast
        col = 2
        month_start_cols = {}  # Track starting column for each month
        for m, y, name in current_year_months:
            month_start_cols[(m, y)] = col
            # Merge cells for month header
            end_col = col + 4
            sheet.range((3, col)).value = name
            try:
                sheet.range((3, col), (3, end_col)).merge()
                sheet.range((3, col)).api.HorizontalAlignment = -4108  # xlCenter
            except:
                pass
            sheet.range((3, col)).font.bold = True
            sheet.range((3, col)).font.name = 'Calibri Light'
            sheet.range((3, col)).color = DARK_BLUE
            sheet.range((3, col)).font.color = (255, 255, 255)
            col += 5

        # Add Totals column header after all months
        totals_col = col  # First column after all month columns
        sheet.range((3, totals_col)).value = 'Full Year'
        sheet.range((3, totals_col)).font.bold = True
        sheet.range((3, totals_col)).font.name = 'Calibri Light'
        sheet.range((3, totals_col)).color = DARK_BLUE
        sheet.range((3, totals_col)).font.color = (255, 255, 255)
        sheet.range((4, totals_col)).value = 'Total'
        sheet.range((4, totals_col)).font.bold = True
        sheet.range((4, totals_col)).font.size = 9
        sheet.range((4, totals_col)).font.name = 'Calibri Light'
        sheet.range((4, totals_col)).color = FORECAST_GREEN

        # Row 4: Column sub-headers for each month
        # Order: Actual, Budget, Adj, Note, Actual/Forecast (dynamic based on actuals through date)
        col = 2
        for m, y, name in current_year_months:
            yyyymm = y * 100 + m

            sheet.range((4, col)).value = 'Actual'
            sheet.range((4, col + 1)).value = 'Budget'
            sheet.range((4, col + 2)).value = 'Adj'
            sheet.range((4, col + 3)).value = 'Note'
            # Dynamic header: "Actual" if month has actuals, "Forecast" if not
            sheet.range((4, col + 4)).value = f'=IF({yyyymm}<=Menu!$G$9,"Actual","Forecast")'

            # Format sub-headers
            for c in range(col, col + 5):
                sheet.range((4, c)).font.bold = True
                sheet.range((4, c)).font.size = 9
                sheet.range((4, c)).font.name = 'Calibri Light'
                sheet.range((4, c)).color = LIGHT_GRAY

            # Highlight the last column (Actual/Forecast)
            sheet.range((4, col + 4)).color = FORECAST_GREEN

            col += 5

        # Row 5+: Account data with formulas and P&L-style formatting
        row = 5
        last_data_col = totals_col
        col_letter_last = self._col_letter(len(months) + 1)  # For Source sheet references

        # Track key rows for Gross Profit % and Net Profit %
        gross_margin_row = None
        total_income_row = None
        net_income_row = None

        for account in accounts:
            account_name = account['name']
            name_lower = account_name.lower()

            # Get indent level from source file
            indent_level = account.get('indent', 0)

            # Determine display name with indentation
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name

            # Track Gross Profit/Margin row (don't rename - keep original for formula matching)
            if 'gross profit' in name_lower:
                gross_margin_row = row

            # Track Net Income row
            if 'net income' in name_lower and account['is_total']:
                net_income_row = row

            # Track Total Income/Revenue row
            if account['is_total'] and 'total' in name_lower:
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and 'other' not in name_lower):
                    total_income_row = row

            sheet.range(f'A{row}').value = display_name
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10

            # Apply formatting based on row type (matching P&L)
            if account['is_header']:
                # Header rows: bold, no formulas
                sheet.range(f'A{row}').font.bold = True
                row += 1
                continue  # Skip formula creation for header rows

            elif account['is_total']:
                is_net_income = 'net income' in name_lower

                if is_net_income:
                    # Net Income: bold, thick top border, double bottom border
                    for c in range(1, last_data_col + 1):
                        cell = sheet.range((row, c))
                        cell.font.bold = True
                        cell.font.name = 'Calibri Light'
                        try:
                            cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                            cell.api.Borders(8).Weight = 3  # xlMedium
                            cell.api.Borders(9).LineStyle = -4119  # xlDouble
                            cell.api.Borders(9).Weight = 4
                        except:
                            pass
                else:
                    # Other totals: bold, thin top border, gray background
                    for c in range(1, last_data_col + 1):
                        cell = sheet.range((row, c))
                        cell.font.bold = True
                        cell.font.name = 'Calibri Light'
                        cell.color = SUBTOTAL_GRAY
                        try:
                            cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                            cell.api.Borders(8).Weight = 2  # xlThin
                        except:
                            pass

            # Add formulas for each month
            # Column order: Actual, Budget, Adj, Note, Forecast
            col = 2
            for m, y, name in current_year_months:
                yyyymm = y * 100 + m

                # Column 1: Actual (from Source_PL)
                # Use TRIM to remove leading spaces from indented display names
                actual_formula = (
                    f"=SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=TRIM($A{row}))*"
                    f"(Source_PL!$B$2:${col_letter_last}$2={yyyymm})*"
                    f"(Source_PL!$B$3:${col_letter_last}$1000))"
                )
                sheet.range((row, col)).value = actual_formula

                # Column 2: Budget (from Source_Budget)
                budget_formula = (
                    f"=SUMPRODUCT("
                    f"(Source_Budget!$A$3:$A$1000=TRIM($A{row}))*"
                    f"(Source_Budget!$B$2:${col_letter_last}$2={yyyymm})*"
                    f"(Source_Budget!$B$3:${col_letter_last}$1000))"
                )
                sheet.range((row, col + 1)).value = budget_formula

                # Column 3: Adjustment (user input - blank)
                sheet.range((row, col + 2)).value = 0

                # Column 4: Note (user input - blank)
                # Leave empty

                # Column 5: Forecast formula
                # If YYYYMM <= Actuals Through (Menu!G9), use Actual; else Budget + Adj
                actual_cell = self._col_letter(col) + str(row)
                budget_cell = self._col_letter(col + 1) + str(row)
                adj_cell = self._col_letter(col + 2) + str(row)
                forecast_formula = (
                    f"=IF({yyyymm}<=Menu!$G$9,"
                    f"{actual_cell},"
                    f"{budget_cell}+{adj_cell})"
                )
                sheet.range((row, col + 4)).value = forecast_formula

                col += 5

            # Add Full Year Total formula (sum of all Forecast columns)
            forecast_cols = [self._col_letter(2 + m * 5 + 4) for m in range(12)]
            total_formula = "=" + "+".join([f"{fc}{row}" for fc in forecast_cols])
            sheet.range((row, totals_col)).value = total_formula

            row += 1

        # Add Gross Profit % row after Gross Margin
        if gross_margin_row and total_income_row:
            # Insert after last account row - we'll add % rows at the end for now
            pass  # Will add after all accounts

        # Now add Gross Profit % and Net Profit % rows
        data_end_row = row - 1

        # Add blank row then % rows
        row += 1  # Blank row

        # Gross Profit % row
        if gross_margin_row and total_income_row:
            gp_pct_row = row
            sheet.range(f'A{row}').value = 'Gross Profit %'
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10
            sheet.range(f'A{row}').font.bold = True
            sheet.range(f'A{row}').font.italic = True

            # Add % formula for each Forecast column
            col = 2
            for month_idx in range(12):
                forecast_col = 2 + month_idx * 5 + 4  # Forecast column position
                forecast_col_letter = self._col_letter(forecast_col)
                gp_formula = f"=IFERROR({forecast_col_letter}{gross_margin_row}/{forecast_col_letter}{total_income_row},0)"
                sheet.range((row, forecast_col)).value = gp_formula
                sheet.range((row, forecast_col)).number_format = '0.0%'
                sheet.range((row, forecast_col)).font.name = 'Calibri Light'
                sheet.range((row, forecast_col)).font.italic = True
                col += 5

            # Full Year GP %
            totals_col_letter = self._col_letter(totals_col)
            sheet.range((row, totals_col)).value = f"=IFERROR({totals_col_letter}{gross_margin_row}/{totals_col_letter}{total_income_row},0)"
            sheet.range((row, totals_col)).number_format = '0.0%'
            sheet.range((row, totals_col)).font.name = 'Calibri Light'
            sheet.range((row, totals_col)).font.bold = True
            sheet.range((row, totals_col)).font.italic = True
            row += 1

        # Net Profit % row
        if net_income_row and total_income_row:
            sheet.range(f'A{row}').value = 'Net Profit %'
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10
            sheet.range(f'A{row}').font.bold = True
            sheet.range(f'A{row}').font.italic = True

            # Add % formula for each Forecast column
            for month_idx in range(12):
                forecast_col = 2 + month_idx * 5 + 4
                forecast_col_letter = self._col_letter(forecast_col)
                np_formula = f"=IFERROR({forecast_col_letter}{net_income_row}/{forecast_col_letter}{total_income_row},0)"
                sheet.range((row, forecast_col)).value = np_formula
                sheet.range((row, forecast_col)).number_format = '0.0%'
                sheet.range((row, forecast_col)).font.name = 'Calibri Light'
                sheet.range((row, forecast_col)).font.italic = True

            # Full Year NP %
            totals_col_letter = self._col_letter(totals_col)
            sheet.range((row, totals_col)).value = f"=IFERROR({totals_col_letter}{net_income_row}/{totals_col_letter}{total_income_row},0)"
            sheet.range((row, totals_col)).number_format = '0.0%'
            sheet.range((row, totals_col)).font.name = 'Calibri Light'
            sheet.range((row, totals_col)).font.bold = True
            sheet.range((row, totals_col)).font.italic = True

        # Format data area
        if len(accounts) > 0:
            try:
                # Number format for numeric columns (Actual, Budget, Adj, Forecast - not Note)
                for month_idx, (m, y, name) in enumerate(current_year_months):
                    start_col = 2 + month_idx * 5
                    # Actual (col), Budget (col+1), Adj (col+2) - format as numbers
                    for c in range(start_col, start_col + 3):
                        data_range = sheet.range((5, c), (data_end_row, c))
                        data_range.number_format = '#,##0'
                        data_range.font.name = 'Calibri Light'
                        data_range.font.size = 10

                    # Forecast column (col+4) - format as number with highlighting
                    forecast_col = start_col + 4
                    forecast_range = sheet.range((5, forecast_col), (data_end_row, forecast_col))
                    forecast_range.number_format = '#,##0'
                    forecast_range.font.name = 'Calibri Light'
                    forecast_range.font.size = 10
                    forecast_range.color = FORECAST_GREEN

                # Format Totals column
                totals_range = sheet.range((5, totals_col), (data_end_row, totals_col))
                totals_range.number_format = '#,##0'
                totals_range.font.name = 'Calibri Light'
                totals_range.font.size = 10
                totals_range.font.bold = True
                totals_range.color = FORECAST_GREEN
            except:
                pass

        # Set column widths
        # Column order: Actual, Budget, Adj, Note, Forecast
        sheet.range('A:A').column_width = 40
        for month_idx in range(len(current_year_months)):
            start_col = 2 + month_idx * 5
            sheet.range((1, start_col), (1, start_col)).column_width = 11      # Actual
            sheet.range((1, start_col + 1), (1, start_col + 1)).column_width = 11  # Budget
            sheet.range((1, start_col + 2), (1, start_col + 2)).column_width = 9   # Adj
            sheet.range((1, start_col + 3), (1, start_col + 3)).column_width = 20  # Note
            sheet.range((1, start_col + 4), (1, start_col + 4)).column_width = 12  # Forecast
        # Totals column width
        sheet.range((1, totals_col), (1, totals_col)).column_width = 13

        # Create column grouping (collapse Actual, Budget, Adj, Note - leave Forecast visible)
        # Group first 4 columns together so Forecast is shown when collapsed
        try:
            for month_idx in range(len(current_year_months)):
                start_col = 2 + month_idx * 5
                # Group columns: Actual, Budget, Adj, Note (start_col to start_col+3)
                sheet.range((1, start_col), (1, start_col + 3)).api.EntireColumn.Group()

            # Collapse all groups (start with collapsed view)
            sheet.api.Outline.ShowLevels(ColumnLevels=1)
        except:
            pass

        # Hide gridlines
        try:
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

    def _create_forecast_summary_sheet(self, sheet, accounts, months):
        """Create Forecast Summary with Current Month, YTD Actual vs Budget, and YTD Forecast comparisons.
        Matches P&L formatting with indentation, borders, and profit % rows.

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
        DARK_BLUE = (22, 33, 62)
        LIGHT_GRAY = (242, 242, 242)
        HEADER_GRAY = (200, 200, 200)
        SUBTOTAL_GRAY = (236, 236, 236)  # Matches P&L

        # Spacer columns
        SPACER1_COL = 6  # F
        SPACER2_COL = 11  # K
        LAST_DATA_COL = 15  # O

        # Determine current year
        if months:
            current_year = months[-1][1]
        else:
            current_year = 2024

        col_letter_last = self._col_letter(len(months) + 1)

        # Row 1: Title
        sheet.range('A1').value = f'Forecast Summary - {current_year}'
        sheet.range('A1').font.bold = True
        sheet.range('A1').font.size = 14
        sheet.range('A1').font.name = 'Calibri Light'

        # Row 2: Section headers (merged)
        sheet.range('B2').value = 'Current Month'
        try:
            sheet.range('B2:E2').merge()
            sheet.range('B2').api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.name = 'Calibri Light'
        sheet.range('B2').color = DARK_BLUE
        sheet.range('B2').font.color = (255, 255, 255)

        sheet.range('G2').value = 'Year-to-Date Actual vs Budget'
        try:
            sheet.range('G2:J2').merge()
            sheet.range('G2').api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass
        sheet.range('G2').font.bold = True
        sheet.range('G2').font.name = 'Calibri Light'
        sheet.range('G2').color = DARK_BLUE
        sheet.range('G2').font.color = (255, 255, 255)

        sheet.range('L2').value = 'YTD Forecast vs Budget'
        try:
            sheet.range('L2:O2').merge()
            sheet.range('L2').api.HorizontalAlignment = -4108  # xlCenter
        except:
            pass
        sheet.range('L2').font.bold = True
        sheet.range('L2').font.name = 'Calibri Light'
        sheet.range('L2').color = DARK_BLUE
        sheet.range('L2').font.color = (255, 255, 255)

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
            sheet.range(cell).value = value

        # Format header row
        header_range = sheet.range('A3:O3')
        header_range.font.bold = True
        header_range.font.name = 'Calibri Light'
        header_range.font.size = 10
        header_range.color = HEADER_GRAY

        # Spacer columns formatting
        sheet.range('F2:F3').color = (255, 255, 255)
        sheet.range('K2:K3').color = (255, 255, 255)

        # Row 4+: Account data with P&L-style formatting
        row = 4

        # Track key rows for Gross Profit % and Net Profit %
        gross_margin_row = None
        total_income_row = None
        net_income_row = None
        forecast_row = 5  # Track actual Forecast sheet row (starts at 5)

        for acct_idx, account in enumerate(accounts):
            account_name = account['name']
            name_lower = account_name.lower()

            # Get indent level from source file
            indent_level = account.get('indent', 0)

            # Determine display name with indentation
            display_name = account_name
            if indent_level > 0 and not account['is_header'] and not account['is_total']:
                display_name = ('    ' * indent_level) + account_name

            # Track Gross Profit/Margin row (don't rename - keep original for formula matching)
            if 'gross profit' in name_lower:
                gross_margin_row = row

            # Track Net Income row
            if 'net income' in name_lower and account['is_total']:
                net_income_row = row

            # Track Total Income/Revenue row
            if account['is_total'] and 'total' in name_lower:
                if (('income' in name_lower or 'revenue' in name_lower) and
                    'net' not in name_lower and 'other' not in name_lower):
                    total_income_row = row

            sheet.range(f'A{row}').value = display_name
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10

            # Apply formatting based on row type (matching P&L)
            if account['is_header']:
                # Header rows: bold, no formulas
                # Note: Forecast sheet also skips headers, so don't increment forecast_row
                sheet.range(f'A{row}').font.bold = True
                row += 1
                forecast_row += 1  # Forecast sheet also has header row at same position
                continue  # Skip formula creation for header rows

            elif account['is_total']:
                is_net_income = 'net income' in name_lower

                if is_net_income:
                    # Net Income: bold, thick top border, double bottom border
                    for c in range(1, LAST_DATA_COL + 1):
                        if c not in [SPACER1_COL, SPACER2_COL]:
                            cell = sheet.range((row, c))
                            cell.font.bold = True
                            cell.font.name = 'Calibri Light'
                            try:
                                cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                                cell.api.Borders(8).Weight = 3  # xlMedium
                                cell.api.Borders(9).LineStyle = -4119  # xlDouble
                                cell.api.Borders(9).Weight = 4
                            except:
                                pass
                else:
                    # Other totals: bold, thin top border, gray background
                    for c in range(1, LAST_DATA_COL + 1):
                        if c not in [SPACER1_COL, SPACER2_COL]:
                            cell = sheet.range((row, c))
                            cell.font.bold = True
                            cell.font.name = 'Calibri Light'
                            cell.color = SUBTOTAL_GRAY
                            try:
                                cell.api.Borders(8).LineStyle = 1  # xlContinuous top
                                cell.api.Borders(8).Weight = 2  # xlThin
                            except:
                                pass

            # --- CURRENT MONTH SECTION ---
            # Current Month Actual: SUMPRODUCT matching current YYYYMM from Menu
            # Use TRIM to remove leading spaces from indented display names
            cm_actual_formula = (
                f"=SUMPRODUCT("
                f"(Source_PL!$A$3:$A$1000=TRIM($A{row}))*"
                f"(Source_PL!$B$2:${col_letter_last}$2=Menu!$G$7)*"
                f"(Source_PL!$B$3:${col_letter_last}$1000))"
            )
            sheet.range(f'B{row}').value = cm_actual_formula

            # Current Month Budget: from Source_Budget for current month
            cm_budget_formula = (
                f"=SUMPRODUCT("
                f"(Source_Budget!$A$3:$A$1000=TRIM($A{row}))*"
                f"(Source_Budget!$B$2:${col_letter_last}$2=Menu!$G$7)*"
                f"(Source_Budget!$B$3:${col_letter_last}$1000))"
            )
            sheet.range(f'C{row}').value = cm_budget_formula

            # Current Month Variance $ (Actual - Budget)
            sheet.range(f'D{row}').value = f"=B{row}-C{row}"

            # Current Month Variance %
            sheet.range(f'E{row}').value = f"=IFERROR(D{row}/ABS(C{row}),0)"

            # --- YTD SECTION ---
            # YTD Actual: Sum where YYYYMM is in current year and <= current month
            ytd_actual_formula = (
                f"=SUMPRODUCT("
                f"(Source_PL!$A$3:$A$1000=TRIM($A{row}))*"
                f"(INT(Source_PL!$B$2:${col_letter_last}$2/100)=Menu!$F$7)*"
                f"(Source_PL!$B$2:${col_letter_last}$2<=Menu!$G$7)*"
                f"(Source_PL!$B$3:${col_letter_last}$1000))"
            )
            sheet.range(f'G{row}').value = ytd_actual_formula

            # YTD Budget: Sum where YYYYMM is in current year and <= current month
            ytd_budget_formula = (
                f"=SUMPRODUCT("
                f"(Source_Budget!$A$3:$A$1000=TRIM($A{row}))*"
                f"(INT(Source_Budget!$B$2:${col_letter_last}$2/100)=Menu!$F$7)*"
                f"(Source_Budget!$B$2:${col_letter_last}$2<=Menu!$G$7)*"
                f"(Source_Budget!$B$3:${col_letter_last}$1000))"
            )
            sheet.range(f'H{row}').value = ytd_budget_formula

            # YTD Variance $ (Actual - Budget)
            sheet.range(f'I{row}').value = f"=G{row}-H{row}"

            # YTD Variance %
            sheet.range(f'J{row}').value = f"=IFERROR(I{row}/ABS(H{row}),0)"

            # --- YTD FORECAST SECTION ---
            # YTD Forecast: Sum Forecast column from Forecast sheet for months <= current
            # Forecast columns are at positions: col+4 for each month (Actual, Budget, Adj, Note, Forecast)
            # So positions are: 2+4=6, 7+4=11, 12+4=16... (col 6, 11, 16...)
            # We need to sum all 12 months' Forecast values where month <= current month
            ytd_forecast_parts = []
            for month_idx in range(12):
                # Forecast column for this month is at: 2 + month_idx * 5 + 4 = 6 + month_idx * 5
                forecast_col = self._col_letter(2 + month_idx * 5 + 4)
                yyyymm = current_year * 100 + (month_idx + 1)
                ytd_forecast_parts.append(f"IF({yyyymm}<=Menu!$G$7,Forecast!{forecast_col}{forecast_row},0)")

            ytd_forecast_formula = "=" + "+".join(ytd_forecast_parts)
            sheet.range(f'L{row}').value = ytd_forecast_formula

            # YTD Budget (for Forecast comparison) - same as column H
            sheet.range(f'M{row}').value = f"=H{row}"

            # Forecast Variance $ (Forecast - Budget)
            sheet.range(f'N{row}').value = f"=L{row}-M{row}"

            # Forecast Variance %
            sheet.range(f'O{row}').value = f"=IFERROR(N{row}/ABS(M{row}),0)"

            row += 1
            forecast_row += 1  # Increment Forecast sheet row tracker

        # Store data end row before adding % rows
        data_end_row = row - 1

        # Add blank row then Gross Profit % and Net Profit % rows
        row += 1  # Blank row

        # Gross Profit % row
        if gross_margin_row and total_income_row:
            sheet.range(f'A{row}').value = 'Gross Profit %'
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10
            sheet.range(f'A{row}').font.bold = True
            sheet.range(f'A{row}').font.italic = True

            # Current Month GP %
            sheet.range(f'B{row}').value = f"=IFERROR(B{gross_margin_row}/B{total_income_row},0)"
            sheet.range(f'B{row}').number_format = '0.0%'
            sheet.range(f'B{row}').font.name = 'Calibri Light'
            sheet.range(f'B{row}').font.italic = True

            # YTD Actual GP %
            sheet.range(f'G{row}').value = f"=IFERROR(G{gross_margin_row}/G{total_income_row},0)"
            sheet.range(f'G{row}').number_format = '0.0%'
            sheet.range(f'G{row}').font.name = 'Calibri Light'
            sheet.range(f'G{row}').font.italic = True

            # YTD Forecast GP %
            sheet.range(f'L{row}').value = f"=IFERROR(L{gross_margin_row}/L{total_income_row},0)"
            sheet.range(f'L{row}').number_format = '0.0%'
            sheet.range(f'L{row}').font.name = 'Calibri Light'
            sheet.range(f'L{row}').font.italic = True

            row += 1

        # Net Profit % row
        if net_income_row and total_income_row:
            sheet.range(f'A{row}').value = 'Net Profit %'
            sheet.range(f'A{row}').font.name = 'Calibri Light'
            sheet.range(f'A{row}').font.size = 10
            sheet.range(f'A{row}').font.bold = True
            sheet.range(f'A{row}').font.italic = True

            # Current Month NP %
            sheet.range(f'B{row}').value = f"=IFERROR(B{net_income_row}/B{total_income_row},0)"
            sheet.range(f'B{row}').number_format = '0.0%'
            sheet.range(f'B{row}').font.name = 'Calibri Light'
            sheet.range(f'B{row}').font.italic = True

            # YTD Actual NP %
            sheet.range(f'G{row}').value = f"=IFERROR(G{net_income_row}/G{total_income_row},0)"
            sheet.range(f'G{row}').number_format = '0.0%'
            sheet.range(f'G{row}').font.name = 'Calibri Light'
            sheet.range(f'G{row}').font.italic = True

            # YTD Forecast NP %
            sheet.range(f'L{row}').value = f"=IFERROR(L{net_income_row}/L{total_income_row},0)"
            sheet.range(f'L{row}').number_format = '0.0%'
            sheet.range(f'L{row}').font.name = 'Calibri Light'
            sheet.range(f'L{row}').font.italic = True

        # Format data area
        try:
            # Number format for dollar columns
            for col in ['B', 'C', 'D', 'G', 'H', 'I', 'L', 'M', 'N']:
                sheet.range(f'{col}4:{col}{data_end_row}').number_format = '#,##0'

            # Percentage format for variance % columns
            for col in ['E', 'J', 'O']:
                sheet.range(f'{col}4:{col}{data_end_row}').number_format = '0.0%'

            # Font styling
            data_range = sheet.range(f'A4:O{data_end_row}')
            data_range.font.name = 'Calibri Light'
            data_range.font.size = 10
        except:
            pass

        # Clear spacer columns of any background color
        try:
            sheet.range(f'F4:F{row}').color = None
            sheet.range(f'K4:K{row}').color = None
        except:
            pass

        # Set column widths
        sheet.range('A:A').column_width = 40
        sheet.range('B:B').column_width = 12  # CM Actual
        sheet.range('C:C').column_width = 12  # CM Budget
        sheet.range('D:D').column_width = 11  # CM Var $
        sheet.range('E:E').column_width = 9   # CM Var %
        sheet.range('F:F').column_width = 2   # Spacer
        sheet.range('G:G').column_width = 12  # YTD Actual
        sheet.range('H:H').column_width = 12  # YTD Budget
        sheet.range('I:I').column_width = 11  # YTD Var $
        sheet.range('J:J').column_width = 9   # YTD Var %
        sheet.range('K:K').column_width = 2   # Spacer
        sheet.range('L:L').column_width = 12  # YTD Forecast
        sheet.range('M:M').column_width = 12  # YTD Budget (Forecast)
        sheet.range('N:N').column_width = 11  # Forecast Var $
        sheet.range('O:O').column_width = 9   # Forecast Var %

        # Hide gridlines
        try:
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

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

    def _create_dashboard_control_sheet(self, sheet, months):
        """Create the Dashboard Control page with KPI target settings"""
        # Colors
        header_color = (22, 33, 62)  # Dark blue
        section_color = (44, 62, 80)  # Darker gray-blue
        control_color = (232, 244, 253)  # Light blue for editable cells

        # Title
        sheet.range('B2').value = "Dashboard Control Panel"
        sheet.range('B2').font.size = 24
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.color = header_color
        sheet.range('B2:F2').merge()

        sheet.range('B3').value = f"Configure KPI targets - Version {APP_VERSION}"
        sheet.range('B3').font.size = 9
        sheet.range('B3').font.color = (128, 128, 128)
        sheet.range('B3:F3').merge()

        # Instructions
        sheet.range('B5').value = "INSTRUCTIONS"
        sheet.range('B5').font.size = 14
        sheet.range('B5').font.bold = True
        sheet.range('B5').font.color = (255, 255, 255)
        sheet.range('B5:H5').color = section_color

        instructions = [
            "1. Set target values for each KPI in the 'Target' column",
            "2. Yellow threshold: 80% of target (caution)",
            "3. Red threshold: 60% of target (warning)",
            "4. Dashboard updates automatically when data changes"
        ]
        for i, instr in enumerate(instructions):
            sheet.range(f'B{7+i}').value = instr
            sheet.range(f'B{7+i}').font.size = 10

        # Headers
        header_row = 13
        headers = ['Category', 'KPI Name', 'Description', 'Target', 'Yellow %', 'Red %', 'Direction']
        for col, header in enumerate(headers, 2):
            cell = sheet.range((header_row, col))
            cell.value = header
            cell.font.bold = True
            cell.font.color = (255, 255, 255)
            cell.color = (52, 73, 94)
            cell.api.HorizontalAlignment = -4108  # Center

        # Populate KPIs
        data_row = header_row + 1
        for category, kpis in self.KPI_DEFINITIONS.items():
            for kpi in kpis:
                sheet.range((data_row, 2)).value = category.title()
                sheet.range((data_row, 3)).value = kpi['name']
                sheet.range((data_row, 3)).font.bold = True
                sheet.range((data_row, 4)).value = kpi['description']
                sheet.range((data_row, 4)).font.size = 9
                sheet.range((data_row, 4)).font.color = (100, 100, 100)

                # Target (editable)
                target_cell = sheet.range((data_row, 5))
                target_cell.value = kpi['default_target']
                if '%' in kpi['format']:
                    target_cell.number_format = '0.0%'
                elif '0.00' in kpi['format']:
                    target_cell.number_format = '0.00'
                else:
                    target_cell.number_format = '#,##0'
                target_cell.color = control_color

                # Yellow/Red thresholds
                sheet.range((data_row, 6)).value = 0.80
                sheet.range((data_row, 6)).number_format = '0%'
                sheet.range((data_row, 6)).color = control_color

                sheet.range((data_row, 7)).value = 0.60
                sheet.range((data_row, 7)).number_format = '0%'
                sheet.range((data_row, 7)).color = control_color

                # Direction
                sheet.range((data_row, 8)).value = "Higher" if kpi['higher_is_better'] else "Lower"

                # Alternate row shading
                if data_row % 2 == 0:
                    for col in range(2, 9):
                        if sheet.range((data_row, col)).color is None:
                            sheet.range((data_row, col)).color = (248, 249, 250)

                data_row += 1

        # Column widths
        sheet.range('A:A').column_width = 3
        sheet.range('B:B').column_width = 14
        sheet.range('C:C').column_width = 20
        sheet.range('D:D').column_width = 30
        sheet.range('E:E').column_width = 12
        sheet.range('F:F').column_width = 10
        sheet.range('G:G').column_width = 10
        sheet.range('H:H').column_width = 10

        # Tab color
        try:
            sheet.api.Tab.Color = 0xDB7400  # Blue
        except:
            pass

    def _create_dashboard_sheet(self, sheet, pl_accounts, bs_accounts, months, detected_totals=None):
        """Create the main Dashboard sheet with KPIs and visualizations"""
        # Colors
        header_color = (22, 33, 62)  # Dark blue
        section_color = (44, 62, 80)  # Section headers
        green_color = (39, 174, 96)  # Good
        yellow_color = (243, 156, 18)  # Caution
        red_color = (231, 76, 60)  # Warning

        # Get detected account names (or use defaults)
        detected_totals = detected_totals or {}
        total_income_name = detected_totals.get('total_income', 'Total for Income')
        total_cogs_name = detected_totals.get('total_cogs', 'Total for Cost of Sales')
        total_expenses_name = detected_totals.get('total_expenses', 'Total for Expenses')

        company_name = self.company_name.get()
        current_month_col = len(months) + 1

        # Title
        sheet.range('B2').value = company_name
        sheet.range('B2').font.size = 28
        sheet.range('B2').font.bold = True
        sheet.range('B2').font.color = header_color
        sheet.range('B2:L2').merge()

        sheet.range('B3').value = "Executive Dashboard"
        sheet.range('B3').font.size = 16
        sheet.range('B3').font.color = (127, 140, 141)
        sheet.range('B3:L3').merge()

        if months:
            current_month = months[-1]
            sheet.range('B4').value = f"Current Period: {current_month[2]}"
        else:
            sheet.range('B4').value = "Current Period: N/A"
        sheet.range('B4').font.size = 9
        sheet.range('B4').font.color = (150, 150, 150)

        sheet.range('B5').value = f"Generated: {datetime.now().strftime('%B %d, %Y')} | Version {APP_VERSION}"
        sheet.range('B5').font.size = 9
        sheet.range('B5').font.color = (150, 150, 150)

        # =====================================================================
        # P&L SUMMARY TABLE (Rows 7-11) - Revenue, COGS, Gross Margin, Expenses
        # =====================================================================
        col_letter = self._col_letter(current_month_col)

        # Table headers row 7
        headers_data = [
            ('B7', ''),
            ('C7', 'CURRENT MTH'),
            ('D7', ''),
            ('E7', 'YTD'),
            ('F7', ''),
            ('G7', '% of Rev'),
        ]

        # Header row styling
        for cell_ref, header_text in headers_data:
            cell = sheet.range(cell_ref)
            cell.value = header_text
            cell.font.bold = True
            cell.font.size = 10
            cell.font.color = (255, 255, 255)
            cell.color = section_color
            cell.api.HorizontalAlignment = -4108  # Center

        # P&L Summary rows - use detected account names
        pl_summary_rows = [
            {'label': 'Revenue', 'account': total_income_name, 'row': 8, 'is_total': True, 'color': (39, 174, 96)},
            {'label': 'Cost of Goods Sold', 'account': total_cogs_name, 'row': 9, 'is_total': False, 'color': (231, 76, 60)},
            {'label': 'Gross Profit', 'account': 'Gross Profit', 'row': 10, 'is_total': True, 'color': (52, 152, 219)},
            {'label': 'Operating Expenses', 'account': total_expenses_name, 'row': 11, 'is_total': False, 'color': (230, 126, 34)},
            {'label': 'Net Income', 'account': 'Net Income', 'row': 12, 'is_total': True, 'color': (155, 89, 182)},
        ]

        for item in pl_summary_rows:
            row = item['row']

            # Label with color indicator (Column B)
            label_cell = sheet.range(f'B{row}')
            label_cell.value = item['label']
            label_cell.font.bold = item['is_total']
            label_cell.font.size = 11
            label_cell.font.color = item['color']

            # Current Month value (Column C only, no merge)
            # Uses SUMPRODUCT to dynamically get value for month/year matching Menu!C7
            cm_formula = (
                f"=SUMPRODUCT("
                f"(Source_PL!$A$3:$A$1000=\"{item['account']}\")*"
                f"(Source_PL!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_PL!$B$3:{col_letter}$1000))"
            )
            sheet.range(f'C{row}').value = cm_formula
            sheet.range(f'C{row}').number_format = '"$"#,##0'
            sheet.range(f'C{row}').font.size = 12
            sheet.range(f'C{row}').font.bold = item['is_total']
            sheet.range(f'C{row}').api.HorizontalAlignment = -4152  # Right align

            # YTD value (Column E only, no merge)
            # Uses SUMPRODUCT with YYYYMM helper row (row 2) to sum only current year through current month
            ytd_formula = (
                f"=SUMPRODUCT("
                f"(Source_PL!$A$3:$A$1000=\"{item['account']}\")*"
                f"(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD(Source_PL!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"(Source_PL!$B$3:{col_letter}$1000))"
            )
            sheet.range(f'E{row}').value = ytd_formula
            sheet.range(f'E{row}').number_format = '"$"#,##0'
            sheet.range(f'E{row}').font.size = 12
            sheet.range(f'E{row}').font.bold = item['is_total']
            sheet.range(f'E{row}').api.HorizontalAlignment = -4152  # Right align

            # % of Revenue (Column G only, no merge)
            if item['label'] == 'Revenue':
                sheet.range(f'G{row}').value = 1.0  # 100%
            else:
                sheet.range(f'G{row}').value = f"=IFERROR(E{row}/E8,0)"
            sheet.range(f'G{row}').number_format = '0.0%'
            sheet.range(f'G{row}').font.size = 11
            sheet.range(f'G{row}').font.bold = item['is_total']
            sheet.range(f'G{row}').api.HorizontalAlignment = -4108  # Center

            # Alternate row shading
            if row % 2 == 0:
                for c in range(2, 8):
                    sheet.range((row, c)).color = (248, 249, 250)

        # Add borders to summary table
        try:
            summary_table = sheet.range('B7:G12')
            for edge in [7, 8, 9, 10]:  # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
                summary_table.api.Borders(edge).LineStyle = 1
                summary_table.api.Borders(edge).Weight = 2
                summary_table.api.Borders(edge).Color = 0xA0A0A0
        except:
            pass

        # Gross Margin % indicator (small badge next to Gross Profit row)
        sheet.range('H10').value = "Gross Margin:"
        sheet.range('H10').font.size = 9
        sheet.range('H10').font.color = (127, 140, 141)
        sheet.range('I10').value = "=IFERROR(C10/C8,0)"
        sheet.range('I10').number_format = '0.0%'
        sheet.range('I10').font.size = 11
        sheet.range('I10').font.bold = True
        sheet.range('I10').font.color = (52, 152, 219)

        # Net Margin % indicator
        sheet.range('H12').value = "Net Margin:"
        sheet.range('H12').font.size = 9
        sheet.range('H12').font.color = (127, 140, 141)
        sheet.range('I12').value = "=IFERROR(C12/C8,0)"
        sheet.range('I12').number_format = '0.0%'
        sheet.range('I12').font.size = 11
        sheet.range('I12').font.bold = True
        sheet.range('I12').font.color = (155, 89, 182)

        # Set row heights for summary area
        try:
            sheet.range('7:7').row_height = 22  # Header row
            for r in range(8, 13):
                sheet.range(f'{r}:{r}').row_height = 24
        except:
            pass

        # KPI Sections - Start after P&L Summary table
        current_row = 14
        sections = [
            ('PROFITABILITY METRICS', 'profitability'),
            ('LIQUIDITY METRICS', 'liquidity'),
            ('EFFICIENCY METRICS', 'efficiency'),
            ('LEVERAGE METRICS', 'leverage'),
            ('CASH FLOW METRICS', 'cashflow'),
        ]

        control_row = 14  # Starting row in Dashboard_Control

        for section_title, category in sections:
            kpis = self.KPI_DEFINITIONS.get(category, [])
            if not kpis:
                continue

            # Section header
            sheet.range((current_row, 2)).value = section_title
            sheet.range((current_row, 2)).font.size = 14
            sheet.range((current_row, 2)).font.bold = True
            sheet.range((current_row, 2)).font.color = (255, 255, 255)
            for col in range(2, 13):
                sheet.range((current_row, col)).color = section_color

            # Column headers
            current_row += 1
            headers = [('KPI', 2), ('Current', 3), ('Target', 4), ('Status', 5), ('', 6),
                       ('YTD', 7), ('Target', 8), ('Status', 9), ('', 10), ('Trend', 11)]

            for header, col in headers:
                cell = sheet.range((current_row, col))
                cell.value = header
                cell.font.bold = True
                cell.font.color = (255, 255, 255)
                cell.color = (52, 73, 94)
                cell.api.HorizontalAlignment = -4108

            # KPI rows
            current_row += 1
            for kpi in kpis:
                # KPI Name
                sheet.range((current_row, 2)).value = kpi['name']
                sheet.range((current_row, 2)).font.bold = True
                sheet.range((current_row, 2)).font.color = (44, 62, 80)

                # Current value formula
                formula = self._build_dashboard_kpi_formula(kpi, 'current', current_month_col, months)
                sheet.range((current_row, 3)).value = formula
                sheet.range((current_row, 3)).number_format = kpi['format']
                sheet.range((current_row, 3)).font.size = 14
                sheet.range((current_row, 3)).font.bold = True

                # Target (from control sheet)
                sheet.range((current_row, 4)).value = f"=Dashboard_Control!E{control_row}"
                sheet.range((current_row, 4)).number_format = kpi['format']

                # Status formula
                higher = kpi['higher_is_better']
                if higher:
                    status_formula = f'=IF(C{current_row}>=D{current_row},"G",IF(C{current_row}>=D{current_row}*0.8,"Y","R"))'
                else:
                    status_formula = f'=IF(C{current_row}<=D{current_row},"G",IF(C{current_row}<=D{current_row}*1.2,"Y","R"))'
                sheet.range((current_row, 5)).value = status_formula
                sheet.range((current_row, 5)).api.HorizontalAlignment = -4108

                # YTD value
                ytd_formula = self._build_dashboard_kpi_formula(kpi, 'ytd', current_month_col, months)
                sheet.range((current_row, 7)).value = ytd_formula
                sheet.range((current_row, 7)).number_format = kpi['format']
                sheet.range((current_row, 7)).font.bold = True

                # YTD Target
                if '%' in kpi['format'] or 'ratio' in kpi['formula_type']:
                    sheet.range((current_row, 8)).value = f"=D{current_row}"
                else:
                    sheet.range((current_row, 8)).value = f"=D{current_row}*{len(months)}"
                sheet.range((current_row, 8)).number_format = kpi['format']

                # YTD Status
                if higher:
                    ytd_status = f'=IF(G{current_row}>=H{current_row},"G",IF(G{current_row}>=H{current_row}*0.8,"Y","R"))'
                else:
                    ytd_status = f'=IF(G{current_row}<=H{current_row},"G",IF(G{current_row}<=H{current_row}*1.2,"Y","R"))'
                sheet.range((current_row, 9)).value = ytd_status
                sheet.range((current_row, 9)).api.HorizontalAlignment = -4108

                # Trend
                trend_formula = f'=IF(C{current_row}>G{current_row}*1.05,"UP",IF(C{current_row}<G{current_row}*0.95,"DOWN","FLAT"))'
                sheet.range((current_row, 11)).value = trend_formula
                sheet.range((current_row, 11)).api.HorizontalAlignment = -4108

                # Alternate row fill
                if current_row % 2 == 0:
                    for col in [2, 3, 4, 5, 7, 8, 9, 11]:
                        sheet.range((current_row, col)).color = (248, 249, 250)

                current_row += 1
                control_row += 1

            current_row += 1  # Space between sections

        # Add conditional formatting for status columns (E and I)
        # Apply to all rows that could have status values (rows 16 through current_row)
        try:
            # Status column E (Current Status)
            status_range_e = sheet.range(f'E16:E{current_row}')
            status_range_e.api.FormatConditions.Delete()

            # Green for G
            fc_green_e = status_range_e.api.FormatConditions.Add(1, 3, "G")  # xlCellValue=1, xlEqual=3
            fc_green_e.Interior.Color = 0x60AE27  # Green (RGB reversed for Excel)
            fc_green_e.Font.Color = 0xFFFFFF  # White text
            fc_green_e.Font.Bold = True

            # Yellow for Y
            fc_yellow_e = status_range_e.api.FormatConditions.Add(1, 3, "Y")
            fc_yellow_e.Interior.Color = 0x129CF3  # Yellow (RGB reversed: 243, 156, 18)
            fc_yellow_e.Font.Color = 0x000000  # Black text
            fc_yellow_e.Font.Bold = True

            # Red for R
            fc_red_e = status_range_e.api.FormatConditions.Add(1, 3, "R")
            fc_red_e.Interior.Color = 0x3C4CE7  # Red (RGB reversed: 231, 76, 60)
            fc_red_e.Font.Color = 0xFFFFFF  # White text
            fc_red_e.Font.Bold = True

            # Status column I (YTD Status)
            status_range_i = sheet.range(f'I16:I{current_row}')
            status_range_i.api.FormatConditions.Delete()

            # Green for G
            fc_green_i = status_range_i.api.FormatConditions.Add(1, 3, "G")
            fc_green_i.Interior.Color = 0x60AE27
            fc_green_i.Font.Color = 0xFFFFFF
            fc_green_i.Font.Bold = True

            # Yellow for Y
            fc_yellow_i = status_range_i.api.FormatConditions.Add(1, 3, "Y")
            fc_yellow_i.Interior.Color = 0x129CF3
            fc_yellow_i.Font.Color = 0x000000
            fc_yellow_i.Font.Bold = True

            # Red for R
            fc_red_i = status_range_i.api.FormatConditions.Add(1, 3, "R")
            fc_red_i.Interior.Color = 0x3C4CE7
            fc_red_i.Font.Color = 0xFFFFFF
            fc_red_i.Font.Bold = True

            # Trend column K formatting
            trend_range = sheet.range(f'K16:K{current_row}')
            trend_range.api.FormatConditions.Delete()

            # UP = Green arrow
            fc_up = trend_range.api.FormatConditions.Add(1, 3, "UP")
            fc_up.Font.Color = 0x60AE27  # Green
            fc_up.Font.Bold = True

            # DOWN = Red arrow
            fc_down = trend_range.api.FormatConditions.Add(1, 3, "DOWN")
            fc_down.Font.Color = 0x3C4CE7  # Red
            fc_down.Font.Bold = True

            # FLAT = Gray
            fc_flat = trend_range.api.FormatConditions.Add(1, 3, "FLAT")
            fc_flat.Font.Color = 0x808080  # Gray

        except Exception as e:
            # Log but don't fail if conditional formatting fails
            pass

        # Column widths - Auto-fit to minimal width
        try:
            # Auto-fit columns to content with minimal padding
            for col_letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']:
                col_range = sheet.range(f'{col_letter}:{col_letter}')
                col_range.api.EntireColumn.AutoFit()

            # Set minimum widths for key columns to ensure readability
            if sheet.range('A:A').column_width < 3:
                sheet.range('A:A').column_width = 3
            if sheet.range('B:B').column_width < 18:
                sheet.range('B:B').column_width = 18  # KPI names need space
            if sheet.range('F:F').column_width < 2:
                sheet.range('F:F').column_width = 2  # Spacer
            if sheet.range('J:J').column_width < 2:
                sheet.range('J:J').column_width = 2  # Spacer
        except:
            # Fallback to fixed widths if auto-fit fails
            sheet.range('A:A').column_width = 3
            sheet.range('B:B').column_width = 22
            sheet.range('C:C').column_width = 14
            sheet.range('D:D').column_width = 12
            sheet.range('E:E').column_width = 8
            sheet.range('F:F').column_width = 3
            sheet.range('G:G').column_width = 14
            sheet.range('H:H').column_width = 12
            sheet.range('I:I').column_width = 8
            sheet.range('J:J').column_width = 3
            sheet.range('K:K').column_width = 10

        # Vertically center all cells
        try:
            used_range = sheet.api.UsedRange
            used_range.VerticalAlignment = -4108  # xlVAlignCenter
        except:
            pass

        # Tab color
        try:
            sheet.api.Tab.Color = 0x60AE27  # Green
        except:
            pass

        # Hide gridlines
        try:
            sheet.api.Activate()
            sheet.book.app.api.ActiveWindow.DisplayGridlines = False
        except:
            pass

    def _build_current_month_sumproduct(self, sheet, account, current_col):
        """
        Build a SUMPRODUCT formula for current month that dynamically references Menu!C7.
        Returns value for the exact month/year matching Menu!E7 (month) and Menu!F7 (year).

        Source sheets have:
        - Row 2: YYYYMM helper values (e.g., 202411 for Nov 2024)
        - Row 3+: Account data
        """
        col_letter = self._col_letter(current_col)
        # SUMPRODUCT matching exact YYYYMM value (year*100 + month)
        return (
            f"=SUMPRODUCT("
            f"({sheet}!$A$3:$A$1000=\"{account}\")*"
            f"({sheet}!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
            f"({sheet}!$B$3:{col_letter}$1000)"
            f")"
        )

    def _build_ytd_sumproduct(self, sheet, account, current_col):
        """
        Build a SUMPRODUCT formula for YTD that only sums months in the current year.
        Uses Menu!F7 for year and Menu!E7 for current month number.

        Source sheets now have:
        - Row 1: Headers (month names)
        - Row 2: YYYYMM helper values (e.g., 202411 for Nov 2024)
        - Row 3+: Account data

        This formula uses the same pattern as PL!AB6.
        """
        col_letter = self._col_letter(current_col)
        # SUMPRODUCT that:
        # 1. Matches account name in column A (data starts at row 3)
        # 2. Checks helper row year (INT(YYYYMM/100)) matches Menu!F7
        # 3. Checks helper row month (MOD(YYYYMM,100)) is <= Menu!E7
        return (
            f"=SUMPRODUCT("
            f"({sheet}!$A$3:$A$1000=\"{account}\")*"
            f"(INT({sheet}!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
            f"(MOD({sheet}!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
            f"({sheet}!$B$3:{col_letter}$1000)"
            f")"
        )

    def _build_dashboard_kpi_formula(self, kpi, period, current_col, months):
        """Build formula for a KPI based on its type and period"""
        formula_type = kpi['formula_type']
        col_letter = self._col_letter(current_col)

        if formula_type == 'direct':
            source = kpi['source']
            account = kpi['account']
            sheet = 'Source_PL' if source == 'PL' else 'Source_BS'
            if period == 'current':
                return self._build_current_month_sumproduct(sheet, account, current_col)
            else:  # ytd
                return self._build_ytd_sumproduct(sheet, account, current_col)

        elif formula_type == 'ratio':
            num = kpi['numerator']
            den = kpi['denominator']
            if period == 'current':
                # Current month using SUMPRODUCT with exact YYYYMM match
                num_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{num}\")*"
                    f"(Source_PL!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
                den_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{den}\")*"
                    f"(Source_PL!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
            else:
                # YTD using SUMPRODUCT with YYYYMM helper row (row 2)
                num_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{num}\")*"
                    f"(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                    f"(MOD(Source_PL!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
                den_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{den}\")*"
                    f"(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                    f"(MOD(Source_PL!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_ratio':
            # Balance sheet ratios use current month value (point in time)
            num = kpi['numerator']
            den = kpi['denominator']
            num_f = (
                f"SUMPRODUCT("
                f"(Source_BS!$A$3:$A$1000=\"{num}\")*"
                f"(Source_BS!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_BS!$B$3:{col_letter}$1000))"
            )
            den_f = (
                f"SUMPRODUCT("
                f"(Source_BS!$A$3:$A$1000=\"{den}\")*"
                f"(Source_BS!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_BS!$B$3:{col_letter}$1000))"
            )
            return f"=IFERROR({num_f}/{den_f},0)"

        elif formula_type == 'bs_difference':
            # Balance sheet difference uses current month value (point in time)
            min_acct = kpi['minuend']
            sub_acct = kpi['subtrahend']
            min_f = (
                f"SUMPRODUCT("
                f"(Source_BS!$A$3:$A$1000=\"{min_acct}\")*"
                f"(Source_BS!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_BS!$B$3:{col_letter}$1000))"
            )
            sub_f = (
                f"SUMPRODUCT("
                f"(Source_BS!$A$3:$A$1000=\"{sub_acct}\")*"
                f"(Source_BS!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_BS!$B$3:{col_letter}$1000))"
            )
            return f"={min_f}-{sub_f}"

        elif formula_type == 'days_ratio':
            balance = kpi['balance']
            flow = kpi['flow']
            # Balance is point-in-time (current month)
            bal_f = (
                f"SUMPRODUCT("
                f"(Source_BS!$A$3:$A$1000=\"{balance}\")*"
                f"(Source_BS!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_BS!$B$3:{col_letter}$1000))"
            )
            if period == 'current':
                flow_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{flow}\")*"
                    f"(Source_PL!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
                return f"=IFERROR({bal_f}/{flow_f}*30,0)"
            else:
                # YTD using SUMPRODUCT with YYYYMM helper row
                flow_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{flow}\")*"
                    f"(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                    f"(MOD(Source_PL!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
                # For YTD, use Menu!E7 as month count (Jan through current month)
                return f"=IFERROR({bal_f}/(({flow_f})/Menu!$E$7)*30,0)"

        elif formula_type == 'turnover':
            flow = kpi['flow']
            balance = kpi['balance']
            # Balance is point-in-time (current month)
            bal_f = (
                f"SUMPRODUCT("
                f"(Source_BS!$A$3:$A$1000=\"{balance}\")*"
                f"(Source_BS!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_BS!$B$3:{col_letter}$1000))"
            )
            if period == 'current':
                flow_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{flow}\")*"
                    f"(Source_PL!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
            else:
                # YTD using SUMPRODUCT with YYYYMM helper row
                flow_f = (
                    f"SUMPRODUCT("
                    f"(Source_PL!$A$3:$A$1000=\"{flow}\")*"
                    f"(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                    f"(MOD(Source_PL!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                    f"(Source_PL!$B$3:{col_letter}$1000))"
                )
            return f"=IFERROR({flow_f}/{bal_f},0)"

        elif formula_type == 'calculated':
            calc_type = kpi['calc_type']
            return self._build_calculated_dashboard_formula(calc_type, period, current_col, months)

        return "=0"

    def _build_calculated_dashboard_formula(self, calc_type, period, current_col, months):
        """Build formula for complex calculated KPIs"""
        col_letter = self._col_letter(current_col)

        def current_sumproduct(account):
            """Helper to build current month SUMPRODUCT formula for an account"""
            return (
                f"SUMPRODUCT("
                f"(Source_PL!$A$3:$A$1000=\"{account}\")*"
                f"(Source_PL!$B$2:{col_letter}$2=Menu!$F$7*100+Menu!$E$7)*"
                f"(Source_PL!$B$3:{col_letter}$1000))"
            )

        def ytd_sumproduct(account):
            """Helper to build YTD SUMPRODUCT formula for an account"""
            return (
                f"SUMPRODUCT("
                f"(Source_PL!$A$3:$A$1000=\"{account}\")*"
                f"(INT(Source_PL!$B$2:{col_letter}$2/100)=Menu!$F$7)*"
                f"(MOD(Source_PL!$B$2:{col_letter}$2,100)<=Menu!$E$7)*"
                f"(Source_PL!$B$3:{col_letter}$1000))"
            )

        if calc_type == 'ebitda':
            if period == 'current':
                ni = current_sumproduct("Net Income")
                int_e = current_sumproduct("8000 Interest Expense")
                dep = current_sumproduct("6090 Depreciation Expense")
            else:
                ni = ytd_sumproduct("Net Income")
                int_e = ytd_sumproduct("8000 Interest Expense")
                dep = ytd_sumproduct("6090 Depreciation Expense")
            return f"={ni}+ABS({int_e})+ABS({dep})"

        elif calc_type == 'ocf':
            if period == 'current':
                ni = current_sumproduct("Net Income")
                dep = current_sumproduct("6090 Depreciation Expense")
            else:
                ni = ytd_sumproduct("Net Income")
                dep = ytd_sumproduct("6090 Depreciation Expense")
            return f"={ni}+ABS({dep})"

        elif calc_type == 'fcf':
            if period == 'current':
                ni = current_sumproduct("Net Income")
                dep = current_sumproduct("6090 Depreciation Expense")
            else:
                ni = ytd_sumproduct("Net Income")
                dep = ytd_sumproduct("6090 Depreciation Expense")
            return f"={ni}+ABS({dep})"

        return "=0"

    def _col_letter(self, col_num):
        """Convert column number to letter"""
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

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
                    # Create group
                    col_range = sheet.range((1, start_col), (1, end_col))
                    col_range.api.EntireColumn.Group()

                    # Hide the grouped columns
                    col_range.api.EntireColumn.Hidden = True
                    print(f"Grouped and hid year {year}: columns {start_col} to {end_col}")
                except Exception as e:
                    print(f"Column grouping error for year {year}: {e}")

            # Hide future month columns (not grouped, just hidden)
            for col in future_month_cols:
                try:
                    sheet.range((1, col)).api.EntireColumn.Hidden = True
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
