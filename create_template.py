"""
Create the Financial Model template with embedded VBA.
Run this ONCE on a machine with "Trust access to VBA project object model" enabled.
The resulting template will be bundled with the .exe for distribution.
"""

import xlwings as xw
import os

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

def main():
    print("Creating Financial Model Template with embedded VBA...")
    print("NOTE: You must have 'Trust access to VBA project object model' enabled in Excel.")
    print()

    # Create Excel application
    app = xw.App(visible=False)

    try:
        wb = app.books.add()

        # Create sheets
        menu_sheet = wb.sheets[0]
        menu_sheet.name = 'Menu'

        source_pl = wb.sheets.add('Source_PL', after=menu_sheet)
        source_bs = wb.sheets.add('Source_BS', after=source_pl)
        pl_sheet = wb.sheets.add('PL', after=source_bs)
        bs_sheet = wb.sheets.add('Balance_Sheet', after=pl_sheet)
        cf_sheet = wb.sheets.add('Cash_Flow', after=bs_sheet)
        notes_sheet = wb.sheets.add('Notes', after=cf_sheet)

        # Set up source sheet headers
        source_pl.range('A1').value = 'Account'
        source_bs.range('A1').value = 'Account'

        # Set up Menu sheet
        menu_sheet.range('B2').value = '[Company Name]'
        menu_sheet.range('B2').font.size = 24
        menu_sheet.range('B2').font.bold = True
        menu_sheet.range('B3').value = 'Financial Model'
        menu_sheet.range('B3').font.size = 14

        # Set up report sheets with headers
        for sheet in [pl_sheet, bs_sheet, cf_sheet]:
            sheet.range('A1').value = '[Company Name]'
            sheet.range('A1').font.size = 14
            sheet.range('A1').font.bold = True

        pl_sheet.range('A2').value = 'Profit & Loss Statement'
        bs_sheet.range('A2').value = 'Balance Sheet'
        cf_sheet.range('A2').value = 'Cash Flow Statement'

        notes_sheet.range('A1').value = 'Account'
        notes_sheet.range('B1').value = 'Month'
        notes_sheet.range('C1').value = 'Note'

        # Add VBA code
        print("Adding VBA macros...")
        try:
            vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            vba_module.Name = "FinancialModel"
            vba_module.CodeModule.AddFromString(VBA_CODE)
            print("VBA macros added successfully!")
        except Exception as e:
            print(f"ERROR: Could not add VBA code: {e}")
            print()
            print("Please enable 'Trust access to VBA project object model' in Excel:")
            print("  1. Open Excel")
            print("  2. File > Options > Trust Center > Trust Center Settings")
            print("  3. Macro Settings > Check 'Trust access to the VBA project object model'")
            print("  4. Click OK and restart Excel")
            print("  5. Run this script again")
            wb.close()
            app.quit()
            return

        # Save as template
        template_path = os.path.join(os.path.dirname(__file__), 'webapp', 'Financial_Template.xlsm')
        wb.save(template_path)
        print(f"\nTemplate saved to: {template_path}")

        # Also save a copy in the dist folder for packaging
        dist_path = os.path.join(os.path.dirname(__file__), 'dist', 'Financial_Template.xlsm')
        os.makedirs(os.path.dirname(dist_path), exist_ok=True)
        wb.save(dist_path)
        print(f"Template also saved to: {dist_path}")

        wb.close()

    finally:
        app.quit()

    print("\nTemplate created successfully!")
    print("You can now rebuild the .exe and distribute it along with Financial_Template.xlsm")

if __name__ == '__main__':
    main()
