Attribute VB_Name = "FinancialModel"
' ============================================
' CFO Financial Model - VBA Module
' Version: 1.0
' ============================================
Option Explicit

' Module-level constants
Private Const SOURCE_PL_SHEET As String = "Source P&L"
Private Const SOURCE_BS_SHEET As String = "Source BS"
Private Const PL_SHEET As String = "P&L"
Private Const BS_SHEET As String = "Balance Sheet"
Private Const CF_SHEET As String = "Cash Flow"
Private Const MENU_SHEET As String = "Menu"
Private Const HELPER_NOTES_SHEET As String = "Helper_Notes"
Private Const DIAGNOSTICS_SHEET As String = "Diagnostics"

' ============================================
' UPLOAD FUNCTIONS
' ============================================

Public Sub UploadPLFile()
    ' Upload and import new P&L data
    Dim filePath As String
    Dim fileFilter As String

    fileFilter = "Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv"
    filePath = Application.GetOpenFilename(fileFilter, , "Select P&L File")

    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportFinancialData(filePath, SOURCE_PL_SHEET)
    Call UpdateReportSheetRows(PL_SHEET, SOURCE_PL_SHEET)
    Call RefreshPLSheet
    Call LogDiagnostic("P&L data imported successfully from: " & filePath)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "P&L data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call LogDiagnostic("Error importing P&L: " & Err.Description)
    MsgBox "Error importing P&L data: " & Err.Description, vbCritical
End Sub

Public Sub UploadBSFile()
    ' Upload and import new Balance Sheet data
    Dim filePath As String
    Dim fileFilter As String

    fileFilter = "Excel Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv"
    filePath = Application.GetOpenFilename(fileFilter, , "Select Balance Sheet File")

    If filePath = "False" Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ImportFinancialData(filePath, SOURCE_BS_SHEET)
    Call UpdateReportSheetRows(BS_SHEET, SOURCE_BS_SHEET)
    Call RefreshBSSheet
    Call RefreshCashFlowSheet
    Call LogDiagnostic("Balance Sheet data imported successfully from: " & filePath)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Balance Sheet data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call LogDiagnostic("Error importing BS: " & Err.Description)
    MsgBox "Error importing Balance Sheet data: " & Err.Description, vbCritical
End Sub

Private Sub ImportFinancialData(filePath As String, targetSheet As String)
    ' Import data from file to target sheet
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim headerRow As Long

    ' Open source file
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)
    Set tgtWs = ThisWorkbook.Sheets(targetSheet)

    ' Find header row (contains month names)
    headerRow = FindHeaderRow(srcWs)

    If headerRow = 0 Then
        srcWb.Close SaveChanges:=False
        Err.Raise vbObjectError + 1, , "Could not find month headers in file"
    End If

    ' Get data dimensions
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.Count).End(xlToLeft).Column

    ' Check for new accounts and add them
    Call CheckForNewAccounts(srcWs, tgtWs, headerRow, lastRow)

    ' Check for new months and add columns
    Call CheckForNewMonths(srcWs, tgtWs, headerRow, lastCol)

    ' Update data values
    Call UpdateSourceData(srcWs, tgtWs, headerRow, lastRow, lastCol)

    srcWb.Close SaveChanges:=False
End Sub

Private Function FindHeaderRow(ws As Worksheet) As Long
    ' Find the row containing month headers
    Dim i As Long, j As Long
    Dim cellVal As String
    Dim monthNames As Variant

    monthNames = Array("january", "february", "march", "april", "may", "june", _
                       "july", "august", "september", "october", "november", "december")

    For i = 1 To 10
        For j = 1 To 20
            cellVal = LCase(Trim(CStr(ws.Cells(i, j).Value)))
            Dim m As Long
            For m = LBound(monthNames) To UBound(monthNames)
                If InStr(cellVal, monthNames(m)) > 0 Then
                    FindHeaderRow = i
                    Exit Function
                End If
            Next m
        Next j
    Next i

    FindHeaderRow = 0
End Function

Private Sub CheckForNewAccounts(srcWs As Worksheet, tgtWs As Worksheet, headerRow As Long, lastRow As Long)
    ' Check for accounts in source that are not in target
    Dim srcAccounts As Object
    Dim tgtAccounts As Object
    Dim i As Long
    Dim accountName As String
    Dim newAccounts As Collection
    Dim tgtLastRow As Long

    Set srcAccounts = CreateObject("Scripting.Dictionary")
    Set tgtAccounts = CreateObject("Scripting.Dictionary")
    Set newAccounts = New Collection

    ' Get source accounts
    For i = headerRow + 1 To lastRow
        accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
        If Len(accountName) > 0 And Not srcAccounts.Exists(accountName) Then
            srcAccounts.Add accountName, i
        End If
    Next i

    ' Get target accounts
    tgtLastRow = tgtWs.Cells(tgtWs.Rows.Count, 1).End(xlUp).Row
    For i = 2 To tgtLastRow
        accountName = Trim(CStr(tgtWs.Cells(i, 1).Value))
        If Len(accountName) > 0 And Not tgtAccounts.Exists(accountName) Then
            tgtAccounts.Add accountName, i
        End If
    Next i

    ' Find new accounts
    Dim key As Variant
    For Each key In srcAccounts.Keys
        If Not tgtAccounts.Exists(key) Then
            newAccounts.Add key
        End If
    Next key

    ' Add new accounts to target
    If newAccounts.Count > 0 Then
        Call AddNewAccounts(tgtWs, newAccounts, tgtLastRow)
        Call LogDiagnostic("Added " & newAccounts.Count & " new account(s)")
    End If
End Sub

Private Sub AddNewAccounts(tgtWs As Worksheet, newAccounts As Collection, afterRow As Long)
    ' Add new accounts to the target sheet
    Dim i As Long
    Dim lastCol As Long

    lastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column

    For i = 1 To newAccounts.Count
        afterRow = afterRow + 1
        tgtWs.Cells(afterRow, 1).Value = newAccounts(i)

        ' Initialize data columns to 0
        Dim j As Long
        For j = 2 To lastCol
            tgtWs.Cells(afterRow, j).Value = 0
            tgtWs.Cells(afterRow, j).NumberFormat = "#,##0"
        Next j
    Next i
End Sub

Private Sub CheckForNewMonths(srcWs As Worksheet, tgtWs As Worksheet, headerRow As Long, lastCol As Long)
    ' Check for months in source that are not in target and add them
    Dim srcMonth As String
    Dim tgtLastCol As Long
    Dim foundMonth As Boolean
    Dim i As Long, j As Long
    Dim lastRow As Long

    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column
    lastRow = tgtWs.Cells(tgtWs.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastCol
        srcMonth = Trim(CStr(srcWs.Cells(headerRow, i).Value))
        If Len(srcMonth) > 0 And LCase(srcMonth) <> "total" Then
            foundMonth = False
            For j = 2 To tgtLastCol
                If LCase(Trim(CStr(tgtWs.Cells(1, j).Value))) = LCase(srcMonth) Then
                    foundMonth = True
                    Exit For
                End If
            Next j

            If Not foundMonth Then
                ' Add new month column at the end
                tgtLastCol = tgtLastCol + 1
                tgtWs.Cells(1, tgtLastCol).Value = srcMonth
                tgtWs.Cells(1, tgtLastCol).Font.Bold = True

                ' Initialize data to 0
                For j = 2 To lastRow
                    tgtWs.Cells(j, tgtLastCol).Value = 0
                    tgtWs.Cells(j, tgtLastCol).NumberFormat = "#,##0"
                Next j

                Call LogDiagnostic("Added new month column: " & srcMonth)
            End If
        End If
    Next i

    ' Update named range
    Call UpdateSourceNamedRange(tgtWs)
End Sub

Private Sub UpdateSourceNamedRange(ws As Worksheet)
    ' Update the named range for source data
    Dim lastRow As Long, lastCol As Long
    Dim rangeName As String
    Dim rangeAddress As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ws.Name = SOURCE_PL_SHEET Then
        rangeName = "SourcePL_Data"
    ElseIf ws.Name = SOURCE_BS_SHEET Then
        rangeName = "SourceBS_Data"
    Else
        Exit Sub
    End If

    rangeAddress = "'" & ws.Name & "'!$A$1:$" & ColLetter(lastCol) & "$" & lastRow

    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:="=" & rangeAddress
End Sub

Private Sub UpdateSourceData(srcWs As Worksheet, tgtWs As Worksheet, headerRow As Long, lastRow As Long, lastCol As Long)
    ' Update source data with new values
    Dim i As Long, j As Long, k As Long
    Dim accountName As String
    Dim monthName As String
    Dim tgtRow As Long, tgtCol As Long
    Dim tgtLastRow As Long, tgtLastCol As Long

    tgtLastRow = tgtWs.Cells(tgtWs.Rows.Count, 1).End(xlUp).Row
    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column

    For i = headerRow + 1 To lastRow
        accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
        If Len(accountName) = 0 Then GoTo NextRow

        ' Find account row in target
        tgtRow = 0
        For j = 2 To tgtLastRow
            If Trim(CStr(tgtWs.Cells(j, 1).Value)) = accountName Then
                tgtRow = j
                Exit For
            End If
        Next j

        If tgtRow = 0 Then GoTo NextRow

        ' Update each month column
        For j = 2 To lastCol
            monthName = Trim(CStr(srcWs.Cells(headerRow, j).Value))
            If Len(monthName) = 0 Or LCase(monthName) = "total" Then GoTo NextCol

            ' Find month column in target
            tgtCol = 0
            For k = 2 To tgtLastCol
                If LCase(Trim(CStr(tgtWs.Cells(1, k).Value))) = LCase(monthName) Then
                    tgtCol = k
                    Exit For
                End If
            Next k

            If tgtCol > 0 Then
                Dim cellValue As Variant
                cellValue = srcWs.Cells(i, j).Value
                If IsNumeric(cellValue) Then
                    tgtWs.Cells(tgtRow, tgtCol).Value = CDbl(cellValue)
                Else
                    tgtWs.Cells(tgtRow, tgtCol).Value = 0
                End If
            End If
NextCol:
        Next j
NextRow:
    Next i
End Sub

Private Sub UpdateReportSheetRows(reportSheet As String, sourceSheet As String)
    ' Update report sheet to include any new accounts from source
    Dim srcWs As Worksheet, rptWs As Worksheet
    Dim srcAccounts As Object, rptAccounts As Object
    Dim i As Long
    Dim accountName As String
    Dim srcLastRow As Long, rptLastRow As Long
    Dim headerRow As Long
    Dim dataStartRow As Long

    Set srcWs = ThisWorkbook.Sheets(sourceSheet)
    Set rptWs = ThisWorkbook.Sheets(reportSheet)
    Set srcAccounts = CreateObject("Scripting.Dictionary")
    Set rptAccounts = CreateObject("Scripting.Dictionary")

    ' Find data start row in report (after headers)
    dataStartRow = 5  ' Typically row 5 after titles and headers

    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    rptLastRow = rptWs.Cells(rptWs.Rows.Count, 1).End(xlUp).Row

    ' Get source accounts
    For i = 2 To srcLastRow
        accountName = Trim(CStr(srcWs.Cells(i, 1).Value))
        If Len(accountName) > 0 And Not srcAccounts.Exists(accountName) Then
            srcAccounts.Add accountName, i
        End If
    Next i

    ' Get report accounts
    For i = dataStartRow To rptLastRow
        accountName = Trim(CStr(rptWs.Cells(i, 1).Value))
        If Len(accountName) > 0 And Not rptAccounts.Exists(accountName) Then
            rptAccounts.Add accountName, i
        End If
    Next i

    ' Add missing accounts to report
    Dim newRow As Long
    newRow = rptLastRow
    Dim key As Variant
    For Each key In srcAccounts.Keys
        If Not rptAccounts.Exists(key) Then
            newRow = newRow + 1
            rptWs.Cells(newRow, 1).Value = key
            ' Add VLOOKUP formulas for each month column
            Call AddRowFormulas(rptWs, newRow, sourceSheet)
        End If
    Next key
End Sub

Private Sub AddRowFormulas(ws As Worksheet, rowNum As Long, sourceSheet As String)
    ' Add VLOOKUP formulas for a new row
    Dim lastCol As Long
    Dim i As Long
    Dim rangeName As String

    If sourceSheet = SOURCE_PL_SHEET Then
        rangeName = "SourcePL_Data"
    Else
        rangeName = "SourceBS_Data"
    End If

    ' Find last month column (before summary columns)
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    ' Add VLOOKUP for each month column
    For i = 2 To lastCol
        If ws.Cells(4, i).Value <> "" And ws.Cells(4, i).Value <> "Notes" Then
            ws.Cells(rowNum, i).Formula = "=IFERROR(VLOOKUP($A" & rowNum & "," & rangeName & "," & i & ",FALSE),0)"
            ws.Cells(rowNum, i).NumberFormat = "#,##0"
        End If
    Next i
End Sub

' ============================================
' REFRESH FUNCTIONS
' ============================================

Public Sub RefreshAllFormulas()
    ' Refresh all sheets
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call RefreshPLSheet
    Call RefreshBSSheet
    Call RefreshCashFlowSheet

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "All formulas refreshed!", vbInformation
End Sub

Private Sub RefreshPLSheet()
    ThisWorkbook.Sheets(PL_SHEET).Calculate
End Sub

Private Sub RefreshBSSheet()
    ThisWorkbook.Sheets(BS_SHEET).Calculate
End Sub

Private Sub RefreshCashFlowSheet()
    ThisWorkbook.Sheets(CF_SHEET).Calculate
End Sub

' ============================================
' NOTES FUNCTIONS
' ============================================

Public Sub SaveVarianceNote()
    ' Allow user to save a note for the selected account/month
    Dim ws As Worksheet
    Dim notesWs As Worksheet
    Dim accountName As String
    Dim currentMonth As String
    Dim noteText As String
    Dim lastRow As Long
    Dim foundRow As Long
    Dim i As Long

    Set ws = ActiveSheet
    Set notesWs = ThisWorkbook.Sheets(HELPER_NOTES_SHEET)

    ' Get account name from current row
    accountName = ws.Cells(ActiveCell.Row, 1).Value
    If Len(accountName) = 0 Then
        MsgBox "Please select a row with an account name.", vbExclamation
        Exit Sub
    End If

    ' Get current month from Menu
    currentMonth = ThisWorkbook.Sheets(MENU_SHEET).Range("C10").Value
    If Len(currentMonth) = 0 Then
        MsgBox "Please set the Current Month on the Menu sheet.", vbExclamation
        Exit Sub
    End If

    ' Prompt for note
    noteText = InputBox("Enter variance note for:" & vbCrLf & _
                       "Account: " & accountName & vbCrLf & _
                       "Month: " & currentMonth, "Variance Note")

    If Len(noteText) = 0 Then Exit Sub

    ' Save to Helper_Notes
    lastRow = notesWs.Cells(notesWs.Rows.Count, 1).End(xlUp).Row

    ' Check if note exists
    foundRow = 0
    For i = 2 To lastRow
        If notesWs.Cells(i, 1).Value = accountName And _
           notesWs.Cells(i, 2).Value = currentMonth Then
            foundRow = i
            Exit For
        End If
    Next i

    If foundRow > 0 Then
        notesWs.Cells(foundRow, 3).Value = noteText
    Else
        lastRow = lastRow + 1
        notesWs.Cells(lastRow, 1).Value = accountName
        notesWs.Cells(lastRow, 2).Value = currentMonth
        notesWs.Cells(lastRow, 3).Value = noteText
        notesWs.Cells(lastRow, 4).Formula = "=A" & lastRow & "&B" & lastRow
    End If

    MsgBox "Note saved successfully!", vbInformation
End Sub

' ============================================
' DIAGNOSTICS FUNCTIONS
' ============================================

Public Sub RunDiagnostics()
    ' Run diagnostic checks on the model
    Dim ws As Worksheet
    Dim issues As Collection
    Dim i As Long
    Dim row As Long

    Set ws = ThisWorkbook.Sheets(DIAGNOSTICS_SHEET)
    Set issues = New Collection

    ' Clear previous diagnostics
    ws.Range("A20:C1000").ClearContents

    ' Check 1: P&L and BS month alignment
    If Not CheckMonthAlignment() Then
        issues.Add "P&L and Balance Sheet months do not align"
    End If

    ' Check 2: Balance Sheet balances
    Call CheckBSBalance(issues)

    ' Check 3: Missing data
    Call CheckMissingData(issues)

    ' Check 4: Cash Flow ties to BS
    Call CheckCashFlowTies(issues)

    ' Log results
    row = 20
    ws.Cells(row - 1, 1).Value = "Diagnostic Results - " & Format(Now(), "yyyy-mm-dd hh:mm")
    ws.Cells(row - 1, 1).Font.Bold = True

    If issues.Count = 0 Then
        ws.Cells(row, 1).Value = Now()
        ws.Cells(row, 2).Value = "All checks passed"
        ws.Cells(row, 3).Value = "No action needed"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    Else
        For i = 1 To issues.Count
            ws.Cells(row, 1).Value = Now()
            ws.Cells(row, 2).Value = issues(i)
            ws.Cells(row, 3).Value = "Review and correct"
            ws.Cells(row, 2).Font.Color = RGB(255, 0, 0)
            row = row + 1
        Next i
    End If

    ws.Activate
    MsgBox "Diagnostics complete. Found " & issues.Count & " issue(s).", vbInformation
End Sub

Private Function CheckMonthAlignment() As Boolean
    Dim plWs As Worksheet, bsWs As Worksheet
    Dim plLastCol As Long, bsLastCol As Long

    Set plWs = ThisWorkbook.Sheets(SOURCE_PL_SHEET)
    Set bsWs = ThisWorkbook.Sheets(SOURCE_BS_SHEET)

    plLastCol = plWs.Cells(1, plWs.Columns.Count).End(xlToLeft).Column
    bsLastCol = bsWs.Cells(1, bsWs.Columns.Count).End(xlToLeft).Column

    CheckMonthAlignment = (plLastCol = bsLastCol)
End Function

Private Sub CheckBSBalance(issues As Collection)
    ' Check if Balance Sheet balances for each month
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim i As Long
    Dim totalAssets As Double, totalLE As Double
    Dim assetsRow As Long, leRow As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets(SOURCE_BS_SHEET)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find Total Assets and Total L&E rows
    For i = 2 To lastRow
        If InStr(LCase(ws.Cells(i, 1).Value), "total for assets") > 0 Then
            assetsRow = i
        ElseIf InStr(LCase(ws.Cells(i, 1).Value), "total for liabilities and equity") > 0 Then
            leRow = i
        End If
    Next i

    If assetsRow = 0 Or leRow = 0 Then
        issues.Add "Cannot find Total Assets or Total L&E rows"
        Exit Sub
    End If

    ' Check each month
    For i = 2 To lastCol
        totalAssets = ws.Cells(assetsRow, i).Value
        totalLE = ws.Cells(leRow, i).Value

        If Abs(totalAssets - totalLE) > 0.01 Then
            issues.Add "BS out of balance in " & ws.Cells(1, i).Value & _
                      " (Assets: " & Format(totalAssets, "#,##0") & _
                      ", L+E: " & Format(totalLE, "#,##0") & ")"
        End If
    Next i
End Sub

Private Sub CheckMissingData(issues As Collection)
    ' Check for accounts with all zeros
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim accountName As String
    Dim hasData As Boolean

    Set ws = ThisWorkbook.Sheets(SOURCE_PL_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = 2 To lastRow
        accountName = ws.Cells(i, 1).Value
        If InStr(LCase(accountName), "total") = 0 And Len(accountName) > 0 Then
            hasData = False
            For j = 2 To lastCol
                If ws.Cells(i, j).Value <> 0 Then
                    hasData = True
                    Exit For
                End If
            Next j

            If Not hasData Then
                issues.Add "P&L Account '" & Left(accountName, 30) & "' has no data"
            End If
        End If
    Next i
End Sub

Private Sub CheckCashFlowTies(issues As Collection)
    ' Verify Cash Flow ending balance ties to BS cash
    ' This is a placeholder for more detailed checking
    issues.Add "Cash Flow to BS tie check: Manual verification recommended"
End Sub

Public Sub LogDiagnostic(message As String)
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets(DIAGNOSTICS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    If lastRow < 20 Then lastRow = 20

    ws.Cells(lastRow, 1).Value = Now()
    ws.Cells(lastRow, 2).Value = message
End Sub

' ============================================
' NAVIGATION FUNCTIONS
' ============================================

Public Sub GoToPL()
    ThisWorkbook.Sheets(PL_SHEET).Activate
End Sub

Public Sub GoToBS()
    ThisWorkbook.Sheets(BS_SHEET).Activate
End Sub

Public Sub GoToCF()
    ThisWorkbook.Sheets(CF_SHEET).Activate
End Sub

Public Sub GoToMenu()
    ThisWorkbook.Sheets(MENU_SHEET).Activate
End Sub

Public Sub GoToDiagnostics()
    ThisWorkbook.Sheets(DIAGNOSTICS_SHEET).Activate
End Sub

' ============================================
' HELPER FUNCTIONS
' ============================================

Private Function ColLetter(colNum As Long) As String
    ' Convert column number to letter
    Dim vArr As Variant
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    ColLetter = vArr(0)
End Function

' ============================================
' WORKBOOK EVENTS (Copy to ThisWorkbook module)
' ============================================
' Private Sub Workbook_Open()
'     ThisWorkbook.Sheets("Menu").Activate
' End Sub
