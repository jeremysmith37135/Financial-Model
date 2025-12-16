
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
' Called when the current month dropdown is changed on Menu sheet
'-----------------------------------------------------
Public Sub OnCurrentMonthChange()
    ' This is called when current month changes on Menu sheet
    Call RefreshDashboard
    Call UpdateColumnVisibility
End Sub

'-----------------------------------------------------
' UpdateColumnVisibility - Hide prior years and future months
' Called when current month changes on Menu sheet
'-----------------------------------------------------
Public Sub UpdateColumnVisibility()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim wsMenu As Worksheet
    Dim wsPL As Worksheet
    Dim wsBS As Worksheet
    Dim currentMonthStr As String

    Set wsMenu = GetWorksheet("Menu")
    Set wsPL = GetWorksheet("PL")
    Set wsBS = GetWorksheet("Balance_Sheet")

    If wsMenu Is Nothing Then
        Call LogDiagnostic("UpdateColumnVisibility: Menu sheet not found")
        GoTo Cleanup
    End If

    ' Get current month from Menu (e.g., "Sep 2024" or "September 2024")
    currentMonthStr = Trim(CStr(wsMenu.Range("C11").Value))
    If Len(currentMonthStr) = 0 Then
        currentMonthStr = Trim(CStr(wsMenu.Range("C7").Value))
    End If

    If Len(currentMonthStr) = 0 Then
        Call LogDiagnostic("UpdateColumnVisibility: No current month found")
        GoTo Cleanup
    End If

    Call LogDiagnostic("UpdateColumnVisibility: Current month = " & currentMonthStr)

    ' Update visibility on P&L sheet
    If Not wsPL Is Nothing Then
        Call SetColumnVisibility(wsPL, currentMonthStr)
    End If

    ' Update visibility on Balance Sheet
    If Not wsBS Is Nothing Then
        Call SetColumnVisibility(wsBS, currentMonthStr)
    End If

    Call LogDiagnostic("Column visibility updated successfully")

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Call LogDiagnostic("Error in UpdateColumnVisibility: " & Err.Description)
    Resume Cleanup
End Sub

'-----------------------------------------------------
' SetColumnVisibility - Show/hide columns based on current month
' Hides prior year columns and months after current month
'-----------------------------------------------------
Private Sub SetColumnVisibility(ws As Worksheet, currentMonthStr As String)
    On Error Resume Next

    Dim headerRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim currentYear As Long
    Dim currentMonthNum As Long
    Dim colYear As Long
    Dim colMonthNum As Long
    Dim headerVal As String
    Dim foundCurrentMonth As Boolean

    headerRow = 4
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' Parse current month/year
    currentYear = ExtractYearFromString(currentMonthStr)
    currentMonthNum = ExtractMonthFromString(currentMonthStr)

    If currentYear = 0 Or currentMonthNum = 0 Then Exit Sub

    ' First, unhide all columns
    ws.Columns.Hidden = False

    ' Clear any existing outline/grouping
    On Error Resume Next
    ws.Cells.ClearOutline
    On Error GoTo 0

    ' Scan columns and set visibility
    foundCurrentMonth = False
    For col = 2 To lastCol
        headerVal = Trim(CStr(ws.Cells(headerRow, col).Value))

        ' Check if this is a month column
        If IsMonthColumn(headerVal) Then
            colYear = ExtractYearFromString(headerVal)
            colMonthNum = ExtractMonthFromString(headerVal)

            If colYear > 0 And colMonthNum > 0 Then
                ' Check if this is after the current month
                If colYear > currentYear Then
                    ' Future year - hide
                    ws.Columns(col).Hidden = True
                ElseIf colYear = currentYear And colMonthNum > currentMonthNum Then
                    ' Same year but future month - hide
                    ws.Columns(col).Hidden = True
                ElseIf colYear < currentYear Then
                    ' Prior year - group and hide
                    ws.Columns(col).Hidden = True
                Else
                    ' Current year, current or prior month - show
                    ws.Columns(col).Hidden = False
                End If
            End If
        End If
    Next col
End Sub

'-----------------------------------------------------
' IsMonthColumn - Check if a header value is a month column
'-----------------------------------------------------
Private Function IsMonthColumn(headerVal As String) As Boolean
    Dim months As Variant
    Dim i As Long

    months = Array("jan", "feb", "mar", "apr", "may", "jun", _
                   "jul", "aug", "sep", "oct", "nov", "dec")

    headerVal = LCase(Trim(headerVal))

    For i = LBound(months) To UBound(months)
        If InStr(headerVal, months(i)) > 0 Then
            IsMonthColumn = True
            Exit Function
        End If
    Next i

    IsMonthColumn = False
End Function

'-----------------------------------------------------
' ExtractYearFromString - Extract year from month string
'-----------------------------------------------------
Private Function ExtractYearFromString(monthStr As String) As Long
    Dim parts() As String
    Dim i As Long
    Dim y As Long

    ExtractYearFromString = 0
    monthStr = Replace(monthStr, "-", " ")
    parts = Split(Trim(monthStr), " ")

    For i = LBound(parts) To UBound(parts)
        If IsNumeric(parts(i)) Then
            y = CLng(parts(i))
            If y >= 2000 And y <= 2100 Then
                ExtractYearFromString = y
                Exit Function
            ElseIf y >= 0 And y <= 99 Then
                ExtractYearFromString = 2000 + y
                Exit Function
            End If
        End If
    Next i
End Function

'-----------------------------------------------------
' ExtractMonthFromString - Extract month number from month string
'-----------------------------------------------------
Private Function ExtractMonthFromString(monthStr As String) As Long
    Dim months As Variant
    Dim i As Long

    months = Array("jan", "feb", "mar", "apr", "may", "jun", _
                   "jul", "aug", "sep", "oct", "nov", "dec")

    monthStr = LCase(Trim(monthStr))

    For i = LBound(months) To UBound(months)
        If InStr(monthStr, months(i)) > 0 Then
            ExtractMonthFromString = i + 1
            Exit Function
        End If
    Next i

    ExtractMonthFromString = 0
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
        n = (n - 1) \ 26
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
' NOTE: This should be placed in ThisWorkbook module, not a standard module
'-----------------------------------------------------
Public Sub Workbook_OpenHandler()
    ' Activate Dashboard sheet and refresh on open
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    Application.OnTime Now + TimeValue("00:00:01"), "RefreshDashboard"
    Application.OnTime Now + TimeValue("00:00:02"), "UpdateYearGrouping"
End Sub

'-----------------------------------------------------
' Menu_CurrentMonth_Change - Worksheet change event handler
' Call this from Menu sheet's Worksheet_Change event when cell C11 changes
' This allows the year grouping to update automatically when the month changes
'
' To use: Add this code to the Menu worksheet's code module:
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Not Intersect(Target, Range("C11")) Is Nothing Then
'           Call OnCurrentMonthChange
'       End If
'   End Sub
'-----------------------------------------------------
