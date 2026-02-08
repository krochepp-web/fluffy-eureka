Attribute VB_Name = "M_Core_Tests"
Option Explicit
'*******************************************************************************
' Module:      M_Core_Tests
' Procedure:   Test_Core_Constants
'
' Purpose:
'   Validate that constants defined in M_Core_Constants are aligned with the
'   actual workbook structure. This includes:
'     - Sheet (tab) name constants (SH_*)
'     - Table (ListObject) name constants (TBL_*)
'     - A small set of key column names on core tables
'
' Inputs (Tabs/Tables/Headers):
'   - Uses ThisWorkbook.Worksheets and ListObjects to discover actual sheets
'     and tables.
'   - Uses constants defined in M_Core_Constants:
'       SH_* sheet name constants
'       TBL_* table name constants
'       COL_* column name constants
'
' Outputs / Side effects:
'   - Creates or clears a sheet named "Core_Tests".
'   - Writes one row per test with:
'       Category   (Sheet, Table, Column)
'       ObjectName (constant name or table.column descriptor)
'       Value      (string value of the constant)
'       Status     ("PASS" / "FAIL")
'       Detail     (human-readable explanation)
'
' Preconditions:
'   - M_Core_Constants module is present and compiled.
'   - Workbook is expected to conform to Schema 3.4.3.
'
' Postconditions:
'   - Does not modify business data; only writes the Core_Tests sheet.
'
' Errors & Guards:
'   - Any unexpected errors are reported via MsgBox and Debug.Print.
'   - Tests that fail are marked as Status = FAIL with a description; the
'     procedure does not stop on first failure.
'
' Version:     v0.1.0
' Author:      ChatGPT (assistant)
' Date:        2025-11-28
'*******************************************************************************

Public Sub Test_Core_Constants()
    Const PROC_NAME As String = "Test_Core_Constants"
    
    Dim wb As Workbook
    Dim wsReport As Worksheet
    Dim dictSheets As Object           ' key: UCase(sheetName), val: True
    Dim dictTables As Object           ' key: UCase(tableName), val: sheetName
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Dim NextRow As Long
    
    On Error GoTo EH
    
    Set wb = ThisWorkbook
    
    '---------------------------------------------------------------------------
    ' Discover actual sheets and tables
    '---------------------------------------------------------------------------
    Set dictSheets = CreateObject("Scripting.Dictionary")
    Set dictTables = CreateObject("Scripting.Dictionary")
    
    For Each ws In wb.Worksheets
        dictSheets(UCase$(ws.Name)) = True
        For Each lo In ws.ListObjects
            ' If duplicate table names ever existed across sheets, last one wins
            dictTables(UCase$(lo.Name)) = ws.Name
        Next lo
    Next ws
    
    '---------------------------------------------------------------------------
    ' Prepare Core_Tests sheet
    '---------------------------------------------------------------------------
    Set wsReport = EnsureReportSheet_Core(wb, "Core_Tests")
    wsReport.Cells.Clear
    
    wsReport.Range("A1").value = "Category"
    wsReport.Range("B1").value = "ObjectName"
    wsReport.Range("C1").value = "Value"
    wsReport.Range("D1").value = "Status"
    wsReport.Range("E1").value = "Detail"
    NextRow = 2
    
    '---------------------------------------------------------------------------
    ' 1) Test Sheet name constants (SH_*)
    '---------------------------------------------------------------------------
    Test_SheetConstant wsReport, NextRow, "SH_LANDING", SH_LANDING, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_SCHEMA", SH_SCHEMA, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_AUTO", SH_AUTO, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_SUPPLIERS", SH_SUPPLIERS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_HELPERS", SH_HELPERS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_LOG", SH_LOG, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_COMPS", SH_COMPS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_USERS", SH_USERS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_INV", SH_INV, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_BOMS", SH_BOMS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_WOS", SH_WOS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_WOCOMPS", SH_WOCOMPS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_DEMAND", SH_DEMAND, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_POS", SH_POS, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_POLINES", SH_POLINES, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_BOM_PDM001", SH_BOM_PDM001, dictSheets
    Test_SheetConstant wsReport, NextRow, "SH_BOM_TEMPLATE", SH_BOM_TEMPLATE, dictSheets
    
    '---------------------------------------------------------------------------
    ' 2) Test Table name constants (TBL_*)
    '---------------------------------------------------------------------------
    Test_TableConstant wsReport, NextRow, "TBL_SCHEMA", TBL_SCHEMA, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_AUTO", TBL_AUTO, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_SUPPLIERS", TBL_SUPPLIERS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_HELPERS", TBL_HELPERS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_LOG", TBL_LOG, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_COMPS", TBL_COMPS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_USERS", TBL_USERS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_INV", TBL_INV, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_BOMS", TBL_BOMS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_WOS", TBL_WOS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_WOCOMPS", TBL_WOCOMPS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_DEMAND", TBL_DEMAND, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_POS", TBL_POS, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_POLINES", TBL_POLINES, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_BOM_PDM001", TBL_BOM_PDM001, dictTables
    Test_TableConstant wsReport, NextRow, "TBL_BOM_TEMPLATE", TBL_BOM_TEMPLATE, dictTables
    
    '---------------------------------------------------------------------------
    ' 3) Spot-check key columns on core tables
    '---------------------------------------------------------------------------
    ' COMPS: CompID, OurPN, OurRev
    Test_TableColumns wsReport, NextRow, _
        "TBL_COMPS core columns", SH_COMPS, TBL_COMPS, _
        Array(COL_COMP_ID, COL_PN, COL_REV)
    
    ' WOS: BuildID, AssemblyID
    Test_TableColumns wsReport, NextRow, _
        "TBL_WOS core columns", SH_WOS, TBL_WOS, _
        Array(COL_BUILD_ID, COL_ASSEMBLY_ID)
    
    ' WOComps: BuildID, OurPN, OurRev, QtyPer
    Test_TableColumns wsReport, NextRow, _
        "TBL_WOCOMPS core columns", SH_WOCOMPS, TBL_WOCOMPS, _
        Array(COL_BUILD_ID, COL_PN, COL_REV, COL_QTY_PER)
    
    ' POLines: POID, POLine, CompID, OurPN, OurRev, POQuantity
    Test_TableColumns wsReport, NextRow, _
        "TBL_POLINES core columns", SH_POLINES, TBL_POLINES, _
        Array(COL_PO_ID, COL_PO_LINE, COL_COMP_ID, COL_PN, COL_REV, COL_PO_QTY)
    
    '---------------------------------------------------------------------------
    ' Autofit and finish
    '---------------------------------------------------------------------------
    wsReport.Columns("A:E").AutoFit
    
    MsgBox "Core constants test complete. See 'Core_Tests' sheet.", vbInformation, PROC_NAME
    
CleanExit:
    Exit Sub
    
EH:
    Debug.Print "Error in " & PROC_NAME & ": " & Err.Number & " - " & Err.Description
    MsgBox "Error " & Err.Number & " in " & PROC_NAME & ": " & Err.Description, vbCritical, PROC_NAME
    Resume CleanExit
End Sub

'*******************************************************************************
' Helper: Ensure a report sheet exists; create if needed
'*******************************************************************************
Private Function EnsureReportSheet_Core(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    
    Set EnsureReportSheet_Core = ws
End Function

'*******************************************************************************
' Helper: Write a single test result row
'*******************************************************************************
Private Sub WriteTestRow( _
    ByVal ws As Worksheet, _
    ByRef NextRow As Long, _
    ByVal category As String, _
    ByVal objectName As String, _
    ByVal value As String, _
    ByVal status As String, _
    ByVal detail As String)
    
    ws.Cells(NextRow, 1).value = category
    ws.Cells(NextRow, 2).value = objectName
    ws.Cells(NextRow, 3).value = value
    ws.Cells(NextRow, 4).value = status
    ws.Cells(NextRow, 5).value = detail
    
    NextRow = NextRow + 1
End Sub

'*******************************************************************************
' Helper: Test a sheet constant
'*******************************************************************************
Private Sub Test_SheetConstant( _
    ByVal wsReport As Worksheet, _
    ByRef NextRow As Long, _
    ByVal constName As String, _
    ByVal sheetName As String, _
    ByVal dictSheets As Object)
    
    Dim key As String
    Dim status As String
    Dim detail As String
    
    key = UCase$(sheetName)
    
    If dictSheets.Exists(key) Then
        status = "PASS"
        detail = "Sheet exists."
    Else
        status = "FAIL"
        detail = "Sheet not found in workbook."
    End If
    
    WriteTestRow wsReport, NextRow, "Sheet", constName, sheetName, status, detail
End Sub

'*******************************************************************************
' Helper: Test a table constant
'*******************************************************************************
Private Sub Test_TableConstant( _
    ByVal wsReport As Worksheet, _
    ByRef NextRow As Long, _
    ByVal constName As String, _
    ByVal tableName As String, _
    ByVal dictTables As Object)
    
    Dim key As String
    Dim status As String
    Dim detail As String
    Dim hostSheet As String
    
    key = UCase$(tableName)
    
    If dictTables.Exists(key) Then
        status = "PASS"
        hostSheet = CStr(dictTables(key))
        detail = "Table exists on sheet '" & hostSheet & "'."
    Else
        status = "FAIL"
        detail = "Table not found in any worksheet."
    End If
    
    WriteTestRow wsReport, NextRow, "Table", constName, tableName, status, detail
End Sub

'*******************************************************************************
' Helper: Test that specified columns exist in a given table
'*******************************************************************************
Private Sub Test_TableColumns( _
    ByVal wsReport As Worksheet, _
    ByRef NextRow As Long, _
    ByVal testName As String, _
    ByVal sheetName As String, _
    ByVal tableName As String, _
    ByVal columnsToCheck As Variant)
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim status As String
    Dim detail As String
    
    Dim colIndex As Long
    Dim i As Long
    Dim colName As String
    Dim found As Boolean
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        WriteTestRow wsReport, NextRow, "Column", testName, "", "FAIL", _
                     "Sheet '" & sheetName & "' not found for table column check."
        Exit Sub
    End If
    
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If lo Is Nothing Then
        WriteTestRow wsReport, NextRow, "Column", testName, "", "FAIL", _
                     "Table '" & tableName & "' not found on sheet '" & sheetName & "'."
        Exit Sub
    End If
    
    ' For each requested column, confirm it exists by name
    For i = LBound(columnsToCheck) To UBound(columnsToCheck)
        colName = CStr(columnsToCheck(i))
        found = False
        
        For colIndex = 1 To lo.ListColumns.Count
            If StrComp(lo.ListColumns(colIndex).Name, colName, vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next colIndex
        
        If found Then
            status = "PASS"
            detail = "Column found in table."
        Else
            status = "FAIL"
            detail = "Column not found in table."
        End If
        
        WriteTestRow wsReport, NextRow, "Column", tableName & "." & colName, colName, status, detail
    Next i
End Sub


