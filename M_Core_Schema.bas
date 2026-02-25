Attribute VB_Name = "M_Core_Schema"
Option Explicit

'===========================================================
' Purpose:
'   Validate that the workbook schema (tabs/tables/columns) matches
'   what is defined in SCHEMA!TBL_SCHEMA.
'   SCHEMA!TBL_SCHEMA is the authoritative source for schema/header checks.
'
' Inputs:
'   - Sheet: "SCHEMA"
'   - Table: "TBL_SCHEMA"
'
' Outputs / Side effects:
'   - Writes results to "Schema_Check" sheet
'   - Optionally shows MsgBox summary (controlled by showUserMessage)
'
' Preconditions:
'   - TBL_SCHEMA exists and contains required headers
'
' Errors & Guards:
'   - Fails fast with clear message if schema table missing
'
' Version: v3.2.1 (patched for silent mode)
' Author:  CUSTOM Tracker
' Date:    2025-12-19
'===========================================================

Public Function RunSchemaCheck(Optional ByVal showUserMessage As Boolean = True) As Boolean
    Schema_Check showUserMessage
    RunSchemaCheck = (CountSchemaIssues() = 0)
End Function

' Canonical schema checker entry point.
' Source of truth: SCHEMA!TBL_SCHEMA defines workbook schema and expected headers.
Public Sub Schema_Check(Optional ByVal showUserMessage As Boolean = True)
    Call Schema_Validate_All(showUserMessage)
End Sub

' DEPRECATED compatibility shim.
' Use RunSchemaCheck or Schema_Check for all new wiring.
Public Sub ValidateSchema(Optional ByVal Strict As Boolean = False, Optional ByVal showUserMessage As Boolean = True)
    Const PROC_NAME As String = "ValidateSchema"
    On Error GoTo EH

    ' Strict reserved for future use
    Schema_Check showUserMessage

CleanExit:
    Exit Sub

EH:
    Debug.Print "Error in " & PROC_NAME & ": " & Err.Number & " - " & Err.Description
    If showUserMessage Then
        MsgBox "Error " & Err.Number & " in " & PROC_NAME & ": " & Err.Description, vbOKOnly, PROC_NAME
    End If
    Resume CleanExit
End Sub

Private Function CountSchemaIssues() As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets("Schema_Check")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        CountSchemaIssues = 0
    Else
        CountSchemaIssues = Application.WorksheetFunction.CountA(ws.Range("A2:A" & CStr(lastRow)))
    End If
    Exit Function

CleanFail:
    CountSchemaIssues = 999999
End Function
'===============================================================================
' Purpose:
'   Determines which tabs should be excluded from schema "extra table" scanning.
'
' Notes:
'   Schema validation in this workbook is table-centric (ListObjects). This skip
'   list prevents user-generated sheets (e.g., BOM_<TAID>) from forcing schema
'   maintenance for each new top assembly.
'===============================================================================
Private Function ShouldSkipSchemaTab(ByVal tabName As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(tabName))

    ' Existing explicit skip
    If t = "SCRIPTSPLAN" Then
        ShouldSkipSchemaTab = True
        Exit Function
    End If

    ' User-generated BOM sheets/spreadsheets: skip dynamic BOM_* tabs,
    ' but DO NOT skip BOM_TEMPLATE (it is schema-required).
    If t Like "BOM_*" And t <> "BOM_TEMPLATE" Then
        ShouldSkipSchemaTab = True
        Exit Function
    End If

    ' Add other skip patterns here if needed
    'If Left$(t, 4) = "TMP_" Then ShouldSkipSchemaTab = True: Exit Function

    ShouldSkipSchemaTab = False
End Function

Public Sub Schema_Validate_All(Optional ByVal showUserMessage As Boolean = True)
    Const PROC_NAME As String = "Schema_Validate_All"

    Dim wb As Workbook
    Dim wsSchema As Worksheet, wsOut As Worksheet
    Dim loSchema As ListObject
    Dim schemaCols As Object
    Dim rules As Object
    Dim issues As Long

    ' Output columns
    Dim outRow As Long

    On Error GoTo EH

    Set wb = ThisWorkbook

    ' Locate schema table
    Set wsSchema = SafeGetWorksheet(wb, "SCHEMA")
    If wsSchema Is Nothing Then
        If showUserMessage Then
            MsgBox "Missing sheet: SCHEMA", vbOKOnly, PROC_NAME
        End If
        Exit Sub
    End If

    Set loSchema = SafeGetListObject(wsSchema, "TBL_SCHEMA")
    If loSchema Is Nothing Then
        If showUserMessage Then
            MsgBox "Missing table: SCHEMA!TBL_SCHEMA", vbOKOnly, PROC_NAME
        End If
        Exit Sub
    End If

    ' Build schema header map and rules
    Set schemaCols = GetSchemaColumnMap(loSchema)
    Set rules = BuildSchemaRules(loSchema, schemaCols, "SCRIPTSPlan")

    ' Prep output sheet
    Set wsOut = EnsureOutputSheet(wb, "Schema_Check")
    wsOut.Cells.Clear

    ' Header row
    wsOut.Range("A1:E1").value = Array("Category", "TabName", "TableName", "ColumnHeader", "Detail")
    outRow = 2
    issues = 0

    ' Validate schema vs workbook
    issues = issues + ValidateTablesExist(wb, rules, wsOut, outRow)
    issues = issues + ValidateColumnsExist(wb, rules, wsOut, outRow)

    ' Summary
    If showUserMessage And issues > 0 Then
        MsgBox "Schema validation complete. Issues found: " & issues & vbCrLf & _
               "See the 'Schema_Check' sheet for details.", vbOKOnly, PROC_NAME
    End If

    Exit Sub

EH:
    If showUserMessage Then
        MsgBox "Error in " & PROC_NAME & vbCrLf & "Err " & Err.Number & ": " & Err.Description, vbOKOnly, PROC_NAME
    End If
End Sub

'-------------------------
' Core validation routines
'-------------------------

Private Function ValidateTablesExist(ByVal wb As Workbook, ByVal rules As Object, ByVal wsOut As Worksheet, ByRef outRow As Long) As Long
    Dim tableKey As String
    Dim ws As Worksheet
    Dim tabName As String, tableName As String
    Dim tKey As Variant
    Dim countIssues As Long

    countIssues = 0

    ' 1) Extra tables: exist in workbook but not in schema rules
   For Each ws In wb.Worksheets
    If Not ShouldSkipSchemaTab(ws.Name) Then
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            tabName = ws.Name
            tableName = lo.Name

            tableKey = MakeTableKey(tabName, tableName)

            If Not rules.Exists(tableKey) Then
                WriteIssue wsOut, outRow, "ExtraTable", tabName, tableName, vbNullString, _
                           "Table exists in workbook but is not defined in TBL_SCHEMA."
                countIssues = countIssues + 1
            End If
        Next lo
    End If
Next ws


    ' 2) Missing tables: exist in rules but not in workbook
    For Each tKey In rules.Keys
        SplitTableKey CStr(tKey), tabName, tableName
        Set ws = SafeGetWorksheet(wb, tabName)

        If ws Is Nothing Then
            WriteIssue wsOut, outRow, "MissingTab", tabName, tableName, vbNullString, _
                       "Tab defined in TBL_SCHEMA but not found in workbook."
            countIssues = countIssues + 1
        Else
            If SafeGetListObject(ws, tableName) Is Nothing Then
                WriteIssue wsOut, outRow, "MissingTable", tabName, tableName, vbNullString, _
                           "Table defined in TBL_SCHEMA but not found in workbook."
                countIssues = countIssues + 1
            End If
        End If
    Next tKey

    ValidateTablesExist = countIssues
End Function

Private Function ValidateColumnsExist(ByVal wb As Workbook, ByVal rules As Object, ByVal wsOut As Worksheet, ByRef outRow As Long) As Long
    Dim tabName As String, tableName As String
    Dim tKey As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim ruleCols As Object
    Dim countIssues As Long

    countIssues = 0

    ' Missing and extra columns per table
    For Each tKey In rules.Keys
        SplitTableKey CStr(tKey), tabName, tableName
        Set ws = SafeGetWorksheet(wb, tabName)
        If ws Is Nothing Then GoTo NextTable

        Set lo = SafeGetListObject(ws, tableName)
        If lo Is Nothing Then GoTo NextTable

        Set ruleCols = rules(CStr(tKey)) ' dictionary of expected column headers

        ' Missing columns (in schema but not workbook)
        Dim colKey As Variant
        For Each colKey In ruleCols.Keys
            If Not HasListColumn(lo, CStr(colKey)) Then
                WriteIssue wsOut, outRow, "MissingColumn", tabName, tableName, CStr(colKey), _
                           "Column defined in TBL_SCHEMA but not found in workbook table."
                countIssues = countIssues + 1
            End If
        Next colKey

        ' Extra columns (in workbook but not schema)
        Dim lc As ListColumn
        For Each lc In lo.ListColumns
            Dim hdr As String
            hdr = Trim$(CStr(lc.Name))
            If Len(hdr) > 0 Then
                If Not ruleCols.Exists(hdr) Then
                    WriteIssue wsOut, outRow, "ExtraColumn", tabName, tableName, hdr, _
                               "Column exists in workbook table but is not defined in TBL_SCHEMA."
                    countIssues = countIssues + 1
                End If
            End If
        Next lc

NextTable:
    Next tKey

    ValidateColumnsExist = countIssues
End Function

Private Sub WriteIssue(ByVal wsOut As Worksheet, ByRef outRow As Long, ByVal category As String, _
                       ByVal tabName As String, ByVal tableName As String, ByVal columnHeader As String, ByVal detail As String)
    wsOut.Cells(outRow, 1).value = category
    wsOut.Cells(outRow, 2).value = tabName
    wsOut.Cells(outRow, 3).value = tableName
    wsOut.Cells(outRow, 4).value = columnHeader
    wsOut.Cells(outRow, 5).value = detail
    outRow = outRow + 1
End Sub

'-------------------------
' Schema parsing helpers
'-------------------------

Private Function MakeTableKey(ByVal tabName As String, ByVal tableName As String) As String
    MakeTableKey = UCase$(Trim$(tabName)) & "|" & UCase$(Trim$(tableName))
End Function

Private Sub SplitTableKey(ByVal tableKey As String, ByRef tabName As String, ByRef tableName As String)
    Dim parts() As String
    parts = Split(tableKey, "|")
    tabName = parts(0)
    tableName = parts(1)
End Sub

Private Function GetSchemaColumnMap(ByVal loSchema As ListObject) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To loSchema.HeaderRowRange.Columns.Count
        dic(UCase$(Trim$(CStr(loSchema.HeaderRowRange.Cells(1, i).value)))) = i
    Next i

    Set GetSchemaColumnMap = dic
End Function

Private Function ResolveSchemaHeader(ByVal headers As Object, ByVal candidates As Variant) As Long
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        Dim key As String
        key = UCase$(Trim$(CStr(candidates(i))))
        If headers.Exists(key) Then
            ResolveSchemaHeader = CLng(headers(key))
            Exit Function
        End If
    Next i
    ResolveSchemaHeader = 0
End Function

Private Function BuildSchemaRules(ByVal loSchema As ListObject, ByVal schemaCols As Object, _
                                 ByVal tabExclude As String) As Object
    Dim dicTables As Object
    Set dicTables = CreateObject("Scripting.Dictionary")

    Dim idxTab As Long, idxTable As Long, idxCol As Long
    idxTab = ResolveSchemaHeader(schemaCols, Array("TAB_NAME", "TABNAME", "TAB"))
    idxTable = ResolveSchemaHeader(schemaCols, Array("TABLE_NAME", "TABLENAME", "TABLE"))
    idxCol = ResolveSchemaHeader(schemaCols, Array("COLUMN_HEADER", "COLUMNHEADER", "COLUMN"))

    If idxTab = 0 Or idxTable = 0 Or idxCol = 0 Then
        Err.Raise vbObjectError + 100, "BuildSchemaRules", "TBL_SCHEMA missing required headers (TAB_NAME, TABLE_NAME, COLUMN_HEADER)."
    End If

    Dim r As Long
    Dim tabName As String, tableName As String, colHeader As String, tableKey As String

    If loSchema.DataBodyRange Is Nothing Then
        Set BuildSchemaRules = dicTables
        Exit Function
    End If

    For r = 1 To loSchema.DataBodyRange.rows.Count
        tabName = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxTab).value))
        tableName = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxTable).value))
        colHeader = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxCol).value))

        If Len(tabName) > 0 And Len(tableName) > 0 Then
            If StrComp(tabName, tabExclude, vbTextCompare) <> 0 Then
                tableKey = MakeTableKey(tabName, tableName)
                If Not dicTables.Exists(tableKey) Then
                    dicTables.Add tableKey, CreateObject("Scripting.Dictionary")
                End If

                If Len(colHeader) > 0 Then
                    Dim cols As Object
                    Set cols = dicTables(tableKey)
                    If Not cols.Exists(colHeader) Then cols.Add colHeader, True
                End If
            End If
        End If
    Next r

    Set BuildSchemaRules = dicTables
End Function

'-------------------------
' Workbook object helpers
'-------------------------

Private Function EnsureOutputSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = SafeGetWorksheet(wb, sheetName)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureOutputSheet = ws
End Function

Private Function SafeGetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SafeGetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function SafeGetListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set SafeGetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function HasListColumn(ByVal lo As ListObject, ByVal colName As String) As Boolean
    On Error GoTo EH
    Dim lc As ListColumn
    Set lc = lo.ListColumns(colName)
    HasListColumn = True
    Exit Function
EH:
    MsgBox "HasListColumn failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Schema Validation"
    HasListColumn = False
End Function


