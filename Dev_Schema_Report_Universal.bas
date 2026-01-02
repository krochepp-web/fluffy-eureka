Attribute VB_Name = "Dev_Schema_Report_Universal"
Option Explicit

'===========================================================
' Purpose:
'   Universal schema export: outputs one row per table column header as:
'     TAB_NAME, TABLE_NAME, COLUMN_HEADER
'   Does NOT assume tables start at A1 (works with A5 or anywhere).
'   Also creates/updates a "Workbook_Schema" sheet containing the same rows.
'
' Inputs:
'   - All ListObjects in ThisWorkbook, excluding TAB_EXCLUDE
'
' Outputs / Side effects:
'   - CSV file selected via Save-As dialog
'   - Sheet "Workbook_Schema" with table "Tbl_Workbook_Schema"
'
' Preconditions:
'   - Workbook is open; tables exist as ListObjects
'
' Errors & Guards:
'   - Skips sheets with no tables
'   - Safe array growth (ReDim Preserve)
'
' Version: v1.1.0
' Author:  ChatGPT (CUSTOM Tracker)
' Date:    2025-12-19
'===========================================================

Public Sub Export_Tables_Headers_LongCSV_Universal()
    Const PROC_NAME As String = "Export_Tables_Headers_LongCSV_Universal"
    Const TAB_EXCLUDE As String = "SCRIPTSPlan"
    Const OUT_SHEET As String = "Workbook_Schema"
    Const OUT_TABLE As String = "Tbl_Workbook_Schema"

    Dim fPath As Variant, ff As Integer
    Dim ws As Worksheet, lo As ListObject
    Dim headers() As String
    Dim i As Long, rowIdx As Long, maxRows As Long
    Dim schemaRows() As Variant

    On Error GoTo EH

    maxRows = 2000
    ReDim schemaRows(1 To maxRows, 1 To 3)

    schemaRows(1, 1) = "TAB_NAME"
    schemaRows(1, 2) = "TABLE_NAME"
    schemaRows(1, 3) = "COLUMN_HEADER"
    rowIdx = 1

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, TAB_EXCLUDE, vbTextCompare) <> 0 Then
            If ws.ListObjects.Count > 0 Then
                For Each lo In ws.ListObjects
                    headers = GetHeaderArray(lo)

                    If HasAnyHeaders(headers) Then
                        For i = LBound(headers) To UBound(headers)
                            rowIdx = rowIdx + 1
                            EnsureCapacity3Col schemaRows, rowIdx, maxRows

                            schemaRows(rowIdx, 1) = ws.Name
                            schemaRows(rowIdx, 2) = lo.Name
                            schemaRows(rowIdx, 3) = headers(i)
                        Next i
                    Else
                        ' Table exists but has no headers (unusual)
                        rowIdx = rowIdx + 1
                        EnsureCapacity3Col schemaRows, rowIdx, maxRows

                        schemaRows(rowIdx, 1) = ws.Name
                        schemaRows(rowIdx, 2) = lo.Name
                        schemaRows(rowIdx, 3) = vbNullString
                    End If
                Next lo
            End If
        End If
    Next ws

    fPath = Application.GetSaveAsFilename( _
                InitialFileName:="Workbook_Schema_Long.csv", _
                FileFilter:="CSV Files (*.csv), *.csv", _
                TITLE:="Save workbook schema CSV")

    If VarType(fPath) = vbBoolean And fPath = False Then Exit Sub

    ff = FreeFile
    Open CStr(fPath) For Output As #ff
    For i = 1 To rowIdx
        Print #ff, CSVQ(CStr(schemaRows(i, 1))) & "," & CSVQ(CStr(schemaRows(i, 2))) & "," & CSVQ(CStr(schemaRows(i, 3)))
    Next i
    Close #ff

    WriteSchemaToSheet OUT_SHEET, OUT_TABLE, schemaRows, rowIdx

    MsgBox "Saved CSV: " & CStr(fPath) & vbCrLf & _
           "Updated sheet: " & OUT_SHEET, vbInformation, PROC_NAME
    Exit Sub

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    MsgBox "Error in " & PROC_NAME & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbCritical, "Schema Export Failed"
End Sub

'-------------------------
' Helpers
'-------------------------

Private Sub WriteSchemaToSheet(ByVal sheetName As String, ByVal tableName As String, ByRef dataArr() As Variant, ByVal rowCount As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRng As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set dataRng = ws.Range("A1").Resize(rowCount, 3)
    dataRng.value = dataArr

    ' Remove existing tables (best-effort) then add
    On Error Resume Next
    For Each tbl In ws.ListObjects
        tbl.UnlistObject
    Next tbl
    On Error GoTo 0

    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = "TableStyleMedium2"
End Sub

Private Sub EnsureCapacity3Col(ByRef arr() As Variant, ByVal neededRow As Long, ByRef maxRows As Long)
    If neededRow <= maxRows Then Exit Sub

    maxRows = maxRows + 2000
    ReDim Preserve arr(1 To maxRows, 1 To 3)
End Sub

Private Function GetHeaderArray(ByVal lo As ListObject) As String()
    Dim i As Long
    Dim tmp() As String

    If lo.HeaderRowRange Is Nothing Then
        ReDim tmp(0 To -1)
        GetHeaderArray = tmp
        Exit Function
    End If

    ReDim tmp(1 To lo.HeaderRowRange.Columns.Count)
    For i = 1 To lo.HeaderRowRange.Columns.Count
        tmp(i) = Trim$(CStr(lo.HeaderRowRange.Cells(1, i).value))
    Next i

    GetHeaderArray = tmp
End Function

Private Function HasAnyHeaders(ByRef headers() As String) As Boolean
    On Error GoTo EH
    HasAnyHeaders = (UBound(headers) >= LBound(headers))
    Exit Function
EH:
    MsgBox "Error in HasAnyHeaders." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Schema Export"
    HasAnyHeaders = False
End Function

Private Function CSVQ(ByVal s As String) As String
    s = Replace(s, Chr$(34), Chr$(34) & Chr$(34))
    CSVQ = Chr$(34) & s & Chr$(34)
End Function



