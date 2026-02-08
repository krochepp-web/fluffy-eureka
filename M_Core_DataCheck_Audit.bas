Attribute VB_Name = "M_Core_DataCheck_Audit"
Option Explicit

'===========================================================
' Purpose:
'   Summarize Data_Check issues by Category/Tab/Table so the
'   worst offenders are obvious immediately.
'
' Inputs:
'   - Sheet: Data_Check (headers in row 1)
'     Columns: Category, TabName, TableName
'
' Outputs:
'   - Sheet: Data_Check_Summary (overwritten)
'
' Version: v1.0.0
' Author: ChatGPT
' Date: 2025-12-21
'===========================================================

Public Sub Audit_DataCheck_Summary()
    Const SRC_SHEET As String = "Data_Check"
    Const OUT_SHEET As String = "Data_Check_Summary"

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long
    Dim dic As Object
    Dim r As Long

    On Error GoTo EH
    Set wb = ThisWorkbook

    If Not WorksheetExists(wb, SRC_SHEET) Then
        MsgBox "Missing sheet: " & SRC_SHEET, vbExclamation, "Audit_DataCheck_Summary"
        Exit Sub
    End If
    Set wsSrc = wb.Worksheets(SRC_SHEET)

    lastRow = wsSrc.Cells(wsSrc.rows.Count, 1).End(xlUp).row
    Set wsOut = EnsureWorksheet(wb, OUT_SHEET)
    wsOut.Cells.ClearContents
    wsOut.Range("A1:D1").value = Array("Category", "TabName", "TableName", "Count")

    If lastRow < 2 Then
        wsOut.Columns("A:D").AutoFit
        MsgBox "No issues found (Data_Check has no rows)." & vbCrLf & _
               "Created '" & OUT_SHEET & "' summary (empty).", vbInformation, "Audit_DataCheck_Summary"
        Exit Sub
    End If

    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = vbTextCompare

    For r = 2 To lastRow
        Dim cat As String, tabN As String, tblN As String, key As String
        cat = Trim$(CStr(wsSrc.Cells(r, 1).value))
        tabN = Trim$(CStr(wsSrc.Cells(r, 2).value))
        tblN = Trim$(CStr(wsSrc.Cells(r, 3).value))

        If Len(cat) = 0 And Len(tabN) = 0 And Len(tblN) = 0 Then GoTo NextR

        key = cat & "|" & tabN & "|" & tblN
        If dic.Exists(key) Then
            dic(key) = CLng(dic(key)) + 1
        Else
            dic.Add key, 1
        End If

NextR:
    Next r

    Dim outRow As Long
    outRow = 2

    Dim k As Variant, parts() As String
    For Each k In dic.Keys
        parts = Split(CStr(k), "|")
        wsOut.Cells(outRow, 1).value = parts(0)
        wsOut.Cells(outRow, 2).value = parts(1)
        wsOut.Cells(outRow, 3).value = parts(2)
        wsOut.Cells(outRow, 4).value = CLng(dic(k))
        outRow = outRow + 1
    Next k

    'Sort by Count desc
    wsOut.Range("A1:D" & outRow - 1).Sort Key1:=wsOut.Range("D2"), Order1:=xlDescending, header:=xlYes
    wsOut.Columns("A:D").AutoFit

    MsgBox "Created '" & OUT_SHEET & "' summary.", vbInformation, "Audit_DataCheck_Summary"
    Exit Sub

EH:
    MsgBox "Audit_DataCheck_Summary failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Audit_DataCheck_Summary"
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If WorksheetExists(wb, sheetName) Then
        Set EnsureWorksheet = wb.Worksheets(sheetName)
    Else
        Dim ws As Worksheet
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
        Set EnsureWorksheet = ws
    End If
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

