Attribute VB_Name = "M_Core_Gate"
Option Explicit

'===========================================================
' Module: M_Core_Gate
'
' Purpose:
'   Central "Gate" that blocks operational macros when the
'   workbook is not in a valid state.
'
'   Gate checks (contract-driven):
'     1) Runs Schema validator macro (canonical entry point)
'     2) Runs Data Integrity validator macro (canonical entry point)
'     3) Evaluates outputs by counting issue rows in:
'          - Schema_Check (schema issues)
'          - Data_Check   (data integrity issues)
'
' Inputs:
'   - Output sheets produced by validators:
'       Schema_Check (row 1 headers, row 2+ issues)
'       Data_Check   (row 1 headers, row 2+ issues)
'
' Outputs / Side effects:
'   - Returns True when both checks have zero issue rows
'   - Logs pass/fail via M_Core_Logging.LogEvent (best-effort)
'   - Optional user-facing MsgBox summary
'
' Preconditions / Postconditions:
'   - Validators exist and are runnable
'   - Output sheets follow "row 1 headers, row 2+ issues"
'
' Errors & Guards:
'   - If validator entry point name is wrong/missing, Gate fails
'     with clear message identifying the expected macro name.
'
' Version: v1.1.1
' Author: Keith + GPT
' Date: 2025-12-27
'===========================================================

Public Function Gate_Ready(Optional ByVal showUserMessage As Boolean = True) As Boolean
    Const PROC_NAME As String = "Gate_Ready"

    Const SCHEMA_VALIDATOR_PROC As String = "M_Core_Schema.Schema_Validate_All"
    Const DATA_VALIDATOR_PROC As String = "M_Core_DataIntegrity.Validate_DataIntegrity_All"

    Const SCHEMA_OUTPUT_SHEET As String = "Schema_Check"
    Const DATA_OUTPUT_SHEET As String = "Data_Check"

    Dim wb As Workbook
    Dim schemaRan As Boolean, dataRan As Boolean
    Dim schemaIssues As Long, dataIssues As Long
    Dim msg As String
    Dim details As String

    On Error GoTo EH
    Set wb = ThisWorkbook

    '--- Run validators
    schemaRan = RunValidatorProc(SCHEMA_VALIDATOR_PROC, False, msg)
    dataRan = RunValidatorProc(DATA_VALIDATOR_PROC, False, msg)

    '--- Count issues from output sheets (row 2+)
    schemaIssues = CountIssueRows(wb, SCHEMA_OUTPUT_SHEET)
    dataIssues = CountIssueRows(wb, DATA_OUTPUT_SHEET)

    '--- Gate decision
    Gate_Ready = (schemaRan And dataRan And schemaIssues = 0 And dataIssues = 0)

    '--- Logging (best-effort; do not fail Gate if logging fails)
    details = "schemaRan=" & CStr(schemaRan) & _
              "; dataRan=" & CStr(dataRan) & _
              "; schemaIssues=" & CStr(schemaIssues) & _
              "; dataIssues=" & CStr(dataIssues) & _
              "; SchemaProc=" & SCHEMA_VALIDATOR_PROC & _
              "; DataProc=" & DATA_VALIDATOR_PROC

    On Error Resume Next
    If Gate_Ready Then
        M_Core_Logging.LogEvent PROC_NAME, LOG_LEVEL_INFO, "Gate PASS", details, 0
    Else
        M_Core_Logging.LogEvent PROC_NAME, LOG_LEVEL_WARN, "Gate FAIL", details, 0
    End If
    On Error GoTo EH

    '--- User message
    If showUserMessage Then
        If Gate_Ready Then
            If ShouldShowGatePassMessage(wb) Then
                MsgBox "Workbook Gate: PASS", vbInformation, "Gate"
            End If
        Else
            MsgBox "Workbook Gate: FAIL" & vbCrLf & _
                   "Schema validator ran: " & CStr(schemaRan) & vbCrLf & _
                   "Data validator ran: " & CStr(dataRan) & vbCrLf & _
                   "Schema issues: " & CStr(schemaIssues) & vbCrLf & _
                   "Data issues: " & CStr(dataIssues) & vbCrLf & _
                   "See '" & SCHEMA_OUTPUT_SHEET & "' and '" & DATA_OUTPUT_SHEET & "'.", _
                   vbExclamation, "Gate"
        End If
    End If

CleanExit:
    Exit Function

EH:
    On Error Resume Next
    M_Core_Logging.LogEvent PROC_NAME, LOG_LEVEL_ERROR, Err.Description, "Unhandled error in Gate_Ready", Err.Number
    On Error GoTo 0

    If showUserMessage Then
        MsgBox "Gate failed." & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Gate"
    End If

    Gate_Ready = False
    Resume CleanExit
End Function

Public Sub Run_Gate_Check()
    Dim ok As Boolean
    ok = Gate_Ready(True)
End Sub

'==========================
' Helpers
'==========================

Private Function RunValidatorProc(ByVal fullyQualifiedProc As String, ByVal showUserMessage As Boolean, ByRef msg As String) As Boolean
    On Error GoTo EH
    Application.Run fullyQualifiedProc, showUserMessage
    RunValidatorProc = True
    Exit Function
EH:
    msg = "Failed to run validator: " & fullyQualifiedProc & " :: " & Err.Description
    MsgBox "Failed to run validator: " & fullyQualifiedProc & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Gate"
    RunValidatorProc = False
End Function

Private Function CountIssueRows(ByVal wb As Workbook, ByVal sheetName As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    If Not WorksheetExists(wb, sheetName) Then
        CountIssueRows = 999999
        Exit Function
    End If

    Set ws = wb.Worksheets(sheetName)
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

    If lastRow < 2 Then
        CountIssueRows = 0
        Exit Function
    End If

    Set rng = ws.Range("A2:A" & CStr(lastRow))
    CountIssueRows = Application.WorksheetFunction.CountA(rng)
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function




Private Function ShouldShowGatePassMessage(ByVal wb As Workbook) As Boolean
    ' Silent-on-success by default. If Landing has DEV MODE? = TRUE,
    ' show pass message to support verbose diagnostic workflows.
    ShouldShowGatePassMessage = LandingFlagValue(wb, "DEV MODE?", False)
End Function

Private Function LandingFlagValue(ByVal wb As Workbook, ByVal headerName As String, ByVal defaultValue As Boolean) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim idx As Long
    Dim headerCell As Range

    On Error GoTo Fallback

    Set ws = wb.Worksheets("Landing")

    For Each lo In ws.ListObjects
        idx = GetListColumnIndex(lo, headerName)
        If idx > 0 Then
            If Not lo.DataBodyRange Is Nothing Then
                LandingFlagValue = ParseBoolean(lo.ListColumns(idx).DataBodyRange.Cells(1, 1).Value, defaultValue)
                Exit Function
            End If
        End If
    Next lo

Fallback:
    On Error Resume Next
    Set headerCell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If Not headerCell Is Nothing Then
        LandingFlagValue = ParseBoolean(ws.Cells(2, headerCell.Column).Value, defaultValue)
    Else
        LandingFlagValue = defaultValue
    End If
End Function

Private Function GetListColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, headerName, vbTextCompare) = 0 Then
            GetListColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    GetListColumnIndex = 0
End Function

Private Function ParseBoolean(ByVal v As Variant, ByVal defaultValue As Boolean) As Boolean
    Dim s As String

    If IsError(v) Or IsNull(v) Then
        ParseBoolean = defaultValue
        Exit Function
    End If

    s = UCase$(Trim$(CStr(v)))
    Select Case s
        Case "TRUE", "YES", "Y", "1", "ON"
            ParseBoolean = True
        Case "FALSE", "NO", "N", "0", "OFF"
            ParseBoolean = False
        Case Else
            ParseBoolean = defaultValue
    End Select
End Function
