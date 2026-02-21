Attribute VB_Name = "M_Core_HealthCheck"
Option Explicit

'===============================================================================
' Purpose:
'   One-click workbook "health check" runner to confirm the workbook is safe
'   to proceed with feature development.
'
'   NOTE (Excel UI):
'     The Macro dialog (Alt+F8) only lists procedures with *no parameters*.
'     Therefore, a parameterless UI wrapper is provided.
'
'   Runs:
'     1) Schema validation -> expected output sheet: "Schema_Check"
'     2) Data integrity validation -> expected output sheet: "Data_Check"
'     3) Optional gate check -> Gate_Ready (or skip if not present)
'
' Inputs (Tabs/Tables/Headers):
'   - Optional validator entry points (if implemented):
'       M_Core_Schema.Schema_Check
'       M_Core_DataIntegrity.Data_Check
'   - Expected output sheets produced by validators:
'       "Schema_Check" (row 1 headers; row 2+ issues)
'       "Data_Check"   (row 1 headers; row 2+ issues)
'
' Outputs / Side effects:
'   - Updates output sheets, shows summary MsgBox, attempts to log via LogEvent
'
' Preconditions / Postconditions:
'   - Macros enabled; output sheets follow header + issue rows convention
'
' Errors & Guards:
'   - Safe if validators/gate/logging are missing; degrades gracefully
'
' Version: v1.0.1 / v1.0.2
' Author: ChatGPT / update to v1.0.2 by KROCHE
' Date: 2025-12-20 / updated 2025-12-20
'===============================================================================

' Deprecated UI wrapper retained for compatibility.
Public Sub UI_Run_HealthCheck()
    RunDiagnostics True
End Sub

' Canonical optional comprehensive diagnostics runner.
' Includes schema/data execution and summary reporting.
Public Sub RunDiagnostics(Optional ByVal showUserMessage As Boolean = True)
    HealthCheck_RunAll showUserMessage
End Sub

Private Sub HealthCheck_RunAll(Optional ByVal showUserMessage As Boolean = True)

    Const PROC_NAME As String = "HealthCheck_RunAll"

    Dim schemaOk As Boolean
    Dim dataOk As Boolean
    Dim gateOk As Variant ' can be Empty if gate macro missing

    Dim schemaIssues As Long
    Dim dataIssues As Long

    Dim msg As String
    Dim ranSchema As Boolean, ranData As Boolean, ranGate As Boolean

    On Error GoTo EH

    ' 1) Run Schema validator (if present)
    ranSchema = CallOptionalMacro("M_Core_Schema.Schema_Check")
    schemaIssues = CountIssuesOnSheet("Schema_Check")
    schemaOk = (schemaIssues = 0) And ranSchema

    ' 2) Run Data integrity validator (if present)
    ranData = CallOptionalMacro("M_Core_DataIntegrity.Data_Check")
    dataIssues = CountIssuesOnSheet("Data_Check")
    dataOk = (dataIssues = 0) And ranData

    ' 3) Optional Gate check (if present)
    ranGate = CallOptionalMacroWithReturn("M_Core_Gate.RunGateCheck", CBool(showUserMessage), gateOk)

    ' Summary
    msg = "Workbook Health Check" & vbCrLf & String(22, "-") & vbCrLf & _
          "Schema validator ran: " & BoolToYesNo(ranSchema) & vbCrLf & _
          "Schema issues found: " & CStr(schemaIssues) & vbCrLf & vbCrLf & _
          "Data validator ran:   " & BoolToYesNo(ranData) & vbCrLf & _
          "Data issues found:    " & CStr(dataIssues) & vbCrLf & vbCrLf

    If ranGate Then
        msg = msg & "RunGateCheck returned: " & VariantToText(gateOk) & vbCrLf & vbCrLf
    Else
        msg = msg & "RunGateCheck: (not found / not run)" & vbCrLf & vbCrLf
    End If

    msg = msg & "Overall status: " & IIf(schemaOk And dataOk, "PASS (safe to proceed)", "FAIL (fix issues before proceeding)")

    TryLog PROC_NAME, 0, "HealthCheck completed", _
           "SchemaRan=" & BoolToYesNo(ranSchema) & "; SchemaIssues=" & schemaIssues & _
           "; DataRan=" & BoolToYesNo(ranData) & "; DataIssues=" & dataIssues & _
           "; GateRan=" & BoolToYesNo(ranGate) & "; Gate=" & VariantToText(gateOk)

    If showUserMessage Then
        MsgBox msg, IIf(schemaOk And dataOk, vbOKOnly, vbOKOnly), "Health Check"
    End If

CleanExit:
    Exit Sub

EH:
    TryLog PROC_NAME, Err.Number, Err.Description, "Unhandled error in health check."
    If showUserMessage Then
        MsgBox "HealthCheck_RunAll failed." & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Health Check"
    End If
    Resume CleanExit
End Sub

'----------------------------
' Helpers
'----------------------------

Private Function CallOptionalMacro(ByVal macroName As String) As Boolean
    On Error GoTo NotFound
    Application.Run macroName
    CallOptionalMacro = True
    Exit Function
NotFound:
    CallOptionalMacro = False
End Function

Private Function CallOptionalMacroWithReturn(ByVal macroName As String, ByVal arg1 As Variant, ByRef retVal As Variant) As Boolean
    On Error GoTo NotFound
    retVal = Application.Run(macroName, arg1)
    CallOptionalMacroWithReturn = True
    Exit Function
NotFound:
    retVal = Empty
    CallOptionalMacroWithReturn = False
End Function

Private Function CountIssuesOnSheet(ByVal sheetName As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets(sheetName)

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        CountIssuesOnSheet = 0
    Else
        CountIssuesOnSheet = lastRow - 1
    End If
    Exit Function

CleanFail:
    CountIssuesOnSheet = 0
End Function

Private Sub TryLog(ByVal procName As String, ByVal errNum As Long, ByVal errDesc As String, Optional ByVal details As String = vbNullString)
    On Error GoTo NoLog
    Application.Run "M_Core_Logging.LogEvent", procName, errNum, errDesc, details
    Exit Sub
NoLog:
    ' no-op
End Sub

Private Function BoolToYesNo(ByVal b As Boolean) As String
    BoolToYesNo = IIf(b, "Yes", "No")
End Function

Private Function VariantToText(ByVal v As Variant) As String
    If IsEmpty(v) Then
        VariantToText = "(empty)"
    ElseIf IsError(v) Then
        VariantToText = "(error)"
    Else
        VariantToText = CStr(v)
    End If
End Function


