Attribute VB_Name = "M_UI_Validation"
Option Explicit

'===========================================================
' Purpose:
'   UI-facing entry points for workbook validation.
'   These appear in Alt+F8 and are safe to bind to buttons.
'
' Entry points:
'   - UI_Run_AllChecks        (primary "Run All Checks" button)
'   - UI_Run_GateCheck        (schema + data gate)
'   - UI_Run_DataIntegrityCheck
'
' Depends on:
'   - Gate_Ready(showUserMessage As Boolean) As Boolean
'   - M_Core_DataIntegrity.Validate_DataIntegrity_All(showUserMessage As Boolean) As Boolean
'
' Version: v1.0.0
'===========================================================

Public Sub UI_Run_AllChecks()
    Const PROC_NAME As String = "UI_Run_AllChecks"
    Const RUN_DATA_INTEGRITY_EVEN_IF_GATE_FAILS As Boolean = True

    Dim gateOk As Boolean
    Dim diOk As Boolean

    On Error GoTo EH

    ' 1) Gate (schema + data per spec)
    gateOk = Gate_Ready(True)

    ' 2) Optional refresh of Data_Check if gate fails early
    diOk = True
    If (gateOk = False And RUN_DATA_INTEGRITY_EVEN_IF_GATE_FAILS) Then
        diOk = M_Core_DataIntegrity.Validate_DataIntegrity_All(False)
    End If

    If gateOk Then
        MsgBox "All Checks: PASS (Workbook ready).", vbInformation, "All Checks"
        SafeLog PROC_NAME, 0, "PASS", "Gate returned True."
    Else
        MsgBox "All Checks: FAIL." & vbCrLf & _
               "Review 'Schema_Check' and/or 'Data_Check'.", vbExclamation, "All Checks"
        SafeLog PROC_NAME, 0, "FAIL", "Gate returned False. DataRefresh=" & CStr(diOk)
    End If

    Exit Sub

EH:
    MsgBox "All checks failed to run." & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, "All Checks"
    SafeLog PROC_NAME, Err.Number, Err.Description, "UI entry point failure."
End Sub

Public Sub UI_Run_GateCheck()
    Const PROC_NAME As String = "UI_Run_GateCheck"

    Dim ok As Boolean
    On Error GoTo EH

    ok = Gate_Ready(True)

    If ok Then
        MsgBox "Workbook Gate: PASS (ready).", vbInformation, "Gate"
        SafeLog PROC_NAME, 0, "PASS", "Gate returned True."
    Else
        MsgBox "Workbook Gate: FAIL (not ready)." & vbCrLf & _
               "Review 'Schema_Check' and/or 'Data_Check'.", vbExclamation, "Gate"
        SafeLog PROC_NAME, 0, "FAIL", "Gate returned False."
    End If

    Exit Sub

EH:
    MsgBox "Gate check failed to run." & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, "Gate"
    SafeLog PROC_NAME, Err.Number, Err.Description, "UI entry point failure."
End Sub

Public Sub UI_Run_DataIntegrityCheck()
    Const PROC_NAME As String = "UI_Run_DataIntegrityCheck"

    Dim ok As Boolean
    On Error GoTo EH

    ok = M_Core_DataIntegrity.Validate_DataIntegrity_All(True)

    If ok Then
        MsgBox "Data Integrity: PASS (0 issues).", vbInformation, "Validation"
        SafeLog PROC_NAME, 0, "PASS", "Data integrity returned True."
    Else
        MsgBox "Data Integrity: FAIL (issues found)." & vbCrLf & _
               "Review the 'Data_Check' sheet.", vbExclamation, "Validation"
        SafeLog PROC_NAME, 0, "FAIL", "Data integrity returned False."
    End If

    Exit Sub

EH:
    MsgBox "Data integrity failed to run." & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, "Validation"
    SafeLog PROC_NAME, Err.Number, Err.Description, "UI entry point failure."
End Sub

'-----------------------------------------------------------
' Helpers
'-----------------------------------------------------------

Private Sub SafeLog(ByVal procName As String, ByVal errNum As Long, ByVal errDesc As String, ByVal details As String)
    On Error Resume Next
    M_Core_Logging.LogEvent procName, errNum, errDesc, details
    On Error GoTo 0
End Sub


