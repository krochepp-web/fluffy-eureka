Attribute VB_Name = "M_Core_Debug"
Option Explicit

'===============================================================================
' Module: M_Core_Debug
'
' Purpose:
'   Centralized runtime error diagnostics for VBA.
'   Shows procedure name, Err.Number, Err.Description, Erl (line number),
'   and optional context. Gives the user an option to break (Stop),
'   resume next, or end execution.
'
' Inputs:
'   - procName: calling procedure name
'   - errObj: Err object
'   - context: optional string key/value dump
'
' Outputs / Side effects:
'   - MsgBox with detailed error info
'   - Optional Stop into debugger
'
' Preconditions / Postconditions:
'   - If you want Erl to be non-zero, the calling procedure must use line numbers.
'
' Errors & Guards:
'   - Safe, no dependencies
'
' Version: v1.0.0
' Author: ChatGPT
' Date: 2026-02-07
'===============================================================================

Public Sub Debug_Report(ByVal procName As String, ByVal errObj As ErrObject, Optional ByVal context As String = "")
    Dim msg As String
    Dim choice As VbMsgBoxResult

    msg = "VBA RUNTIME ERROR" & vbCrLf & _
          "-----------------" & vbCrLf & _
          "Procedure : " & procName & vbCrLf & _
          "Error #   : " & errObj.Number & vbCrLf & _
          "Description: " & errObj.Description & vbCrLf & _
          "Line (Erl): " & CStr(Erl) & vbCrLf

    If Len(context) > 0 Then
        msg = msg & vbCrLf & "Context:" & vbCrLf & context & vbCrLf
    End If

    msg = msg & vbCrLf & _
          "Choose:" & vbCrLf & _
          "Yes    = Break into debugger (Stop)" & vbCrLf & _
          "No     = Continue (Resume Next)" & vbCrLf & _
          "Cancel = Stop execution"

    choice = MsgBox(msg, vbYesNoCancel + vbCritical, "Debug Report")

    Select Case choice
        Case vbYes
            Stop
        Case vbNo
            Resume Next
        Case Else
            End
    End Select
End Sub


