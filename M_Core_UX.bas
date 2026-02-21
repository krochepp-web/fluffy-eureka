Attribute VB_Name = "M_Core_UX"
Option Explicit

'===============================================================================
' Module: M_Core_UX
'
' Purpose:
'   Central UX policy helpers for operational macros.
'   - Success confirmations are suppressed by default.
'   - Success confirmations can be enabled from Landing[DEV MODE?].
'   - Failure messaging remains explicit and points users to Log.
'
' Version: v0.1.0
'===============================================================================

Public Function ShouldShowSuccessMessage(ByVal actionName As String) As Boolean
    Dim wb As Workbook

    On Error GoTo Fallback

    Set wb = ThisWorkbook
    ShouldShowSuccessMessage = LandingFlagValue(wb, "DEV MODE?", False)
    Exit Function

Fallback:
    ShouldShowSuccessMessage = False
End Function

Public Sub ShowFailureMessageWithLogFocus( _
    ByVal procName As String, _
    ByVal uiTitle As String, _
    ByVal userMessage As String, _
    Optional ByVal details As String = "", _
    Optional ByVal errNum As Long = 0, _
    Optional ByVal activityId As String = "")

    Dim msg As String

    On Error Resume Next

    M_Core_Logging.LogEvent procName, LOG_LEVEL_ERROR, userMessage, details, errNum, activityId

    msg = userMessage
    If errNum <> 0 Then
        msg = msg & vbCrLf & "Error " & CStr(errNum) & ": " & ErrDescriptionOrFallback(details)
    End If
    msg = msg & vbCrLf & "See Log sheet for details."

    MsgBox msg, vbExclamation, uiTitle
End Sub

Private Function ErrDescriptionOrFallback(ByVal details As String) As String
    If Len(Trim$(details)) > 0 Then
        ErrDescriptionOrFallback = details
    Else
        ErrDescriptionOrFallback = "Unexpected error"
    End If
End Function

Private Function LandingFlagValue(ByVal wb As Workbook, ByVal headerName As String, ByVal defaultValue As Boolean) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim idx As Long
    Dim headerCell As Range

    On Error GoTo Fallback

    Set ws = wb.Worksheets(SH_LANDING)

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
