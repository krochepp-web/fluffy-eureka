Attribute VB_Name = "M_UI_Comps"
Option Explicit

'===============================================================================
' Module: M_UI_Comps
'
' Purpose:
'   UI entrypoints for Components actions. Parameterless macros suitable for buttons.
'   Enforces Gate, logs start/end, and calls worker procedures.
'
' Inputs:
'   - Gate: M_Core_Gate.Gate_Ready
'   - Worker: M_Data_Comps_Entry.NewComponent
'
' Outputs / Side effects:
'   - Creates a new component record
'
' Version: v3.5.4
' Author: Keith + GPT
' Date: 2025-12-27
'===============================================================================

Private Const MODULE_VERSION As String = "3.5.4"

Public Sub UI_New_Component()
    Const PROC_NAME As String = "M_UI_Comps.UI_New_Component"

    On Error GoTo EH

    FocusCompsAndSortNewest

    If Not M_Core_Gate.Gate_Ready(False) Then
        M_Core_Logging.LogWarn PROC_NAME, "Blocked by Gate", "ModuleVersion=" & MODULE_VERSION
        Exit Sub
    End If

    Dim attemptedCompId As String
    Dim failureReason As String
    Dim createdOk As Boolean

    M_Core_Logging.LogInfo PROC_NAME, "Start: New Component", "ModuleVersion=" & MODULE_VERSION
    createdOk = M_Data_Comps_Entry.NewComponent_Create(attemptedCompId, failureReason)

    If createdOk Then
        M_Core_Logging.LogInfo PROC_NAME, "End: New Component", _
            "Result=SUCCESS; CompID=" & attemptedCompId & "; ModuleVersion=" & MODULE_VERSION
    Else
        If Len(Trim$(failureReason)) = 0 Then failureReason = "Unknown reason"
        M_Core_Logging.LogWarn PROC_NAME, "New component not created", _
            "Result=FAIL; CompID=" & attemptedCompId & "; Reason=" & failureReason & "; ModuleVersion=" & MODULE_VERSION
    End If
    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "UI_New_Component failed", "Err " & Err.Number & ": " & Err.Description, Err.Number
    GoToLogSheet
    MsgBox "UI_New_Component failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "See Log sheet for details.", vbOKOnly, "New Component"
End Sub


Private Sub FocusCompsAndSortNewest()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sortIdx As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Comps")
    If ws Is Nothing Then Exit Sub
    ws.Activate

    Set lo = ws.ListObjects("TBL_COMPS")
    If lo Is Nothing Then Exit Sub

    sortIdx = GetSortColumnIndex(lo, "CompID")
    If sortIdx <= 0 Then Exit Sub

    lo.Sort.SortFields.Clear
    lo.Sort.SortFields.Add Key:=lo.ListColumns(sortIdx).Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With lo.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Function GetSortColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(headerName), vbTextCompare) = 0 Then
            GetSortColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function


Private Sub GoToLogSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets("Log").Activate
    On Error GoTo 0
End Sub
