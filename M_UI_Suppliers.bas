Attribute VB_Name = "M_UI_Suppliers"
Option Explicit

'===============================================================================
' Module: M_UI_Suppliers
'
' Purpose:
'   UI entrypoints for supplier actions. Parameterless macros suitable for
'   buttons. Enforces Gate, logs start/end, calls worker procedures.
'
' Inputs:
'   - Gate: M_Core_Gate.Gate_Ready
'   - Worker: M_Data_Suppliers_Entry.NewSupplier
'
' Outputs / Side effects:
'   - Creates supplier rows (via worker)
'   - Logs to Log.TBL_LOG
'
' Version: v3.5.0
' Author: Keith + GPT
' Date: 2025-12-27
'===============================================================================

Public Sub UI_New_Supplier()
    Const PROC_NAME As String = "M_UI_Suppliers.UI_New_Supplier"
    Dim ok As Boolean

    On Error GoTo EH

    FocusSuppliersAndSortNewest

    ok = M_Core_Gate.Gate_Ready(False)
    If Not ok Then
        M_Core_Logging.LogWarn PROC_NAME, "Blocked by Gate"
        Exit Sub
    End If

    M_Core_Logging.LogInfo PROC_NAME, "Start: New Supplier"

    M_Data_Suppliers_Entry.NewSupplier

    M_Core_Logging.LogInfo PROC_NAME, "Success: New Supplier"

CleanExit:
    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "Failed: New Supplier", "Err " & Err.Number & ": " & Err.Description, Err.Number
    GoToLogSheet
    MsgBox "UI_New_Supplier failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "See Log sheet for details.", vbOKOnly, "New Supplier"
    Resume CleanExit
End Sub


Private Sub FocusSuppliersAndSortNewest()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sortIdx As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SUPPLIERS)
    If ws Is Nothing Then Exit Sub
    ws.Activate

    Set lo = ws.ListObjects(TBL_SUPPLIERS)
    If lo Is Nothing Then Exit Sub

    sortIdx = GetSortColumnIndex(lo, "SupplierID")
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
