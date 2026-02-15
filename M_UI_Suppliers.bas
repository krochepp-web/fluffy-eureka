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

    ok = M_Core_Gate.Gate_Ready(True)
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
    MsgBox "UI_New_Supplier failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "See Log sheet for details.", vbExclamation, "New Supplier"
    Resume CleanExit
End Sub

