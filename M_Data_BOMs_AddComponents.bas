Attribute VB_Name = "M_Data_BOMs_AddComponents"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_AddComponents
'
' Purpose:
'   Deprecated compatibility wrapper for legacy add-to-BOM macro entry points.
'   Canonical BOM add flows now live in M_Data_BOMs_Picker:
'     - UI_OP_AddSelectedPickerRowsToActiveBOM
'     - UI_OP_AddComponentByPNRevToActiveBOM
'
' Version: v2.0.0
' Author: ChatGPT (assistant)
' Date: 2026-02-20
'===============================================================================

'==========================
' PUBLIC COMPAT ENTRY POINT
'==========================
Public Sub DEV_LegacyAddComponentsToBOM()
    On Error GoTo EH

    MsgBox "DEV_LegacyAddComponentsToBOM is deprecated." & vbCrLf & _
           "Use picker-based add flows instead." & vbCrLf & vbCrLf & _
           "Routing to UI_OP_AddComponentByPNRevToActiveBOM...", _
           vbOKOnly, "Add Components to BOM"

    M_Data_BOMs_Picker.UI_OP_AddComponentByPNRevToActiveBOM
    Exit Sub

EH:
    MsgBox "Legacy wrapper failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Add Components to BOM"
End Sub
