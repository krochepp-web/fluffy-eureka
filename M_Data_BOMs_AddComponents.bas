Attribute VB_Name = "M_Data_BOMs_AddComponents"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_AddComponents
'
' Purpose:
'   Deprecated compatibility wrapper for legacy add-to-BOM macro entry points.
'   Canonical BOM add flows now live in M_Data_BOMs_Picker:
'     - UI_Add_SelectedPickerRows_To_ActiveBOM
'     - UI_Add_ComponentByPNRev_To_ActiveBOM
'
' Version: v2.0.0
' Author: ChatGPT (assistant)
' Date: 2026-02-20
'===============================================================================

'==========================
' PUBLIC COMPAT ENTRY POINT
'==========================
Public Sub UI_Add_Components_To_BOM()
    On Error GoTo EH

    MsgBox "UI_Add_Components_To_BOM is deprecated." & vbCrLf & _
           "Use picker-based add flows instead." & vbCrLf & vbCrLf & _
           "Routing to UI_Add_ComponentByPNRev_To_ActiveBOM...", _
           vbOKOnly, "Add Components to BOM"

    M_Data_BOMs_Picker.UI_Add_ComponentByPNRev_To_ActiveBOM
    Exit Sub

EH:
    MsgBox "Legacy wrapper failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Add Components to BOM"
End Sub
