Attribute VB_Name = "M_Data_BOMs_AddComponents"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_AddComponents
'
' Purpose:
'   Add components to the active BOM sheet's BOM table by OurPN + OurRev + QtyPer.
'   Only allows (OurPN, OurRev) combinations that are "Active" in Comps.TBL_COMPS
'   based on RevStatus = "Active" (case-insensitive).
'
' Inputs (Tabs/Tables/Headers):
'   - Active BOM sheet: first ListObject on the sheet (the BOM table)
'       Required headers (per schema TBL_BOM_TEMPLATE):
'         CompID, OurPN, OurRev, Description, UOM, QtyPer, CompNotes
'       Optional headers:
'         CreatedAt, CreatedBy, UpdatedAt, UpdatedBy
'
'   - Comps sheet: TBL_COMPS
'       Required headers:
'         CompID, OurPN, OurRev, ComponentDescription, UOM, ComponentNotes, RevStatus
'
' Outputs / Side effects:
'   - Adds or updates rows in the active BOM table
'   - If OurPN+OurRev already exists in BOM table, increases QtyPer deterministically
'
' Preconditions / Postconditions:
'   - ActiveSheet is a BOM sheet containing a ListObject table with BOM headers.
'
' Errors & Guards:
'   - Fails fast on missing tables/headers
'   - Blocks if the component PN+Rev is not found in Comps or RevStatus <> "Active"
'   - Blocks if QtyPer <= 0
'
' Version: v0.1.0
' Author: ChatGPT (assistant)
' Date: 2026-02-07
'===============================================================================

'==========================
' PUBLIC ENTRY POINT
'==========================
Public Sub UI_Add_Components_To_BOM()
    Const PROC_NAME As String = "M_Data_BOMs_AddComponents.UI_Add_Components_To_BOM"

    Const SH_COMPS As String = "Comps"
    Const LO_COMPS As String = "TBL_COMPS"

    Const ACTIVE_REVSTATUS As String = "Active"

    Dim wb As Workbook
    Dim wsBom As Worksheet
    Dim wsComps As Worksheet

    Dim loBom As ListObject
    Dim loComps As ListObject

    Dim pn As String, rev As String
    Dim qtyPer As Double

    On Error GoTo EH

    If Not GateReady_Safe(True) Then Exit Sub

    Set wb = ThisWorkbook
    Set wsBom = ActiveSheet

    If wsBom Is Nothing Then Err.Raise vbObjectError + 7000, PROC_NAME, "No active sheet."
    If wsBom.ListObjects.Count < 1 Then Err.Raise vbObjectError + 7001, PROC_NAME, "Active sheet has no table (ListObject)."

    Set loBom = wsBom.ListObjects(1)

    ' BOM required headers
    RequireColumn loBom, "CompID"
    RequireColumn loBom, "OurPN"
    RequireColumn loBom, "OurRev"
    RequireColumn loBom, "Description"
    RequireColumn loBom, "UOM"
    RequireColumn loBom, "QtyPer"
    RequireColumn loBom, "CompNotes"

    ' Comps table
    Set wsComps = wb.Worksheets(SH_COMPS)
    Set loComps = wsComps.ListObjects(LO_COMPS)

    RequireColumn loComps, "CompID"
    RequireColumn loComps, "OurPN"
    RequireColumn loComps, "OurRev"
    RequireColumn loComps, "ComponentDescription"
    RequireColumn loComps, "UOM"
    RequireColumn loComps, "ComponentNotes"
    RequireColumn loComps, "RevStatus"

    ' Loop: user enters PN, Rev, QtyPer repeatedly
    Do
        pn = Trim$(InputBox("Enter component OurPN (blank to stop).", "Add Components to BOM"))
        If Len(pn) = 0 Then Exit Do

        rev = Trim$(InputBox("Enter component OurRev.", "Add Components to BOM (" & pn & ")"))
        If Len(rev) = 0 Then
            MsgBox "Revision is required.", vbExclamation, "Add Components to BOM"
            GoTo NextLoop
        End If

        qtyPer = Prompt_Double_Simple("Enter QtyPer (> 0).", "Add Components to BOM (" & pn & " / " & rev & ")", 1#)
        If qtyPer <= 0 Then
            MsgBox "QtyPer must be > 0.", vbExclamation, "Add Components to BOM"
            GoTo NextLoop
        End If

        ' Lookup in Comps + enforce Active
        Dim compId As String, desc As String, uom As String, compNotes As String
        If Not Comps_LookupActive(loComps, pn, rev, ACTIVE_REVSTATUS, compId, desc, uom, compNotes) Then
            ' Comps_LookupActive shows a user-friendly message
            GoTo NextLoop
        End If

        ' Upsert into BOM table (add qty if already exists)
        Bom_UpsertComponent loBom, compId, pn, rev, desc, uom, qtyPer, compNotes

NextLoop:
        ' continue
    Loop

    Exit Sub

EH:
    MsgBox "Add components failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Add Components to BOM"
End Sub

'==========================
' Helpers
'==========================
Private Function GateReady_Safe(Optional ByVal showUserMessage As Boolean = True) As Boolean
    On Error GoTo EH
    GateReady_Safe = M_Core_Gate.Gate_Ready(showUserMessage)
    Exit Function
EH:
    MsgBox "Gate check failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Add Components to BOM"
    GateReady_Safe = False
End Function

Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 7200, "RequireColumn", "Missing column '" & header & "' in table '" & lo.Name & "'."
    End If
End Sub

Private Function GetColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, header, vbTextCompare) = 0 Then
            GetColIndex = lc.Index
            Exit Function
        End If
    Next lc
    GetColIndex = 0
End Function

Private Function ColumnExists(ByVal lo As ListObject, ByVal header As String) As Boolean
    ColumnExists = (GetColIndex(lo, header) > 0)
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Err.Raise vbObjectError + 7201, "SetByHeader", "Missing column '" & header & "' in table '" & lo.Name & "'."
    lr.Range.Cells(1, idx).value = v
End Sub

' Simple numeric prompt (no dependency on your other prompt helpers)
Private Function Prompt_Double_Simple(ByVal prompt As String, ByVal title As String, ByVal defaultVal As Double) As Double
    Dim s As String
    s = Trim$(InputBox(prompt, title, CStr(defaultVal)))
    If Len(s) = 0 Then
        Prompt_Double_Simple = -1#
        Exit Function
    End If
    If Not IsNumeric(s) Then
        Prompt_Double_Simple = -1#
        Exit Function
    End If
    Prompt_Double_Simple = CDbl(s)
End Function

' Look up a component by PN+Rev, require RevStatus = activeLabel (case-insensitive).
' Returns False (with MsgBox) if not found or inactive.
Private Function Comps_LookupActive(ByVal loComps As ListObject, ByVal pn As String, ByVal rev As String, ByVal activeLabel As String, _
                                   ByRef compIdOut As String, ByRef descOut As String, ByRef uomOut As String, ByRef notesOut As String) As Boolean
    Dim idxId As Long, idxPn As Long, idxRev As Long, idxDesc As Long, idxUom As Long, idxNotes As Long, idxRS As Long
    Dim arrPn As Variant, arrRev As Variant, arrId As Variant, arrDesc As Variant, arrUom As Variant, arrNotes As Variant, arrRS As Variant
    Dim i As Long

    Comps_LookupActive = False
    compIdOut = vbNullString
    descOut = vbNullString
    uomOut = vbNullString
    notesOut = vbNullString

    If loComps Is Nothing Then Exit Function
    If loComps.DataBodyRange Is Nothing Then
        MsgBox "Comps table has no data.", vbExclamation, "Add Components to BOM"
        Exit Function
    End If

    idxId = GetColIndex(loComps, "CompID")
    idxPn = GetColIndex(loComps, "OurPN")
    idxRev = GetColIndex(loComps, "OurRev")
    idxDesc = GetColIndex(loComps, "ComponentDescription")
    idxUom = GetColIndex(loComps, "UOM")
    idxNotes = GetColIndex(loComps, "ComponentNotes")
    idxRS = GetColIndex(loComps, "RevStatus")

    If idxId = 0 Or idxPn = 0 Or idxRev = 0 Or idxDesc = 0 Or idxUom = 0 Or idxNotes = 0 Or idxRS = 0 Then Exit Function

    arrId = loComps.ListColumns(idxId).DataBodyRange.value
    arrPn = loComps.ListColumns(idxPn).DataBodyRange.value
    arrRev = loComps.ListColumns(idxRev).DataBodyRange.value
    arrDesc = loComps.ListColumns(idxDesc).DataBodyRange.value
    arrUom = loComps.ListColumns(idxUom).DataBodyRange.value
    arrNotes = loComps.ListColumns(idxNotes).DataBodyRange.value
    arrRS = loComps.ListColumns(idxRS).DataBodyRange.value

    For i = 1 To UBound(arrId, 1)
        If StrComp(Trim$(CStr(arrPn(i, 1))), pn, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(arrRev(i, 1))), rev, vbTextCompare) = 0 Then

            If StrComp(Trim$(CStr(arrRS(i, 1))), activeLabel, vbTextCompare) <> 0 Then
                MsgBox "Component exists but is not Active in Comps (RevStatus <> '" & activeLabel & "')." & vbCrLf & _
                       "PN/Rev: " & pn & " / " & rev & vbCrLf & _
                       "RevStatus: " & CStr(arrRS(i, 1)), vbExclamation, "Add Components to BOM"
                Exit Function
            End If

            compIdOut = Trim$(CStr(arrId(i, 1)))
            descOut = Trim$(CStr(arrDesc(i, 1)))
            uomOut = Trim$(CStr(arrUom(i, 1)))
            notesOut = Trim$(CStr(arrNotes(i, 1)))

            Comps_LookupActive = True
            Exit Function
        End If
    Next i

    MsgBox "Component not found in Comps for PN/Rev:" & vbCrLf & _
           pn & " / " & rev, vbExclamation, "Add Components to BOM"
End Function

' Adds a new row if PN/Rev not present; otherwise increases QtyPer on the existing row.
Private Sub Bom_UpsertComponent(ByVal loBom As ListObject, ByVal compId As String, ByVal pn As String, ByVal rev As String, _
                               ByVal desc As String, ByVal uom As String, ByVal qtyPer As Double, ByVal compNotes As String)
    Dim idxPn As Long, idxRev As Long, idxQty As Long
    Dim arrPn As Variant, arrRev As Variant
    Dim i As Long

    idxPn = GetColIndex(loBom, "OurPN")
    idxRev = GetColIndex(loBom, "OurRev")
    idxQty = GetColIndex(loBom, "QtyPer")

    If idxPn = 0 Or idxRev = 0 Or idxQty = 0 Then Err.Raise vbObjectError + 7300, "Bom_UpsertComponent", "BOM table missing OurPN/OurRev/QtyPer."

    ' Check for existing row
    If Not loBom.DataBodyRange Is Nothing Then
        arrPn = loBom.ListColumns(idxPn).DataBodyRange.value
        arrRev = loBom.ListColumns(idxRev).DataBodyRange.value

        For i = 1 To UBound(arrPn, 1)
            If StrComp(Trim$(CStr(arrPn(i, 1))), pn, vbTextCompare) = 0 And _
               StrComp(Trim$(CStr(arrRev(i, 1))), rev, vbTextCompare) = 0 Then

                ' Update qty (increase)
                Dim currentQty As Double
                currentQty = 0#
                If IsNumeric(loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).value) Then
                    currentQty = CDbl(loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).value)
                End If
                loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).value = currentQty + qtyPer

                ' Touch Updated fields if present
                If ColumnExists(loBom, "UpdatedAt") Then loBom.ListColumns(GetColIndex(loBom, "UpdatedAt")).DataBodyRange.Cells(i, 1).value = Now
                If ColumnExists(loBom, "UpdatedBy") Then loBom.ListColumns(GetColIndex(loBom, "UpdatedBy")).DataBodyRange.Cells(i, 1).value = GetUserNameSafe()

                Exit Sub
            End If
        Next i
    End If

    ' Add a new row
    Dim lr As ListRow
    Set lr = loBom.ListRows.Add

    SetByHeader loBom, lr, "CompID", compId
    SetByHeader loBom, lr, "OurPN", pn
    SetByHeader loBom, lr, "OurRev", rev
    SetByHeader loBom, lr, "Description", desc
    SetByHeader loBom, lr, "UOM", uom
    SetByHeader loBom, lr, "QtyPer", qtyPer
    SetByHeader loBom, lr, "CompNotes", compNotes

    If ColumnExists(loBom, "CreatedAt") Then SetByHeader loBom, lr, "CreatedAt", Now
    If ColumnExists(loBom, "CreatedBy") Then SetByHeader loBom, lr, "CreatedBy", GetUserNameSafe()
    If ColumnExists(loBom, "UpdatedAt") Then SetByHeader loBom, lr, "UpdatedAt", Now
    If ColumnExists(loBom, "UpdatedBy") Then SetByHeader loBom, lr, "UpdatedBy", GetUserNameSafe()
End Sub

Private Function GetUserNameSafe() As String
    Dim u As String
    u = Trim$(Environ$("Username"))
    If Len(u) = 0 Then u = Application.userName
    If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
    GetUserNameSafe = u
End Function


