Attribute VB_Name = "M_Data_BOMs_Picker"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_Picker
'
' Purpose:
'   Provide a reusable in-sheet component picker UX and route selected components
'   into BOM, PO Lines, or Inventory targets.
'
' Inputs (Tabs/Tables/Headers):
'   - Comps sheet: TBL_COMPS
'       Required headers:
'         CompID, OurPN, OurRev, ComponentDescription, UOM, ComponentNotes, RevStatus
'
'   - Pickers sheet (created if missing):
'       Input cells:
'         B2 = SearchText (contains match)
'         B3 = OurRev filter (optional, exact match)
'         B4 = ActiveOnly (TRUE/FALSE)
'         B5 = MaxResults (numeric)
'       Results table:
'         TBL_PICK_RESULTS with headers:
'           CompID, OurPN, OurRev, Description, UOM, ComponentNotes, RevStatus
'
'   - Targets by context:
'       BOM: active sheet ListObject with BOM headers
'       PO:  POLines.TBL_POLINES
'       INV: Inv.TBL_INV
'
' Outputs / Side effects:
'   - Creates/updates Pickers sheet and results table
'   - Adds selected components to target table by context
'   - Enforces active components and guards duplicate PN+Rev active mappings
'
' Preconditions / Postconditions:
'   - Comps.TBL_COMPS exists and is well-formed
'   - User selects one or more rows in TBL_PICK_RESULTS
'
' Errors & Guards:
'   - No Select/Activate in core write logic
'   - SafeText() prevents Type mismatch on Excel error values
'   - Fail-fast header checks and context/table guards
'
' Version: v1.2.0
' Author: ChatGPT
' Date: 2026-02-15
'===============================================================================

Private Enum PickerTargetContext
    PickerTarget_BOM = 1
    PickerTarget_PO = 2
    PickerTarget_INV = 3
End Enum

'==========================
' Constants (schema contract)
'==========================
Private Const SH_COMPS As String = "Comps"
Private Const LO_COMPS As String = "TBL_COMPS"

Private Const SH_PICKERS As String = "Pickers"
Private Const LO_PICK_RESULTS As String = "TBL_PICK_RESULTS"

Private Const SH_POLINES As String = "POLines"
Private Const LO_POLINES As String = "TBL_POLINES"

Private Const SH_INV As String = "Inv"
Private Const LO_INV As String = "TBL_INV"

Private Const SH_SUPPLIERS As String = "Suppliers"
Private Const LO_SUPPLIERS As String = "TBL_SUPPLIERS"

' Picker input layout
Private Const CELL_SEARCH As String = "B2"
Private Const CELL_REV As String = "B3"
Private Const CELL_ACTIVEONLY As String = "B4"
Private Const CELL_MAXRESULTS As String = "B5"
Private Const CELL_COMPID As String = "B6"
Private Const CELL_SUPPLIER As String = "B7"
Private Const CELL_DESCRIPTION As String = "B8"

' Picker top-left for results table
Private Const RESULTS_TOPLEFT As String = "A10"
Private Const HELPER_COMPID_TOPLEFT As String = "J2"
Private Const HELPER_SUPPLIER_TOPLEFT As String = "K2"
Private Const HELPER_DESCRIPTION_TOPLEFT As String = "L2"
Private Const HELPER_REV_TOPLEFT As String = "M2"

' Default settings
Private Const DEFAULT_ACTIVEONLY As Boolean = True
Private Const DEFAULT_MAXRESULTS As Long = 250

Private Const ACTIVE_LABEL As String = "Active"

'==========================
' Public entry points
'==========================

Public Sub UI_Open_ComponentPicker()
    On Error GoTo EH

    If Not GateReady_Safe(True) Then GoTo CleanExit

    EnsurePickerSheetAndTable ThisWorkbook
    RefreshPickerResults ThisWorkbook
    ThisWorkbook.Worksheets(SH_PICKERS).Activate

CleanExit:
    Exit Sub

EH:
    MsgBox "Picker open/refresh failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

' Optional userform wrapper hook (if a custom UF_ComponentPicker exists).
Public Sub UI_Open_ComponentPicker_Form_Optional()
    On Error GoTo Fallback

    VBA.UserForms.Add("UF_ComponentPicker").Show
    Exit Sub

Fallback:
    Err.Clear
    UI_Open_ComponentPicker
    MsgBox "UF_ComponentPicker was not found. Opened the sheet-based picker instead.", vbInformation, "Component Picker"
End Sub

Public Sub UI_Refresh_PickerResults()
    On Error GoTo EH

    If Not GateReady_Safe(True) Then GoTo CleanExit

    EnsurePickerSheetAndTable ThisWorkbook
    RefreshPickerResults ThisWorkbook
    ThisWorkbook.Worksheets(SH_PICKERS).Activate

CleanExit:
    Exit Sub

EH:
    MsgBox "Picker refresh failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Public Sub UI_Add_SelectedPickerRows_To_ActiveBOM()
    UI_Add_SelectedPickerRows_ToContext PickerTarget_BOM
End Sub

Public Sub UI_Add_SelectedPickerRows_To_POLines()
    UI_Add_SelectedPickerRows_ToContext PickerTarget_PO
End Sub

Public Sub UI_Add_SelectedPickerRows_To_Inventory()
    UI_Add_SelectedPickerRows_ToContext PickerTarget_INV
End Sub

Public Sub UI_Add_ComponentByPNRev_To_ActiveBOM()
    Dim wb As Workbook
    Dim loComps As ListObject
    Dim pn As String
    Dim rev As String
    Dim qtyPer As Double

    On Error GoTo EH

    If Not GateReady_Safe(True) Then Exit Sub

    Set wb = ThisWorkbook
    Set loComps = wb.Worksheets(SH_COMPS).ListObjects(LO_COMPS)

    pn = Trim$(InputBox("Enter component OurPN (blank to cancel).", "Add Component By PN/Rev"))
    If Len(pn) = 0 Then Exit Sub

    rev = Trim$(InputBox("Enter component OurRev (optional)." & vbCrLf & _
                         "Leave blank to choose from active revisions for this PN.", _
                         "Add Component By PN/Rev (" & pn & ")"))

    If Len(rev) = 0 Then
        If Not ResolveRevisionForPN(loComps, pn, rev) Then Exit Sub
    End If

    qtyPer = PromptDouble_Simple("Enter QtyPer (> 0).", "Add Component By PN/Rev", 1#)
    If qtyPer <= 0 Then Exit Sub

    AddComponentToActiveBOM pn, rev, qtyPer
    Exit Sub

EH:
    MsgBox "Add-by-PN/Rev failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Private Function ResolveRevisionForPN(ByVal loComps As ListObject, ByVal pn As String, ByRef revOut As String) As Boolean
    Dim idxPn As Long, idxRev As Long, idxRS As Long
    Dim arrPn As Variant, arrRev As Variant, arrRS As Variant
    Dim i As Long
    Dim key As Variant
    Dim revPrompt As String
    Dim enteredRev As String
    Dim dicRev As Object

    ResolveRevisionForPN = False
    revOut = vbNullString

    If loComps Is Nothing Then Exit Function
    If loComps.DataBodyRange Is Nothing Then Exit Function

    idxPn = GetColIndex(loComps, "OurPN")
    idxRev = GetColIndex(loComps, "OurRev")
    idxRS = GetColIndex(loComps, "RevStatus")
    If idxPn = 0 Or idxRev = 0 Or idxRS = 0 Then Exit Function

    arrPn = ColumnRangeTo2D(loComps.ListColumns(idxPn).DataBodyRange)
    arrRev = ColumnRangeTo2D(loComps.ListColumns(idxRev).DataBodyRange)
    arrRS = ColumnRangeTo2D(loComps.ListColumns(idxRS).DataBodyRange)

    Set dicRev = CreateObject("Scripting.Dictionary")
    dicRev.CompareMode = vbTextCompare

    For i = 1 To UBound(arrPn, 1)
        If StrComp(SafeText(arrPn(i, 1)), pn, vbTextCompare) = 0 And _
           StrComp(SafeText(arrRS(i, 1)), ACTIVE_LABEL, vbTextCompare) = 0 Then
            If Len(SafeText(arrRev(i, 1))) > 0 Then dicRev(SafeText(arrRev(i, 1))) = True
        End If
    Next i

    If dicRev.Count = 0 Then
        MsgBox "No active revisions found for PN: " & pn, vbExclamation, "Component Picker"
        Exit Function
    End If

    If dicRev.Count = 1 Then
        revOut = CStr(dicRev.Keys()(0))
        ResolveRevisionForPN = True
        Exit Function
    End If

    revPrompt = "Multiple active revisions found for PN " & pn & "." & vbCrLf & _
                "Enter one of:" & vbCrLf
    For Each key In dicRev.Keys
        revPrompt = revPrompt & "- " & CStr(key) & vbCrLf
    Next key

    enteredRev = Trim$(InputBox(revPrompt, "Choose Revision"))
    If Len(enteredRev) = 0 Then Exit Function

    If Not dicRev.Exists(enteredRev) Then
        MsgBox "Revision '" & enteredRev & "' is not an active revision for PN " & pn & ".", vbExclamation, "Component Picker"
        Exit Function
    End If

    revOut = enteredRev
    ResolveRevisionForPN = True
End Function

Public Sub AddComponentToActiveBOM(ByVal pn As String, ByVal rev As String, ByVal qtyPer As Double)
    Dim wb As Workbook
    Dim loBom As ListObject
    Dim wsComps As Worksheet
    Dim loComps As ListObject

    Dim bomId As String
    Dim compId As String, desc As String, uom As String, notes As String

    On Error GoTo EH

    If qtyPer <= 0 Then
        MsgBox "QtyPer must be > 0.", vbExclamation, "Component Picker"
        Exit Sub
    End If

    Set wb = ThisWorkbook
    Set loBom = ResolveTargetTable(wb, PickerTarget_BOM)
    Set wsComps = wb.Worksheets(SH_COMPS)
    Set loComps = wsComps.ListObjects(LO_COMPS)

    If Not Comps_LookupActive(loComps, pn, rev, ACTIVE_LABEL, compId, desc, uom, notes) Then Exit Sub

    WriteComponentToTarget loBom, PickerTarget_BOM, compId, pn, rev, desc, uom, qtyPer, notes
    Exit Sub

EH:
    MsgBox "Add component failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

'==========================
' Context-aware add orchestration
'==========================
Private Sub UI_Add_SelectedPickerRows_ToContext(ByVal targetContext As PickerTargetContext)
    Dim wb As Workbook
    Dim wsPick As Worksheet
    Dim loPick As ListObject
    Dim rowIndices As Collection
    Dim qtyDefault As Double
    Dim promptPerRowQty As Boolean

    On Error GoTo EH

    If Not GateReady_Safe(True) Then Exit Sub

    Set wb = ThisWorkbook
    Set wsPick = wb.Worksheets(SH_PICKERS)
    Set loPick = wsPick.ListObjects(LO_PICK_RESULTS)

    If loPick.DataBodyRange Is Nothing Then
        MsgBox "No picker results to add.", vbInformation, "Component Picker"
        Exit Sub
    End If

    Set rowIndices = GetSelectedPickerRowIndices(loPick)
    If rowIndices.Count = 0 Then
        Dim selChoice As VbMsgBoxResult

        selChoice = PromptYesNoCancel( _
            "No picker rows are currently selected." & vbCrLf & _
            "(Tip: if you run this from a BOM sheet, your picker selection is not active.)" & vbCrLf & vbCrLf & _
            "Yes = use ALL rows currently shown in Pickers!" & LO_PICK_RESULTS & vbCrLf & _
            "No = use PN/Rev dialog (BOM only)" & vbCrLf & _
            "Cancel = stop", _
            "Component Picker", vbDefaultButton1)

        Select Case selChoice
            Case vbYes
                Set rowIndices = GetAllPickerRowIndices(loPick)

            Case vbNo
                If targetContext = PickerTarget_BOM Then
                    UI_Add_ComponentByPNRev_To_ActiveBOM
                    Exit Sub
                End If

            Case Else
                Exit Sub
        End Select
    End If

    If rowIndices.Count = 0 Then
        MsgBox "No picker rows were provided to add.", vbExclamation, "Component Picker"
        Exit Sub
    End If

    qtyDefault = PromptDouble_Simple("Enter default quantity (> 0):", "Component Picker", 1#)
    If qtyDefault <= 0 Then Exit Sub

    promptPerRowQty = PromptYesNo("Do you want to enter quantity per selected row?" & vbCrLf & _
                                  "Yes = prompt for each row" & vbCrLf & _
                                  "No = apply default quantity to all rows", _
                                  "Quantity Mode", False)

    AddPickedRowsToTarget wb, loPick, rowIndices, targetContext, qtyDefault, promptPerRowQty

    MsgBox "Selected components processed for " & ContextLabel(targetContext) & ".", vbInformation, "Component Picker"
    Exit Sub

EH:
    MsgBox "Add selected components failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Private Sub AddPickedRowsToTarget(ByVal wb As Workbook, ByVal loPick As ListObject, ByVal rowIndices As Collection, _
                                  ByVal targetContext As PickerTargetContext, ByVal qtyDefault As Double, _
                                  ByVal promptPerRowQty As Boolean)
    Dim loTarget As ListObject
    Dim i As Long
    Dim pickRowIndex As Long

    Dim compId As String, pn As String, rev As String, desc As String, uom As String, notes As String, rs As String
    Dim pickDescHeader As String
    Dim qtyVal As Double

    On Error GoTo EH

    ValidateUniqueActiveMappings wb

    Set loTarget = ResolveTargetTable(wb, targetContext)
    pickDescHeader = PickerResultDescriptionHeader(loPick)

    For i = 1 To rowIndices.Count
        pickRowIndex = CLng(rowIndices(i))

        compId = SafeText(loPick.ListColumns("CompID").DataBodyRange.Cells(pickRowIndex, 1).Value)
        pn = SafeText(loPick.ListColumns("OurPN").DataBodyRange.Cells(pickRowIndex, 1).Value)
        rev = SafeText(loPick.ListColumns("OurRev").DataBodyRange.Cells(pickRowIndex, 1).Value)
        desc = SafeText(loPick.ListColumns(pickDescHeader).DataBodyRange.Cells(pickRowIndex, 1).Value)
        uom = SafeText(loPick.ListColumns("UOM").DataBodyRange.Cells(pickRowIndex, 1).Value)
        notes = SafeText(loPick.ListColumns("ComponentNotes").DataBodyRange.Cells(pickRowIndex, 1).Value)
        rs = SafeText(loPick.ListColumns("RevStatus").DataBodyRange.Cells(pickRowIndex, 1).Value)

        If StrComp(rs, ACTIVE_LABEL, vbTextCompare) <> 0 Then
            MsgBox "Skipping non-active component: " & pn & " / " & rev & " (RevStatus=" & rs & ")", vbExclamation, "Component Picker"
            GoTo NextRow
        End If

        qtyVal = qtyDefault
        If promptPerRowQty Then
            qtyVal = PromptDouble_Simple("Qty for " & pn & " / " & rev & " (> 0):", "Component Picker", qtyDefault)
            If qtyVal <= 0 Then GoTo NextRow
        End If

        WriteComponentToTarget loTarget, targetContext, compId, pn, rev, desc, uom, qtyVal, notes

NextRow:
    Next i

    Exit Sub

EH:
    Err.Raise Err.Number, "AddPickedRowsToTarget", Err.Description
End Sub

Private Function ResolveTargetTable(ByVal wb As Workbook, ByVal targetContext As PickerTargetContext) As ListObject
    Dim ws As Worksheet

    Select Case targetContext
        Case PickerTarget_BOM
            Set ResolveTargetTable = GetActiveBomTable()

        Case PickerTarget_PO
            Set ws = wb.Worksheets(SH_POLINES)
            Set ResolveTargetTable = ws.ListObjects(LO_POLINES)
            RequireColumn ResolveTargetTable, "CompID"
            RequireColumn ResolveTargetTable, "OurPN"
            RequireColumn ResolveTargetTable, "OurRev"
            RequireColumn ResolveTargetTable, "Description"
            RequireColumn ResolveTargetTable, "UOM"
            RequireColumn ResolveTargetTable, "POQuantity"

        Case PickerTarget_INV
            Set ws = wb.Worksheets(SH_INV)
            Set ResolveTargetTable = ws.ListObjects(LO_INV)
            RequireColumn ResolveTargetTable, "CompID"
            RequireColumn ResolveTargetTable, "OurPN"
            RequireColumn ResolveTargetTable, "OurRev"
            RequireColumn ResolveTargetTable, "ComponentDescription"
            RequireColumn ResolveTargetTable, "UOM"
            RequireColumn ResolveTargetTable, "ADD/SUBTRACT"

        Case Else
            Err.Raise vbObjectError + 8601, "ResolveTargetTable", "Unsupported picker target context."
    End Select
End Function

Private Sub WriteComponentToTarget(ByVal loTarget As ListObject, ByVal targetContext As PickerTargetContext, _
                                   ByVal compId As String, ByVal pn As String, ByVal rev As String, _
                                   ByVal desc As String, ByVal uom As String, ByVal qtyVal As Double, ByVal notes As String)
    If qtyVal <= 0 Then Exit Sub

    Select Case targetContext
        Case PickerTarget_BOM
            Bom_UpsertComponent loTarget, compId, pn, rev, desc, uom, qtyVal, notes

        Case PickerTarget_PO
            POLine_AppendComponent loTarget, compId, pn, rev, desc, uom, qtyVal, notes

        Case PickerTarget_INV
            Inv_AppendTransaction loTarget, compId, pn, rev, desc, uom, qtyVal, notes

        Case Else
            Err.Raise vbObjectError + 8602, "WriteComponentToTarget", "Unsupported picker target context."
    End Select
End Sub

Private Function GetSelectedPickerRowIndices(ByVal loPick As ListObject) As Collection
    Dim sel As Range
    Dim selInTable As Range
    Dim area As Range
    Dim rowCell As Range
    Dim dicRows As Object
    Dim key As Variant
    Dim rowIndex As Long

    Set GetSelectedPickerRowIndices = New Collection

    If loPick Is Nothing Then Exit Function
    If loPick.DataBodyRange Is Nothing Then Exit Function

    On Error Resume Next
    Set sel = Selection
    On Error GoTo 0
    If sel Is Nothing Then Exit Function

    If Not sel.Parent Is loPick.Parent Then Exit Function

    Set selInTable = Application.Intersect(sel, loPick.DataBodyRange)
    If selInTable Is Nothing Then Exit Function

    Set dicRows = CreateObject("Scripting.Dictionary")
    dicRows.CompareMode = vbTextCompare

    For Each area In selInTable.Areas
        For Each rowCell In area.Cells
            dicRows(CStr(rowCell.Row)) = True
        Next rowCell
    Next area

    For Each key In dicRows.Keys
        rowIndex = CLng(key) - loPick.DataBodyRange.Row + 1
        If rowIndex >= 1 And rowIndex <= loPick.DataBodyRange.Rows.Count Then
            GetSelectedPickerRowIndices.Add rowIndex
        End If
    Next key
End Function

Private Function GetAllPickerRowIndices(ByVal loPick As ListObject) As Collection
    Dim i As Long

    Set GetAllPickerRowIndices = New Collection

    If loPick Is Nothing Then Exit Function
    If loPick.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To loPick.DataBodyRange.Rows.Count
        GetAllPickerRowIndices.Add i
    Next i
End Function


Private Function ContextLabel(ByVal targetContext As PickerTargetContext) As String
    Select Case targetContext
        Case PickerTarget_BOM: ContextLabel = "BOM"
        Case PickerTarget_PO: ContextLabel = "PO Lines"
        Case PickerTarget_INV: ContextLabel = "Inventory"
        Case Else: ContextLabel = "Unknown"
    End Select
End Function

'==========================
' Core logic
'==========================
Private Sub EnsurePickerSheetAndTable(ByVal wb As Workbook)
    Const PROC_NAME As String = "M_Data_BOMs_Picker.EnsurePickerSheetAndTable"

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rngTopLeft As Range
    Dim headers As Variant
    Dim i As Long

    On Error GoTo EH

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(SH_PICKERS)
    On Error GoTo EH

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SH_PICKERS
    End If

    ws.Range("A1").Value = "Component Picker"
    ws.Range("A2").Value = "Search (Description/Notes/PN/CompID contains)"
    ws.Range("A3").Value = "Revision (optional exact match)"
    ws.Range("A4").Value = "Active only (TRUE/FALSE)"
    ws.Range("A5").Value = "Max results"
    ws.Range("A6").Value = "CompID (optional exact match)"
    ws.Range("A7").Value = "Supplier (optional exact match; dropdown)"
    ws.Range("A8").Value = "Description (optional contains/wildcard)"

    If Len(SafeText(ws.Range(CELL_SEARCH).Value)) = 0 Then ws.Range(CELL_SEARCH).Value = ""
    If Len(SafeText(ws.Range(CELL_REV).Value)) = 0 Then ws.Range(CELL_REV).Value = ""
    If Len(SafeText(ws.Range(CELL_ACTIVEONLY).Value)) = 0 Then ws.Range(CELL_ACTIVEONLY).Value = IIf(DEFAULT_ACTIVEONLY, "TRUE", "FALSE")
    If Len(SafeText(ws.Range(CELL_MAXRESULTS).Value)) = 0 Then ws.Range(CELL_MAXRESULTS).Value = DEFAULT_MAXRESULTS
    If Len(SafeText(ws.Range(CELL_COMPID).Value)) = 0 Then ws.Range(CELL_COMPID).Value = ""
    If Len(SafeText(ws.Range(CELL_SUPPLIER).Value)) = 0 Then ws.Range(CELL_SUPPLIER).Value = ""
    If Len(SafeText(ws.Range(CELL_DESCRIPTION).Value)) = 0 Then ws.Range(CELL_DESCRIPTION).Value = ""

    ws.Range("A9").Value = "Results (filter/sort, select rows, then run add macro):"

    Set lo = Nothing
    On Error Resume Next
    Set lo = ws.ListObjects(LO_PICK_RESULTS)
    On Error GoTo EH

    headers = Array("CompID", "OurPN", "OurRev", "ComponentDescription", "UOM", "ComponentNotes", "RevStatus")

    If lo Is Nothing Then
        Set rngTopLeft = ws.Range(RESULTS_TOPLEFT)

        For i = LBound(headers) To UBound(headers)
            rngTopLeft.Offset(0, i).Value = headers(i)
        Next i

        Dim rngTable As Range
        Set rngTable = ws.Range(rngTopLeft, rngTopLeft.Offset(1, UBound(headers)))

        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngTable, XlListObjectHasHeaders:=xlYes)
        lo.Name = LO_PICK_RESULTS
        lo.TableStyle = "TableStyleLight9"

        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Else
        Dim tableNeedsReset As Boolean
        tableNeedsReset = (lo.ListColumns.Count <> (UBound(headers) - LBound(headers) + 1))

        If Not tableNeedsReset Then
            For i = LBound(headers) To UBound(headers)
                If StrComp(lo.ListColumns(i + 1).Name, CStr(headers(i)), vbTextCompare) <> 0 Then
                    tableNeedsReset = True
                    Exit For
                End If
            Next i
        End If

        If tableNeedsReset Then
            Set lo = ResetPickerResultsTable(ws, lo, headers)
        End If
    End If

    ws.Columns("J:M").EntireColumn.Hidden = True
    RebuildPickerDropdownLists wb, ws

    Exit Sub

EH:
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Private Function ResetPickerResultsTable(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal headers As Variant) As ListObject
    Dim i As Long
    Dim rngTopLeft As Range
    Dim rngTable As Range

    If Not lo Is Nothing Then lo.Delete

    Set rngTopLeft = ws.Range(RESULTS_TOPLEFT)

    For i = LBound(headers) To UBound(headers)
        rngTopLeft.Offset(0, i).Value = headers(i)
    Next i

    Set rngTable = ws.Range(rngTopLeft, rngTopLeft.Offset(1, UBound(headers)))
    Set ResetPickerResultsTable = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngTable, XlListObjectHasHeaders:=xlYes)
    ResetPickerResultsTable.Name = LO_PICK_RESULTS
    ResetPickerResultsTable.TableStyle = "TableStyleLight9"

    If Not ResetPickerResultsTable.DataBodyRange Is Nothing Then ResetPickerResultsTable.DataBodyRange.ClearContents
End Function

Private Function PickerResultDescriptionHeader(ByVal loPick As ListObject) As String
    If GetColIndex(loPick, "Description") > 0 Then
        PickerResultDescriptionHeader = "Description"
    ElseIf GetColIndex(loPick, "ComponentDescription") > 0 Then
        PickerResultDescriptionHeader = "ComponentDescription"
    Else
        Err.Raise vbObjectError + 8605, "PickerResultDescriptionHeader", _
                  "Picker results table is missing Description column."
    End If
End Function

Private Sub RebuildPickerDropdownLists(ByVal wb As Workbook, ByVal wsPick As Worksheet)
    Dim loComps As ListObject
    Dim idxCompId As Long, idxSupplier As Long, idxDesc As Long, idxRev As Long
    Dim arr As Variant
    Dim r As Long
    Dim dicCompId As Object, dicSup As Object, dicDesc As Object, dicRev As Object
    Dim compIdTop As Range, supTop As Range, descTop As Range, revTop As Range
    Dim key As Variant
    Dim outRow As Long

    On Error GoTo EH

    Set loComps = wb.Worksheets(SH_COMPS).ListObjects(LO_COMPS)
    If loComps.DataBodyRange Is Nothing Then Exit Sub

    idxCompId = GetColIndex(loComps, "CompID")
    idxSupplier = GetColIndex(loComps, "SupplierName")
    idxDesc = GetColIndex(loComps, "ComponentDescription")
    idxRev = GetColIndex(loComps, "OurRev")
    If idxCompId = 0 Or idxDesc = 0 Then Exit Sub

    arr = loComps.DataBodyRange.Value
    Set dicCompId = CreateObject("Scripting.Dictionary")
    dicCompId.CompareMode = vbTextCompare
    Set dicSup = CreateObject("Scripting.Dictionary")
    dicSup.CompareMode = vbTextCompare
    Set dicDesc = CreateObject("Scripting.Dictionary")
    dicDesc.CompareMode = vbTextCompare
    Set dicRev = CreateObject("Scripting.Dictionary")
    dicRev.CompareMode = vbTextCompare

    For r = 1 To UBound(arr, 1)
        If Len(SafeText(arr(r, idxCompId))) > 0 Then dicCompId(SafeText(arr(r, idxCompId))) = True
        If idxSupplier > 0 Then
            If Len(SafeText(arr(r, idxSupplier))) > 0 Then dicSup(SafeText(arr(r, idxSupplier))) = True
        End If
        If Len(SafeText(arr(r, idxDesc))) > 0 Then dicDesc(SafeText(arr(r, idxDesc))) = True
        If idxRev > 0 Then
            If Len(SafeText(arr(r, idxRev))) > 0 Then dicRev(SafeText(arr(r, idxRev))) = True
        End If
    Next r

    ' Prefer full supplier catalog when available, so dropdown is not limited to current comps rows.
    Dim loSuppliers As ListObject
    Dim idxSupName As Long
    Dim supArr As Variant

    On Error Resume Next
    Set loSuppliers = wb.Worksheets(SH_SUPPLIERS).ListObjects(LO_SUPPLIERS)
    On Error GoTo EH

    If Not loSuppliers Is Nothing Then
        idxSupName = GetColIndex(loSuppliers, "SupplierName")
        If idxSupName > 0 And Not loSuppliers.DataBodyRange Is Nothing Then
            supArr = ColumnRangeTo2D(loSuppliers.ListColumns(idxSupName).DataBodyRange)
            For r = 1 To UBound(supArr, 1)
                If Len(SafeText(supArr(r, 1))) > 0 Then dicSup(SafeText(supArr(r, 1))) = True
            Next r
        End If
    End If

    wsPick.Range("J1").Value = "CompIDOptions"
    wsPick.Range("K1").Value = "SupplierOptions"
    wsPick.Range("L1").Value = "DescriptionOptions"
    wsPick.Range("M1").Value = "RevisionOptions"
    wsPick.Range("J2:M5000").ClearContents

    Set compIdTop = wsPick.Range(HELPER_COMPID_TOPLEFT)
    outRow = 0
    For Each key In dicCompId.Keys
        outRow = outRow + 1
        compIdTop.Offset(outRow - 1, 0).Value = CStr(key)
    Next key

    Set supTop = wsPick.Range(HELPER_SUPPLIER_TOPLEFT)
    outRow = 0
    For Each key In dicSup.Keys
        outRow = outRow + 1
        supTop.Offset(outRow - 1, 0).Value = CStr(key)
    Next key

    Set descTop = wsPick.Range(HELPER_DESCRIPTION_TOPLEFT)
    outRow = 0
    For Each key In dicDesc.Keys
        outRow = outRow + 1
        descTop.Offset(outRow - 1, 0).Value = CStr(key)
    Next key

    Set revTop = wsPick.Range(HELPER_REV_TOPLEFT)
    outRow = 0
    For Each key In dicRev.Keys
        outRow = outRow + 1
        revTop.Offset(outRow - 1, 0).Value = CStr(key)
    Next key

    ApplyValidationListFromRange wsPick.Range(CELL_COMPID), wsPick.Range("J2:J5000")
    ApplyValidationListFromRange wsPick.Range(CELL_SUPPLIER), wsPick.Range("K2:K5000")
    ApplyValidationListFromRange wsPick.Range(CELL_DESCRIPTION), wsPick.Range("L2:L5000")
    ApplyValidationListFromRange wsPick.Range(CELL_REV), wsPick.Range("M2:M5000")
    ApplyValidationInline wsPick.Range(CELL_ACTIVEONLY), "TRUE,FALSE"
    Exit Sub

EH:
    ' Non-fatal helper; picker remains usable without dropdowns
End Sub

Private Sub ApplyValidationInline(ByVal targetCell As Range, ByVal csvList As String)
    On Error Resume Next
    targetCell.Validation.Delete
    On Error GoTo 0

    targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=csvList
    targetCell.Validation.IgnoreBlank = True
    targetCell.Validation.InCellDropdown = True
End Sub

Private Sub ApplyValidationListFromRange(ByVal targetCell As Range, ByVal listRange As Range)
    On Error Resume Next
    targetCell.Validation.Delete
    On Error GoTo 0

    targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='" & targetCell.Worksheet.Name & "'!" & listRange.Address
    targetCell.Validation.IgnoreBlank = True
    targetCell.Validation.InCellDropdown = True
End Sub

Private Sub RefreshPickerResults(ByVal wb As Workbook)
    Const PROC_NAME As String = "M_Data_BOMs_Picker.RefreshPickerResults"

    Dim wsPick As Worksheet
    Dim loPick As ListObject

    Dim searchText As String
    Dim revFilter As String
    Dim activeOnly As Boolean
    Dim maxResults As Long
    Dim compIdFilter As String
    Dim supplierFilter As String
    Dim descFilter As String

    On Error GoTo EH

    Set wsPick = wb.Worksheets(SH_PICKERS)
    Set loPick = wsPick.ListObjects(LO_PICK_RESULTS)

    searchText = LCase$(Trim$(SafeText(wsPick.Range(CELL_SEARCH).Value)))
    revFilter = Trim$(SafeText(wsPick.Range(CELL_REV).Value))
    activeOnly = ParseBoolDefault(wsPick.Range(CELL_ACTIVEONLY).Value, DEFAULT_ACTIVEONLY)
    maxResults = ParseLongDefault(wsPick.Range(CELL_MAXRESULTS).Value, DEFAULT_MAXRESULTS)
    compIdFilter = Trim$(SafeText(wsPick.Range(CELL_COMPID).Value))
    supplierFilter = Trim$(SafeText(wsPick.Range(CELL_SUPPLIER).Value))
    descFilter = Trim$(SafeText(wsPick.Range(CELL_DESCRIPTION).Value))
    If maxResults < 1 Then maxResults = DEFAULT_MAXRESULTS

    Dim outArr As Variant
    Dim outCount As Long
    outArr = Picker_GetResults(wb, searchText, revFilter, activeOnly, maxResults, outCount, compIdFilter, supplierFilter, descFilter)

    If outCount = 0 Then
        ClearPickResults loPick
    Else
        WritePickResults loPick, outArr, outCount
    End If

    Exit Sub

EH:
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Function Picker_GetResults( _
    ByVal wb As Workbook, _
    ByVal searchText As String, _
    ByVal revFilter As String, _
    ByVal activeOnly As Boolean, _
    ByVal maxResults As Long, _
    ByRef outCount As Long, _
    Optional ByVal compIdFilter As String = "", _
    Optional ByVal supplierFilter As String = "", _
    Optional ByVal descFilter As String = "") As Variant
    Const PROC_NAME As String = "M_Data_BOMs_Picker.Picker_GetResults"

    Dim wsComps As Worksheet
    Dim loComps As ListObject

    Dim idxCompID As Long, idxPn As Long, idxRev As Long, idxDesc As Long, idxUom As Long, idxNotes As Long, idxRS As Long, idxSupplier As Long
    Dim compsArr As Variant
    Dim outArr() As Variant
    Dim r As Long

    On Error GoTo EH

    Set wsComps = wb.Worksheets(SH_COMPS)
    Set loComps = wsComps.ListObjects(LO_COMPS)

    idxCompID = GetColIndex(loComps, "CompID")
    idxPn = GetColIndex(loComps, "OurPN")
    idxRev = GetColIndex(loComps, "OurRev")
    idxDesc = GetColIndex(loComps, "ComponentDescription")
    idxUom = GetColIndex(loComps, "UOM")
    idxNotes = GetColIndex(loComps, "ComponentNotes")
    idxRS = GetColIndex(loComps, "RevStatus")
    idxSupplier = GetColIndex(loComps, "SupplierName")

    If idxCompID = 0 Or idxPn = 0 Or idxRev = 0 Or idxDesc = 0 Or idxUom = 0 Or idxNotes = 0 Or idxRS = 0 Then
        Err.Raise vbObjectError + 8101, PROC_NAME, "Comps.TBL_COMPS missing one or more required headers."
    End If

    outCount = 0
    If loComps.DataBodyRange Is Nothing Then
        Picker_GetResults = Empty
        Exit Function
    End If

    compsArr = loComps.DataBodyRange.Value
    If maxResults < 1 Then maxResults = DEFAULT_MAXRESULTS
    ReDim outArr(1 To maxResults, 1 To 7)

    For r = 1 To UBound(compsArr, 1)
        Dim cId As String, cDesc As String, cNotes As String, cPN As String, cRev As String, cRS As String, cSupplier As String

        cId = SafeText(compsArr(r, idxCompID))
        cDesc = SafeText(compsArr(r, idxDesc))
        cNotes = SafeText(compsArr(r, idxNotes))
        cPN = SafeText(compsArr(r, idxPn))
        cRev = SafeText(compsArr(r, idxRev))
        cRS = SafeText(compsArr(r, idxRS))
        If idxSupplier > 0 Then cSupplier = SafeText(compsArr(r, idxSupplier))

        If activeOnly Then
            If StrComp(cRS, ACTIVE_LABEL, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If Len(revFilter) > 0 Then
            If StrComp(cRev, revFilter, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If Len(compIdFilter) > 0 Then
            If StrComp(cId, compIdFilter, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If Len(supplierFilter) > 0 Then
            If StrComp(cSupplier, supplierFilter, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If Len(descFilter) > 0 Then
            If Not TextMatchesWildcardOrContains(cDesc, descFilter) Then GoTo NextRow
        End If

        If Len(searchText) > 0 Then
            If InStr(1, LCase$(cDesc), searchText, vbTextCompare) = 0 And _
               InStr(1, LCase$(cNotes), searchText, vbTextCompare) = 0 And _
               InStr(1, LCase$(cPN), searchText, vbTextCompare) = 0 And _
               InStr(1, LCase$(cId), searchText, vbTextCompare) = 0 Then
                GoTo NextRow
            End If
        End If

        outCount = outCount + 1
        If outCount > maxResults Then Exit For

        outArr(outCount, 1) = SafeText(compsArr(r, idxCompID))
        outArr(outCount, 2) = cPN
        outArr(outCount, 3) = cRev
        outArr(outCount, 4) = cDesc
        outArr(outCount, 5) = SafeText(compsArr(r, idxUom))
        outArr(outCount, 6) = cNotes
        outArr(outCount, 7) = cRS

NextRow:
    Next r

    If outCount = 0 Then
        Picker_GetResults = Empty
    Else
        Picker_GetResults = Slice2D(outArr, outCount, 7)
    End If
    Exit Function

EH:
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Function

Public Function GetActiveBomTable_Public() As ListObject
    Set GetActiveBomTable_Public = GetActiveBomTable()
End Function

'==========================
' Target writers
'==========================
Private Sub Bom_UpsertComponent(ByVal loBom As ListObject, ByVal compId As String, ByVal pn As String, ByVal rev As String, _
                               ByVal desc As String, ByVal uom As String, ByVal qtyPer As Double, ByVal compNotes As String)
    Dim idxPn As Long, idxRev As Long, idxQty As Long
    Dim arrPn As Variant, arrRev As Variant
    Dim i As Long

    idxPn = GetColIndex(loBom, "OurPN")
    idxRev = GetColIndex(loBom, "OurRev")
    idxQty = GetColIndex(loBom, "QtyPer")
    If idxPn = 0 Or idxRev = 0 Or idxQty = 0 Then Err.Raise vbObjectError + 8400, "Bom_UpsertComponent", "BOM table missing OurPN/OurRev/QtyPer."

    If Not loBom.DataBodyRange Is Nothing Then
        arrPn = ColumnRangeTo2D(loBom.ListColumns(idxPn).DataBodyRange)
        arrRev = ColumnRangeTo2D(loBom.ListColumns(idxRev).DataBodyRange)

        For i = 1 To UBound(arrPn, 1)
            If StrComp(SafeText(arrPn(i, 1)), pn, vbTextCompare) = 0 And _
               StrComp(SafeText(arrRev(i, 1)), rev, vbTextCompare) = 0 Then

                Dim currentQty As Double
                currentQty = 0#
                If IsNumeric(loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).Value) Then
                    currentQty = CDbl(loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).Value)
                End If
                loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).Value = currentQty + qtyPer

                If ColumnExists(loBom, "UpdatedAt") Then loBom.ListColumns(GetColIndex(loBom, "UpdatedAt")).DataBodyRange.Cells(i, 1).Value = Now
                If ColumnExists(loBom, "UpdatedBy") Then loBom.ListColumns(GetColIndex(loBom, "UpdatedBy")).DataBodyRange.Cells(i, 1).Value = GetUserNameSafe()
                Exit Sub
            End If
        Next i
    End If

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

Private Sub POLine_AppendComponent(ByVal loPo As ListObject, ByVal compId As String, ByVal pn As String, ByVal rev As String, _
                                   ByVal desc As String, ByVal uom As String, ByVal qtyVal As Double, ByVal compNotes As String)
    Dim lr As ListRow

    Set lr = loPo.ListRows.Add

    SetByHeader loPo, lr, "CompID", compId
    SetByHeader loPo, lr, "OurPN", pn
    SetByHeader loPo, lr, "OurRev", rev
    SetByHeader loPo, lr, "Description", desc
    SetByHeader loPo, lr, "UOM", uom
    SetByHeader loPo, lr, "POQuantity", qtyVal

    If ColumnExists(loPo, "POLineComment") Then SetByHeader loPo, lr, "POLineComment", compNotes
    If ColumnExists(loPo, "CreatedAt") Then SetByHeader loPo, lr, "CreatedAt", Now
    If ColumnExists(loPo, "CreatedBy") Then SetByHeader loPo, lr, "CreatedBy", GetUserNameSafe()
    If ColumnExists(loPo, "UpdatedAt") Then SetByHeader loPo, lr, "UpdatedAt", Now
    If ColumnExists(loPo, "UpdatedBy") Then SetByHeader loPo, lr, "UpdatedBy", GetUserNameSafe()
End Sub

Private Sub Inv_AppendTransaction(ByVal loInv As ListObject, ByVal compId As String, ByVal pn As String, ByVal rev As String, _
                                  ByVal desc As String, ByVal uom As String, ByVal qtyVal As Double, ByVal compNotes As String)
    Dim lr As ListRow
    Dim signedDelta As Double

    signedDelta = qtyVal

    Set lr = loInv.ListRows.Add

    SetByHeader loInv, lr, "CompID", compId
    SetByHeader loInv, lr, "OurPN", pn
    SetByHeader loInv, lr, "OurRev", rev
    SetByHeader loInv, lr, "ComponentDescription", desc
    SetByHeader loInv, lr, "UOM", uom
    SetByHeader loInv, lr, "ADD/SUBTRACT", signedDelta

    If ColumnExists(loInv, "CreatedAt") Then SetByHeader loInv, lr, "CreatedAt", Now
    If ColumnExists(loInv, "CreatedBy") Then SetByHeader loInv, lr, "CreatedBy", GetUserNameSafe()
    If ColumnExists(loInv, "UpdatedAt") Then SetByHeader loInv, lr, "UpdatedAt", Now
    If ColumnExists(loInv, "UpdatedBy") Then SetByHeader loInv, lr, "UpdatedBy", GetUserNameSafe()

    If ColumnExists(loInv, "ComponentNotes") Then SetByHeader loInv, lr, "ComponentNotes", compNotes
End Sub

'==========================
' Picker table writers
'==========================
Private Sub ClearPickResults(ByVal loPick As ListObject)
    If Not loPick.DataBodyRange Is Nothing Then loPick.DataBodyRange.Delete
End Sub

Private Sub WritePickResults(ByVal loPick As ListObject, ByRef outArr As Variant, ByVal outCount As Long)
    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    If Not loPick.DataBodyRange Is Nothing Then loPick.DataBodyRange.Delete
    If outCount <= 0 Then GoTo CleanExit

    Dim i As Long
    For i = 1 To outCount
        loPick.ListRows.Add
    Next i

    loPick.DataBodyRange.Value = Slice2D(outArr, outCount, 7)

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "WritePickResults", Err.Description
End Sub

Private Function Slice2D(ByRef src As Variant, ByVal rows As Long, ByVal cols As Long) As Variant
    Dim out As Variant
    Dim r As Long, c As Long
    ReDim out(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            out(r, c) = src(r, c)
        Next c
    Next r
    Slice2D = out
End Function

'==========================
' Integrity checks
'==========================
Private Sub ValidateUniqueActiveMappings(ByVal wb As Workbook)
    Const PROC_NAME As String = "ValidateUniqueActiveMappings"

    Dim loComps As ListObject
    Dim idxId As Long, idxPn As Long, idxRev As Long, idxRS As Long
    Dim arr As Variant
    Dim i As Long
    Dim keyPnRev As String
    Dim keyCompId As String
    Dim dicPnRev As Object
    Dim dicCompId As Object

    Set loComps = wb.Worksheets(SH_COMPS).ListObjects(LO_COMPS)
    If loComps.DataBodyRange Is Nothing Then Exit Sub

    idxId = GetColIndex(loComps, "CompID")
    idxPn = GetColIndex(loComps, "OurPN")
    idxRev = GetColIndex(loComps, "OurRev")
    idxRS = GetColIndex(loComps, "RevStatus")

    If idxId = 0 Or idxPn = 0 Or idxRev = 0 Or idxRS = 0 Then
        Err.Raise vbObjectError + 8701, PROC_NAME, "Comps table missing keys for uniqueness validation."
    End If

    arr = loComps.DataBodyRange.Value

    Set dicPnRev = CreateObject("Scripting.Dictionary")
    dicPnRev.CompareMode = vbTextCompare

    Set dicCompId = CreateObject("Scripting.Dictionary")
    dicCompId.CompareMode = vbTextCompare

    For i = 1 To UBound(arr, 1)
        If StrComp(SafeText(arr(i, idxRS)), ACTIVE_LABEL, vbTextCompare) = 0 Then
            keyPnRev = UCase$(SafeText(arr(i, idxPn))) & "|" & UCase$(SafeText(arr(i, idxRev)))
            keyCompId = UCase$(SafeText(arr(i, idxId)))

            If Len(keyPnRev) > 1 Then
                If dicPnRev.Exists(keyPnRev) Then
                    Err.Raise vbObjectError + 8702, PROC_NAME, "Duplicate active PN+Rev found in Comps: " & keyPnRev
                End If
                dicPnRev(keyPnRev) = True
            End If

            If Len(keyCompId) > 0 Then
                If dicCompId.Exists(keyCompId) Then
                    Err.Raise vbObjectError + 8703, PROC_NAME, "Duplicate active CompID found in Comps: " & keyCompId
                End If
                dicCompId(keyCompId) = True
            End If
        End If
    Next i
End Sub

'==========================
' Utilities / guards
'==========================
Private Function GateReady_Safe(Optional ByVal showUserMessage As Boolean = True) As Boolean
    On Error GoTo EH
    GateReady_Safe = M_Core_Gate.Gate_Ready(showUserMessage)
    Exit Function
EH:
    GateReady_Safe = False
End Function

Private Function GetActiveBomTable() As ListObject
    Const PROC_NAME As String = "M_Data_BOMs_Picker.GetActiveBomTable"

    Dim wsBom As Worksheet
    Dim loBom As ListObject

    Set wsBom = ActiveSheet
    If wsBom Is Nothing Then Err.Raise vbObjectError + 8300, PROC_NAME, "No active sheet."
    If wsBom.ListObjects.Count < 1 Then Err.Raise vbObjectError + 8301, PROC_NAME, "Active sheet has no BOM table (ListObject)."
    Set loBom = wsBom.ListObjects(1)

    RequireColumn loBom, "CompID"
    RequireColumn loBom, "OurPN"
    RequireColumn loBom, "OurRev"
    RequireColumn loBom, "Description"
    RequireColumn loBom, "UOM"
    RequireColumn loBom, "QtyPer"
    RequireColumn loBom, "CompNotes"

    Set GetActiveBomTable = loBom
End Function

Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 8500, "RequireColumn", "Missing column '" & header & "' in table '" & lo.Name & "'."
    End If
End Sub

Private Function ColumnExists(ByVal lo As ListObject, ByVal header As String) As Boolean
    ColumnExists = (GetColIndex(lo, header) > 0)
End Function

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
        MsgBox "Comps table has no data.", vbExclamation, "Component Picker"
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

    arrId = ColumnRangeTo2D(loComps.ListColumns(idxId).DataBodyRange)
    arrPn = ColumnRangeTo2D(loComps.ListColumns(idxPn).DataBodyRange)
    arrRev = ColumnRangeTo2D(loComps.ListColumns(idxRev).DataBodyRange)
    arrDesc = ColumnRangeTo2D(loComps.ListColumns(idxDesc).DataBodyRange)
    arrUom = ColumnRangeTo2D(loComps.ListColumns(idxUom).DataBodyRange)
    arrNotes = ColumnRangeTo2D(loComps.ListColumns(idxNotes).DataBodyRange)
    arrRS = ColumnRangeTo2D(loComps.ListColumns(idxRS).DataBodyRange)

    For i = 1 To UBound(arrPn, 1)
        If StrComp(SafeText(arrPn(i, 1)), pn, vbTextCompare) = 0 And _
           StrComp(SafeText(arrRev(i, 1)), rev, vbTextCompare) = 0 Then

            If StrComp(SafeText(arrRS(i, 1)), activeLabel, vbTextCompare) <> 0 Then
                MsgBox "Component is not active: " & pn & " / " & rev, vbExclamation, "Component Picker"
                Exit Function
            End If

            compIdOut = SafeText(arrId(i, 1))
            descOut = SafeText(arrDesc(i, 1))
            uomOut = SafeText(arrUom(i, 1))
            notesOut = SafeText(arrNotes(i, 1))
            Comps_LookupActive = True
            Exit Function
        End If
    Next i

    MsgBox "Component not found in Comps: " & pn & " / " & rev, vbExclamation, "Component Picker"
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Err.Raise vbObjectError + 8501, "SetByHeader", "Missing column '" & header & "' in table '" & lo.Name & "'."
    lr.Range.Cells(1, idx).Value = v
End Sub

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = vbNullString
    ElseIf IsNull(v) Then
        SafeText = vbNullString
    Else
        SafeText = Trim$(CStr(v))
    End If
End Function

Private Function ColumnRangeTo2D(ByVal rng As Range) As Variant
    Dim raw As Variant
    Dim out As Variant

    If rng Is Nothing Then
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = vbNullString
        ColumnRangeTo2D = out
        Exit Function
    End If

    raw = rng.Value

    If IsArray(raw) Then
        ColumnRangeTo2D = raw
    Else
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = raw
        ColumnRangeTo2D = out
    End If
End Function


Private Function TextMatchesWildcardOrContains(ByVal candidate As String, ByVal filterText As String) As Boolean
    Dim normalizedCandidate As String
    Dim normalizedFilter As String

    normalizedCandidate = LCase$(SafeText(candidate))
    normalizedFilter = LCase$(Trim$(SafeText(filterText)))

    If Len(normalizedFilter) = 0 Then
        TextMatchesWildcardOrContains = True
        Exit Function
    End If

    If InStr(1, normalizedFilter, "*", vbBinaryCompare) > 0 Or _
       InStr(1, normalizedFilter, "?", vbBinaryCompare) > 0 Then
        TextMatchesWildcardOrContains = (normalizedCandidate Like normalizedFilter)
    Else
        TextMatchesWildcardOrContains = (InStr(1, normalizedCandidate, normalizedFilter, vbBinaryCompare) > 0)
    End If
End Function

Private Function ParseBoolDefault(ByVal v As Variant, ByVal defaultVal As Boolean) As Boolean
    Dim s As String
    If IsError(v) Or IsNull(v) Then
        ParseBoolDefault = defaultVal
        Exit Function
    End If
    s = UCase$(Trim$(CStr(v)))
    If s = "TRUE" Or s = "YES" Or s = "1" Then
        ParseBoolDefault = True
    ElseIf s = "FALSE" Or s = "NO" Or s = "0" Then
        ParseBoolDefault = False
    Else
        ParseBoolDefault = defaultVal
    End If
End Function

Private Function ParseLongDefault(ByVal v As Variant, ByVal defaultVal As Long) As Long
    If IsError(v) Or IsNull(v) Then
        ParseLongDefault = defaultVal
    ElseIf IsNumeric(v) Then
        ParseLongDefault = CLng(v)
    Else
        ParseLongDefault = defaultVal
    End If
End Function

Private Function PromptDouble_Simple(ByVal prompt As String, ByVal title As String, ByVal defaultVal As Double) As Double
    Dim s As String
    s = Trim$(InputBox(prompt, title, CStr(defaultVal)))
    If Len(s) = 0 Then
        PromptDouble_Simple = -1#
        Exit Function
    End If
    If Not IsNumeric(s) Then
        PromptDouble_Simple = -1#
        Exit Function
    End If
    PromptDouble_Simple = CDbl(s)
End Function

Private Function PromptYesNo(ByVal prompt As String, ByVal title As String, ByVal defaultYes As Boolean) As Boolean
    Dim btn As VbMsgBoxResult
    btn = MsgBox(prompt, vbQuestion + vbYesNo + IIf(defaultYes, vbDefaultButton1, vbDefaultButton2), title)
    PromptYesNo = (btn = vbYes)
End Function

Private Function PromptYesNoCancel(ByVal prompt As String, ByVal title As String, Optional ByVal defaultButton As VbMsgBoxStyle = vbDefaultButton1) As VbMsgBoxResult
    PromptYesNoCancel = MsgBox(prompt, vbQuestion + vbYesNoCancel + defaultButton, title)
End Function


Private Function GetUserNameSafe() As String
    Dim u As String
    u = Trim$(Environ$("Username"))
    If Len(u) = 0 Then u = Application.UserName
    If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
    GetUserNameSafe = u
End Function
