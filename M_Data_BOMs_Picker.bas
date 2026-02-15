Attribute VB_Name = "M_Data_BOMs_Picker"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_Picker
'
' Purpose:
'   Provide an in-sheet "picker" UX to search Components by
'   description (and optionally revision) and add selected rows to the active BOM sheet.
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
'           CompID, OurPN, OurRev, ComponentDescription, UOM, ComponentNotes, RevStatus
'
'   - Active BOM sheet:
'       Uses the first ListObject on the active sheet as the BOM table.
'       Required headers:
'         CompID, OurPN, OurRev, Description, UOM, QtyPer, CompNotes
'       Optional headers:
'         CreatedAt, CreatedBy, UpdatedAt, UpdatedBy
'
' Outputs / Side effects:
'   - Creates/updates Pickers sheet and results table
'   - Adds selected components from picker to active BOM table
'   - If component already exists in BOM (PN+Rev), increases QtyPer deterministically
'
' Preconditions / Postconditions:
'   - Comps.TBL_COMPS exists and is well-formed
'   - User selects one or more rows in TBL_PICK_RESULTS before adding
'
' Errors & Guards:
'   - No Select/Activate used
'   - SafeText() prevents Type mismatch on Excel error values
'   - Clear error messages for missing tables/headers
'
' Version: v1.0.0
' Author: ChatGPT
' Date: 2026-02-07
'===============================================================================

'==========================
' Constants (schema contract)
'==========================
Private Const SH_COMPS As String = "Comps"
Private Const LO_COMPS As String = "TBL_COMPS"

Private Const SH_PICKERS As String = "Pickers"
Private Const LO_PICK_RESULTS As String = "TBL_PICK_RESULTS"

' Picker input layout
Private Const CELL_SEARCH As String = "B2"
Private Const CELL_REV As String = "B3"
Private Const CELL_ACTIVEONLY As String = "B4"
Private Const CELL_MAXRESULTS As String = "B5"

' Picker top-left for results table
Private Const RESULTS_TOPLEFT As String = "A8"

' Default settings
Private Const DEFAULT_ACTIVEONLY As Boolean = True
Private Const DEFAULT_MAXRESULTS As Long = 250

Private Const ACTIVE_LABEL As String = "Active"

'==========================
' Public entry points
'==========================

Public Sub UI_Open_ComponentPicker()
    Const PROC_NAME As String = "M_Data_BOMs_Picker.UI_Open_ComponentPicker"
    On Error GoTo EH

    If Not GateReady_Safe(True) Then GoTo CleanExit

    EnsurePickerSheetAndTable ThisWorkbook
    RefreshPickerResults ThisWorkbook

CleanExit:
    Exit Sub

EH:
    MsgBox "Picker open/refresh failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Public Sub UI_Refresh_PickerResults()
    Const PROC_NAME As String = "M_Data_BOMs_Picker.UI_Refresh_PickerResults"
    On Error GoTo EH

    If Not GateReady_Safe(True) Then GoTo CleanExit

    EnsurePickerSheetAndTable ThisWorkbook
    RefreshPickerResults ThisWorkbook

CleanExit:
    Exit Sub

EH:
    MsgBox "Picker refresh failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Public Sub UI_Add_SelectedPickerRows_To_ActiveBOM()
    Const PROC_NAME As String = "M_Data_BOMs_Picker.UI_Add_SelectedPickerRows_To_ActiveBOM"
    On Error GoTo EH

    If Not GateReady_Safe(True) Then GoTo CleanExit

    AddSelectedPickerRowsToActiveBOM ThisWorkbook

CleanExit:
    Exit Sub

EH:
    MsgBox "Add selected components failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Public Sub AddComponentToActiveBOM(ByVal pn As String, ByVal rev As String, ByVal qtyPer As Double)
    Const PROC_NAME As String = "M_Data_BOMs_Picker.AddComponentToActiveBOM"

    Dim wb As Workbook
    Dim loBom As ListObject
    Dim wsComps As Worksheet
    Dim loComps As ListObject

    Dim compId As String, desc As String, uom As String, notes As String

    On Error GoTo EH

    If qtyPer <= 0 Then
        MsgBox "QtyPer must be > 0.", vbExclamation, "Component Picker"
        Exit Sub
    End If

    Set wb = ThisWorkbook
    Set loBom = GetActiveBomTable()
    Set wsComps = wb.Worksheets(SH_COMPS)
    Set loComps = wsComps.ListObjects(LO_COMPS)

    If Not Comps_LookupActive(loComps, pn, rev, ACTIVE_LABEL, compId, desc, uom, notes) Then
        Exit Sub
    End If

    Bom_UpsertComponent loBom, compId, pn, rev, desc, uom, qtyPer, notes
    Exit Sub

EH:
    MsgBox "Add component failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, PROC_NAME
End Sub

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

    ' Layout labels + defaults (idempotent)
    ws.Range("A1").value = "Component Picker"
    ws.Range("A2").value = "Search (Description contains)"
    ws.Range("A3").value = "Revision (optional exact match)"
    ws.Range("A4").value = "Active only (TRUE/FALSE)"
    ws.Range("A5").value = "Max results"

    If Len(SafeText(ws.Range(CELL_SEARCH).value)) = 0 Then ws.Range(CELL_SEARCH).value = ""
    If Len(SafeText(ws.Range(CELL_REV).value)) = 0 Then ws.Range(CELL_REV).value = ""
    If Len(SafeText(ws.Range(CELL_ACTIVEONLY).value)) = 0 Then ws.Range(CELL_ACTIVEONLY).value = IIf(DEFAULT_ACTIVEONLY, "TRUE", "FALSE")
    If Len(SafeText(ws.Range(CELL_MAXRESULTS).value)) = 0 Then ws.Range(CELL_MAXRESULTS).value = DEFAULT_MAXRESULTS

    ws.Range("A7").value = "Results (filter/sort normally, then select rows and run Add Selected):"

    ' Ensure results table exists
    Set lo = Nothing
    On Error Resume Next
    Set lo = ws.ListObjects(LO_PICK_RESULTS)
    On Error GoTo EH

    headers = Array("CompID", "OurPN", "OurRev", "ComponentDescription", "UOM", "ComponentNotes", "RevStatus")

    If lo Is Nothing Then
        Set rngTopLeft = ws.Range(RESULTS_TOPLEFT)

        ' Write headers
        For i = LBound(headers) To UBound(headers)
            rngTopLeft.Offset(0, i).value = headers(i)
        Next i

        ' Create a 2-row table (headers + one blank row)
        Dim rngTable As Range
        Set rngTable = ws.Range(rngTopLeft, rngTopLeft.Offset(1, UBound(headers)))

        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngTable, XlListObjectHasHeaders:=xlYes)
        lo.Name = LO_PICK_RESULTS
        lo.TableStyle = "TableStyleLight9"

        ' Clear the single blank data row
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Else
        ' Validate headers match (fail fast if user edited them)
        For i = LBound(headers) To UBound(headers)
            If i + 1 > lo.ListColumns.Count Then
                Err.Raise vbObjectError + 8001, PROC_NAME, "Picker table header mismatch. Recreate TBL_PICK_RESULTS."
            End If
            If StrComp(lo.ListColumns(i + 1).Name, CStr(headers(i)), vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 8002, PROC_NAME, "Picker table header mismatch. Expected '" & headers(i) & "'. Found '" & lo.ListColumns(i + 1).Name & "'."
            End If
        Next i
    End If

    Exit Sub

EH:
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Private Sub RefreshPickerResults(ByVal wb As Workbook)
    Const PROC_NAME As String = "M_Data_BOMs_Picker.RefreshPickerResults"

    Dim wsPick As Worksheet
    Dim loPick As ListObject

    Dim searchText As String
    Dim revFilter As String
    Dim activeOnly As Boolean
    Dim maxResults As Long

    On Error GoTo EH

    Set wsPick = wb.Worksheets(SH_PICKERS)
    Set loPick = wsPick.ListObjects(LO_PICK_RESULTS)

    searchText = LCase$(Trim$(SafeText(wsPick.Range(CELL_SEARCH).value)))
    revFilter = Trim$(SafeText(wsPick.Range(CELL_REV).value))
    activeOnly = ParseBoolDefault(wsPick.Range(CELL_ACTIVEONLY).value, DEFAULT_ACTIVEONLY)
    maxResults = ParseLongDefault(wsPick.Range(CELL_MAXRESULTS).value, DEFAULT_MAXRESULTS)
    If maxResults < 1 Then maxResults = DEFAULT_MAXRESULTS

    Dim outArr As Variant
    Dim outCount As Long
    outArr = Picker_GetResults(wb, searchText, revFilter, activeOnly, maxResults, outCount)

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
    ByRef outCount As Long) As Variant
    Const PROC_NAME As String = "M_Data_BOMs_Picker.Picker_GetResults"

    Dim wsComps As Worksheet
    Dim loComps As ListObject

    Dim idxCompID As Long, idxPn As Long, idxRev As Long, idxDesc As Long, idxUom As Long, idxNotes As Long, idxRS As Long
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

    If idxCompID = 0 Or idxPn = 0 Or idxRev = 0 Or idxDesc = 0 Or idxUom = 0 Or idxNotes = 0 Or idxRS = 0 Then
        Err.Raise vbObjectError + 8101, PROC_NAME, "Comps.TBL_COMPS missing one or more required headers."
    End If

    outCount = 0
    If loComps.DataBodyRange Is Nothing Then
        Picker_GetResults = Empty
        Exit Function
    End If

    compsArr = loComps.DataBodyRange.value
    If maxResults < 1 Then maxResults = DEFAULT_MAXRESULTS
    ReDim outArr(1 To maxResults, 1 To 7)

    For r = 1 To UBound(compsArr, 1)
        Dim cDesc As String, cNotes As String, cPN As String, cRev As String, cRS As String

        cDesc = SafeText(compsArr(r, idxDesc))
        cNotes = SafeText(compsArr(r, idxNotes))
        cPN = SafeText(compsArr(r, idxPn))
        cRev = SafeText(compsArr(r, idxRev))
        cRS = SafeText(compsArr(r, idxRS))

        If activeOnly Then
            If StrComp(cRS, ACTIVE_LABEL, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If Len(revFilter) > 0 Then
            If StrComp(cRev, revFilter, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If Len(searchText) > 0 Then
            If InStr(1, LCase$(cDesc), searchText, vbTextCompare) = 0 And _
               InStr(1, LCase$(cNotes), searchText, vbTextCompare) = 0 And _
               InStr(1, LCase$(cPN), searchText, vbTextCompare) = 0 Then
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

Private Sub AddSelectedPickerRowsToActiveBOM(ByVal wb As Workbook)
    Const PROC_NAME As String = "M_Data_BOMs_Picker.AddSelectedPickerRowsToActiveBOM"

    Dim wsPick As Worksheet
    Dim loPick As ListObject
    Dim sel As Range
    Dim selInTable As Range
    Dim area As Range
    Dim rowCell As Range

    Dim loBom As ListObject

    Dim qtyPer As Double
    qtyPer = PromptDouble_Simple("Enter QtyPer (> 0) to apply to EACH selected component:", "Add Selected Components", 1#)
    If qtyPer <= 0 Then Exit Sub

    On Error GoTo EH

    Set wsPick = wb.Worksheets(SH_PICKERS)
    Set loPick = wsPick.ListObjects(LO_PICK_RESULTS)

    If loPick.DataBodyRange Is Nothing Then
        MsgBox "No picker results to add.", vbInformation, "Add Selected Components"
        Exit Sub
    End If

    Set sel = Selection
    If sel Is Nothing Then
        MsgBox "Select one or more rows in the picker results table first.", vbExclamation, "Add Selected Components"
        Exit Sub
    End If

    Set selInTable = Intersect(sel, loPick.DataBodyRange)
    If selInTable Is Nothing Then
        MsgBox "Your selection is not inside the picker results table (data rows).", vbExclamation, "Add Selected Components"
        Exit Sub
    End If

    ' Active BOM is the active sheet at run time
    Set loBom = GetActiveBomTable()

    ' Iterate distinct rows in selection (by row index)
    Dim dicRows As Object
    Set dicRows = CreateObject("Scripting.Dictionary")
    dicRows.compareMode = vbTextCompare

    For Each area In selInTable.Areas
        For Each rowCell In area.Cells
            dicRows(CStr(rowCell.row)) = True
        Next rowCell
    Next area

    Dim key As Variant
    For Each key In dicRows.Keys
        Dim pickRowIndex As Long
        pickRowIndex = CLng(key) - loPick.DataBodyRange.row + 1 ' 1-based within DataBodyRange

        If pickRowIndex >= 1 And pickRowIndex <= loPick.DataBodyRange.rows.Count Then
            Dim compId As String, pn As String, rev As String, desc As String, uom As String, notes As String, rs As String

            compId = SafeText(loPick.ListColumns("CompID").DataBodyRange.Cells(pickRowIndex, 1).value)
            pn = SafeText(loPick.ListColumns("OurPN").DataBodyRange.Cells(pickRowIndex, 1).value)
            rev = SafeText(loPick.ListColumns("OurRev").DataBodyRange.Cells(pickRowIndex, 1).value)
            desc = SafeText(loPick.ListColumns("ComponentDescription").DataBodyRange.Cells(pickRowIndex, 1).value)
            uom = SafeText(loPick.ListColumns("UOM").DataBodyRange.Cells(pickRowIndex, 1).value)
            notes = SafeText(loPick.ListColumns("ComponentNotes").DataBodyRange.Cells(pickRowIndex, 1).value)
            rs = SafeText(loPick.ListColumns("RevStatus").DataBodyRange.Cells(pickRowIndex, 1).value)

            Picker_AddComponentToActiveBOM loBom, compId, pn, rev, desc, uom, qtyPer, notes, rs
        End If
    Next key

    MsgBox "Selected components processed.", vbInformation, "Add Selected Components"
    Exit Sub

EH:
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub Picker_AddComponentToActiveBOM( _
    ByVal loBom As ListObject, _
    ByVal compId As String, _
    ByVal pn As String, _
    ByVal rev As String, _
    ByVal desc As String, _
    ByVal uom As String, _
    ByVal qtyPer As Double, _
    ByVal notes As String, _
    ByVal revStatus As String)
    Const PROC_NAME As String = "M_Data_BOMs_Picker.Picker_AddComponentToActiveBOM"

    On Error GoTo EH

    If qtyPer <= 0 Then
        MsgBox "QtyPer must be > 0.", vbExclamation, "Component Picker"
        Exit Sub
    End If

    If StrComp(revStatus, ACTIVE_LABEL, vbTextCompare) <> 0 Then
        MsgBox "Skipping non-active component: " & pn & " / " & rev & " (RevStatus=" & revStatus & ")", vbExclamation, "Component Picker"
        Exit Sub
    End If

    Bom_UpsertComponent loBom, compId, pn, rev, desc, uom, qtyPer, notes
    Exit Sub

EH:
    MsgBox "Add component failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Public Function GetActiveBomTable_Public() As ListObject
    Set GetActiveBomTable_Public = GetActiveBomTable()
End Function

'==========================
' BOM upsert (deterministic)
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
        arrPn = loBom.ListColumns(idxPn).DataBodyRange.value
        arrRev = loBom.ListColumns(idxRev).DataBodyRange.value

        For i = 1 To UBound(arrPn, 1)
            If StrComp(SafeText(arrPn(i, 1)), pn, vbTextCompare) = 0 And _
               StrComp(SafeText(arrRev(i, 1)), rev, vbTextCompare) = 0 Then

                Dim currentQty As Double
                currentQty = 0#
                If IsNumeric(loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).value) Then
                    currentQty = CDbl(loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).value)
                End If
                loBom.ListColumns(idxQty).DataBodyRange.Cells(i, 1).value = currentQty + qtyPer

                If ColumnExists(loBom, "UpdatedAt") Then loBom.ListColumns(GetColIndex(loBom, "UpdatedAt")).DataBodyRange.Cells(i, 1).value = Now
                If ColumnExists(loBom, "UpdatedBy") Then loBom.ListColumns(GetColIndex(loBom, "UpdatedBy")).DataBodyRange.Cells(i, 1).value = GetUserNameSafe()
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

'==========================
' Picker table writers
'==========================
Private Sub ClearPickResults(ByVal loPick As ListObject)
    If Not loPick.DataBodyRange Is Nothing Then
        loPick.DataBodyRange.Delete
    End If
End Sub

Private Sub WritePickResults(ByVal loPick As ListObject, ByRef outArr As Variant, ByVal outCount As Long)
    Dim ws As Worksheet
    Set ws = loPick.Parent

    Application.ScreenUpdating = False

    ' Clear existing rows
    If Not loPick.DataBodyRange Is Nothing Then
        loPick.DataBodyRange.Delete
    End If

    If outCount <= 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Add required rows
    Dim i As Long
    For i = 1 To outCount
        loPick.ListRows.Add
    Next i

    ' Write values in one shot
    loPick.DataBodyRange.value = Slice2D(outArr, outCount, 7)

    Application.ScreenUpdating = True
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

    arrId = loComps.ListColumns(idxId).DataBodyRange.value
    arrPn = loComps.ListColumns(idxPn).DataBodyRange.value
    arrRev = loComps.ListColumns(idxRev).DataBodyRange.value
    arrDesc = loComps.ListColumns(idxDesc).DataBodyRange.value
    arrUom = loComps.ListColumns(idxUom).DataBodyRange.value
    arrNotes = loComps.ListColumns(idxNotes).DataBodyRange.value
    arrRS = loComps.ListColumns(idxRS).DataBodyRange.value

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
    lr.Range.Cells(1, idx).value = v
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

Private Function GetUserNameSafe() As String
    Dim u As String
    u = Trim$(Environ$("Username"))
    If Len(u) = 0 Then u = Application.userName
    If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
    GetUserNameSafe = u
End Function
