Attribute VB_Name = "M_Data_Calc_Demand"
Option Explicit

'===============================================================================
' Module: M_Data_Calc_Demand
'
' Purpose:
'   Recalculate Demand.TBL_DEMAND from open build quantities in WOS and BOM lines.
'   For each build row, component demand = BOM.QtyPer * WOS.BuildQuantity.
'   Demand is aggregated by component identity (CompID + OurPN + OurRev) across
'   all build rows (no per-build split retained in the final demand table).
'
' Inputs (Tabs/Tables/Headers):
'   - BOMS!TBL_BOMS: TAID, BOMTab
'   - WOS!TBL_WOS: AssemblyID, BuildQuantity
'   - Each BOM sheet listed in BOMS.BOMTab, first table on the sheet:
'       CompID, OurPN, OurRev, Description (or ComponentDescription), UOM, QtyPer
'   - Demand!TBL_DEMAND (target): expected quantity column TotalDemand (fallbacks supported)
'
' Outputs / Side effects:
'   - Clears and repopulates Demand.TBL_DEMAND with aggregated demand lines.
'   - Writes available identity fields + quantity fields where matching headers exist.
'
' Version: v0.1.0
' Author: ChatGPT
' Date: 2026-02-20
'===============================================================================

Private Const MODULE_VERSION As String = "0.1.0"

Public Sub UI_Recalc_Demand_From_WOS_BOM()
    Const PROC_NAME As String = "M_Data_Calc_Demand.UI_Recalc_Demand_From_WOS_BOM"

    On Error GoTo EH

    If Not M_Core_Gate.Gate_Ready(True) Then Exit Sub
    Recalc_Demand_From_WOS_BOM M_Core_UX.ShouldShowSuccessMessage("UI_Recalc_Demand_From_WOS_BOM")
    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "UI demand recalc failed", Err.Description, Err.Number
    MsgBox "Demand recalculation failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Demand"
End Sub

Public Sub Recalc_Demand_From_WOS_BOM(Optional ByVal showUserMessage As Boolean = True)
    Const PROC_NAME As String = "M_Data_Calc_Demand.Recalc_Demand_From_WOS_BOM"

    Dim wb As Workbook
    Dim wsBoms As Worksheet, wsWos As Worksheet, wsDemand As Worksheet
    Dim loBoms As ListObject, loWos As ListObject, loDemand As ListObject

    Dim dictBomTabByTaid As Object
    Dim dictDemand As Object

    On Error GoTo EH

    Set wb = ThisWorkbook
    Set wsBoms = wb.Worksheets(SH_BOMS)
    Set wsWos = wb.Worksheets(SH_WOS)
    Set wsDemand = wb.Worksheets(SH_DEMAND)

    Set loBoms = wsBoms.ListObjects(TBL_BOMS)
    Set loWos = wsWos.ListObjects(TBL_WOS)
    Set loDemand = wsDemand.ListObjects(TBL_DEMAND)

    RequireColumn loBoms, "TAID"
    RequireColumn loBoms, "BOMTab"
    RequireColumn loWos, "AssemblyID"
    RequireColumn loWos, "BuildQuantity"

    Set dictBomTabByTaid = CreateObject("Scripting.Dictionary")
    dictBomTabByTaid.CompareMode = vbTextCompare

    Set dictDemand = CreateObject("Scripting.Dictionary")
    dictDemand.CompareMode = vbTextCompare

    BuildBomLookup loBoms, dictBomTabByTaid
    AccumulateDemand loWos, dictBomTabByTaid, dictDemand

    WriteDemandRows loDemand, dictDemand

    M_Core_Logging.LogInfo PROC_NAME, "Demand recalculated", _
        "Rows=" & CStr(dictDemand.Count) & "; Version=" & MODULE_VERSION

    If showUserMessage Then
        MsgBox "Demand recalculated successfully." & vbCrLf & _
               "Demand rows: " & CStr(dictDemand.Count), vbInformation, "Demand"
    End If

    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "Demand recalculation failed", Err.Description, Err.Number
    If showUserMessage Then
        MsgBox "Demand recalculation failed." & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Demand"
    End If
End Sub

Private Sub BuildBomLookup(ByVal loBoms As ListObject, ByVal dictBomTabByTaid As Object)
    Dim idxTaid As Long, idxBomTab As Long
    Dim arrTaid As Variant, arrBomTab As Variant
    Dim i As Long
    Dim taid As String, bomTab As String

    idxTaid = GetColIndex(loBoms, "TAID")
    idxBomTab = GetColIndex(loBoms, "BOMTab")

    If loBoms.DataBodyRange Is Nothing Then Exit Sub

    arrTaid = loBoms.ListColumns(idxTaid).DataBodyRange.Value
    arrBomTab = loBoms.ListColumns(idxBomTab).DataBodyRange.Value

    For i = 1 To UBound(arrTaid, 1)
        taid = Trim$(CStr(arrTaid(i, 1)))
        bomTab = Trim$(CStr(arrBomTab(i, 1)))
        If Len(taid) > 0 And Len(bomTab) > 0 Then
            dictBomTabByTaid(taid) = bomTab
        End If
    Next i
End Sub

Private Sub AccumulateDemand(ByVal loWos As ListObject, ByVal dictBomTabByTaid As Object, ByVal dictDemand As Object)
    Dim idxAssembly As Long, idxBuildQty As Long
    Dim arrAssembly As Variant, arrBuildQty As Variant
    Dim i As Long

    Dim assemblyId As String
    Dim buildQty As Double
    Dim bomTab As String

    idxAssembly = GetColIndex(loWos, "AssemblyID")
    idxBuildQty = GetColIndex(loWos, "BuildQuantity")

    If loWos.DataBodyRange Is Nothing Then Exit Sub

    arrAssembly = loWos.ListColumns(idxAssembly).DataBodyRange.Value
    arrBuildQty = loWos.ListColumns(idxBuildQty).DataBodyRange.Value

    For i = 1 To UBound(arrAssembly, 1)
        assemblyId = Trim$(CStr(arrAssembly(i, 1)))
        buildQty = ParsePositiveDouble(arrBuildQty(i, 1))

        If Len(assemblyId) = 0 Or buildQty <= 0 Then GoTo ContinueLoop
        If Not dictBomTabByTaid.Exists(assemblyId) Then GoTo ContinueLoop

        bomTab = CStr(dictBomTabByTaid(assemblyId))
        AccumulateBomLines bomTab, buildQty, dictDemand

ContinueLoop:
    Next i
End Sub

Private Sub AccumulateBomLines(ByVal bomTab As String, ByVal buildQty As Double, ByVal dictDemand As Object)
    Dim wsBom As Worksheet
    Dim loBom As ListObject

    Dim idxCompId As Long, idxPn As Long, idxRev As Long, idxDesc As Long, idxUom As Long, idxQtyPer As Long
    Dim arrCompId As Variant, arrPn As Variant, arrRev As Variant, arrDesc As Variant, arrUom As Variant, arrQtyPer As Variant
    Dim i As Long

    Dim compId As String, ourPn As String, ourRev As String, descr As String, uom As String
    Dim qtyPer As Double, lineDemand As Double
    Dim key As String
    Dim rec As Variant

    Set wsBom = GetWorksheetSafe(ThisWorkbook, bomTab)
    If wsBom Is Nothing Then Exit Sub
    If wsBom.ListObjects.Count = 0 Then Exit Sub

    Set loBom = wsBom.ListObjects(1)
    If loBom.DataBodyRange Is Nothing Then Exit Sub

    idxCompId = GetColIndex(loBom, "CompID")
    idxPn = GetColIndex(loBom, "OurPN")
    idxRev = GetColIndex(loBom, "OurRev")
    idxDesc = GetColIndex(loBom, "Description")
    If idxDesc = 0 Then idxDesc = GetColIndex(loBom, "ComponentDescription")
    idxUom = GetColIndex(loBom, "UOM")
    idxQtyPer = GetColIndex(loBom, "QtyPer")

    If idxQtyPer = 0 Then Exit Sub

    If idxCompId > 0 Then arrCompId = loBom.ListColumns(idxCompId).DataBodyRange.Value
    If idxPn > 0 Then arrPn = loBom.ListColumns(idxPn).DataBodyRange.Value
    If idxRev > 0 Then arrRev = loBom.ListColumns(idxRev).DataBodyRange.Value
    If idxDesc > 0 Then arrDesc = loBom.ListColumns(idxDesc).DataBodyRange.Value
    If idxUom > 0 Then arrUom = loBom.ListColumns(idxUom).DataBodyRange.Value
    arrQtyPer = loBom.ListColumns(idxQtyPer).DataBodyRange.Value

    For i = 1 To UBound(arrQtyPer, 1)
        compId = ReadArrText(arrCompId, i)
        ourPn = ReadArrText(arrPn, i)
        ourRev = ReadArrText(arrRev, i)
        descr = ReadArrText(arrDesc, i)
        uom = ReadArrText(arrUom, i)

        If Len(compId) = 0 And Len(ourPn) = 0 And Len(ourRev) = 0 Then GoTo ContinueLoop

        qtyPer = ParsePositiveDouble(arrQtyPer(i, 1))
        If qtyPer <= 0 Then GoTo ContinueLoop

        lineDemand = qtyPer * buildQty
        key = UCase$(compId) & "|" & UCase$(ourPn) & "|" & UCase$(ourRev)

        If dictDemand.Exists(key) Then
            rec = dictDemand(key)
            rec(6) = CDbl(rec(6)) + lineDemand
        Else
            ReDim rec(1 To 6)
            rec(1) = compId
            rec(2) = ourPn
            rec(3) = ourRev
            rec(4) = descr
            rec(5) = uom
            rec(6) = lineDemand
        End If

        dictDemand(key) = rec

ContinueLoop:
    Next i
End Sub

Private Sub WriteDemandRows(ByVal loDemand As ListObject, ByVal dictDemand As Object)
    Dim qtyCol As String
    Dim k As Variant
    Dim rec As Variant
    Dim lr As ListRow

    qtyCol = ResolveDemandQtyColumn(loDemand)

    ClearTableRows loDemand

    For Each k In dictDemand.Keys
        rec = dictDemand(k)
        Set lr = loDemand.ListRows.Add

        SetIfExists loDemand, lr, "CompID", rec(1)
        SetIfExists loDemand, lr, "OurPN", rec(2)
        SetIfExists loDemand, lr, "OurRev", rec(3)

        If ColumnExists(loDemand, "Description") Then
            SetIfExists loDemand, lr, "Description", rec(4)
        ElseIf ColumnExists(loDemand, "ComponentDescription") Then
            SetIfExists loDemand, lr, "ComponentDescription", rec(4)
        End If

        SetIfExists loDemand, lr, "UOM", rec(5)

        If Len(qtyCol) > 0 Then SetIfExists loDemand, lr, qtyCol, CDbl(rec(6))
        If ColumnExists(loDemand, "NetDemand") Then SetIfExists loDemand, lr, "NetDemand", CDbl(rec(6))

        If ColumnExists(loDemand, "UpdatedAt") Then SetIfExists loDemand, lr, "UpdatedAt", Now
        If ColumnExists(loDemand, "UpdatedBy") Then SetIfExists loDemand, lr, "UpdatedBy", SafeActorName()
    Next k
End Sub

Private Function ResolveDemandQtyColumn(ByVal loDemand As ListObject) As String
    If ColumnExists(loDemand, "TotalDemand") Then
        ResolveDemandQtyColumn = "TotalDemand"
    ElseIf ColumnExists(loDemand, "BuildQuantityDemand") Then
        ResolveDemandQtyColumn = "BuildQuantityDemand"
    ElseIf ColumnExists(loDemand, "DemandQuantity") Then
        ResolveDemandQtyColumn = "DemandQuantity"
    ElseIf ColumnExists(loDemand, "Quantity") Then
        ResolveDemandQtyColumn = "Quantity"
    Else
        ResolveDemandQtyColumn = vbNullString
    End If
End Function

Private Function ParsePositiveDouble(ByVal v As Variant) As Double
    If IsError(v) Or IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    ParsePositiveDouble = CDbl(v)
    If ParsePositiveDouble < 0 Then ParsePositiveDouble = 0
End Function

Private Function GetWorksheetSafe(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function ReadArrText(ByVal arr As Variant, ByVal idx As Long) As String
    On Error GoTo EH
    If IsEmpty(arr) Then Exit Function
    ReadArrText = Trim$(CStr(arr(idx, 1)))
    Exit Function
EH:
    ReadArrText = vbNullString
End Function

Private Sub ClearTableRows(ByVal lo As ListObject)
    Dim i As Long
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = lo.ListRows.Count To 1 Step -1
        lo.ListRows(i).Delete
    Next i
End Sub

Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 8600, "M_Data_Calc_Demand.RequireColumn", _
                  "Missing column '" & header & "' in table '" & lo.Name & "'."
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

Private Sub SetIfExists(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal value As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Exit Sub
    lr.Range.Cells(1, idx).Value = value
End Sub

Private Function SafeActorName() As String
    SafeActorName = Trim$(Environ$("Username"))
    If Len(SafeActorName) = 0 Then SafeActorName = Application.UserName
    If Len(SafeActorName) = 0 Then SafeActorName = "UNKNOWN"
End Function
