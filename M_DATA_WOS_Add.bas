Attribute VB_Name = "M_DATA_WOS_Add"
Option Explicit

'===============================================================================
' Module: M_DATA_WOS_Add
'
' Purpose:
'   Controlled creation and controlled edits for WOS (build/work-order) rows.
'   This pattern intentionally separates:
'     1) user entry prompts
'     2) strict validation/guards
'     3) whitelisted write operations with audit stamping
'
'   The same pattern can be reused for other tables to reduce accidental edits,
'   data loss, and schema corruption.
'
' Inputs (Tabs/Tables/Headers):
'   - WOS!TBL_WOS (required): BuildID, AssemblyID, BuildQuantity, ShipTo
'   - Optional WOS columns: ShipTargetDate/DockDate, BuildName, BuildStatus,
'     BuildNotes, CreatedAt/By, UpdatedAt/By
'   - Optional Comps!TBL_COMPS for AssemblyID validation
'
' Outputs / Side effects:
'   - Adds one row in WOS!TBL_WOS for each successful create
'   - Updates an existing row only through a whitelisted updater
'   - Stamps audit fields where present
'
' Version: v0.1.0
' Author: ChatGPT
' Date: 2026-02-20
'===============================================================================

Private Const MODULE_VERSION As String = "0.1.0"
Private Const BUILD_ID_PREFIX As String = "NSWO-"
Private Const BUILD_STATUS_PLANNED As String = "PLANNED"

Public Sub UI_Add_WOS_Build()
    Const PROC_NAME As String = "M_DATA_WOS_Add.UI_Add_WOS_Build"

    Dim assemblyId As String
    Dim dueDateText As String
    Dim buildQtyText As String
    Dim destination As String
    Dim dueDate As Date
    Dim buildQty As Long
    Dim deliveryMethod As String
    Dim dockDateText As String
    Dim dockDate As Date
    Dim transitTimeText As String
    Dim transitTime As Long

    On Error GoTo EH

    FocusWosSheet

    If Not Gate_Ready(False) Then Exit Sub

    assemblyId = Trim$(InputBox("Top assembly part number to build (AssemblyID / TAID):", "New Build"))
    If Len(assemblyId) = 0 Then Exit Sub

    dueDateText = Trim$(InputBox("Due date:", "New Build", Format$(Date + 14, "yyyy-mm-dd")))
    If Len(dueDateText) = 0 Then Exit Sub
    If Not IsDate(dueDateText) Then
        MsgBox "Due date must be a valid date.", vbOKOnly, "New Build"
        Exit Sub
    End If
    dueDate = CDate(dueDateText)

    buildQtyText = Trim$(InputBox("Build quantity:", "New Build", "1"))
    If Len(buildQtyText) = 0 Then Exit Sub
    If Not IsNumeric(buildQtyText) Then
        MsgBox "Build quantity must be numeric.", vbOKOnly, "New Build"
        Exit Sub
    End If
    buildQty = CLng(buildQtyText)

    destination = Trim$(InputBox("Destination (ShipTo):", "New Build"))
    If Len(destination) = 0 Then
        MsgBox "Destination is required.", vbOKOnly, "New Build"
        Exit Sub
    End If

    deliveryMethod = PromptRequiredTextWithDefault("Delivery method:", "New Build", _
                        GetSchemaDefaultValue(SH_WOS, TBL_WOS, "DeliveryMethod"))
    If Len(deliveryMethod) = 0 Then Exit Sub

    dockDateText = PromptRequiredTextWithDefault("Dock date:", "New Build", _
                        GetSchemaDefaultValue(SH_WOS, TBL_WOS, "DockDate"))
    If Len(dockDateText) = 0 Then Exit Sub
    If Not IsDate(dockDateText) Then
        MsgBox "Dock date must be a valid date.", vbOKOnly, "New Build"
        Exit Sub
    End If
    dockDate = CDate(dockDateText)

    transitTimeText = PromptRequiredTextWithDefault("Transit time (days):", "New Build", _
                          GetSchemaDefaultValue(SH_WOS, TBL_WOS, "TransitTime"))
    If Len(transitTimeText) = 0 Then Exit Sub
    If Not IsNumeric(transitTimeText) Then
        MsgBox "Transit time must be numeric.", vbOKOnly, "New Build"
        Exit Sub
    End If
    transitTime = CLng(transitTimeText)
    If transitTime < 0 Then
        MsgBox "Transit time cannot be negative.", vbOKOnly, "New Build"
        Exit Sub
    End If

    Add_WOS_Build_FromInputs assemblyId, dueDate, buildQty, destination, "", "", deliveryMethod, dockDate, transitTime
    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "UI add build failed", Err.Description, Err.Number
    GoToLogSheet
    M_Core_UX.ShowFailureMessageWithLogFocus PROC_NAME, "New Build", "New build failed.", Err.Description, Err.Number
End Sub

Public Sub Add_WOS_Build_FromInputs(ByVal assemblyId As String, ByVal dueDate As Date, ByVal buildQty As Long, ByVal destination As String, _
                                    Optional ByVal buildName As String = "", Optional ByVal buildNotes As String = "", _
                                    Optional ByVal deliveryMethod As String = "", Optional ByVal dockDate As Variant, Optional ByVal transitTime As Variant)
    Const PROC_NAME As String = "M_DATA_WOS_Add.Add_WOS_Build_FromInputs"

    Dim wb As Workbook
    Dim wsWos As Worksheet
    Dim loWos As ListObject
    Dim lr As ListRow
    Dim dueDateCol As String

    Dim buildId As String
    Dim nowStamp As Date
    Dim actor As String

    On Error GoTo EH

    FocusWosSheet

    If Not Gate_Ready(False) Then Exit Sub

    assemblyId = Trim$(assemblyId)
    destination = Trim$(destination)
    buildName = Trim$(buildName)
    buildNotes = Trim$(buildNotes)
    deliveryMethod = Trim$(deliveryMethod)

    If Len(assemblyId) = 0 Then Err.Raise vbObjectError + 7001, PROC_NAME, "AssemblyID is required."
    If buildQty <= 0 Then Err.Raise vbObjectError + 7002, PROC_NAME, "BuildQuantity must be > 0."
    If Len(destination) = 0 Then Err.Raise vbObjectError + 7003, PROC_NAME, "ShipTo is required."

    Set wb = ThisWorkbook
    Set wsWos = wb.Worksheets(SH_WOS)
    Set loWos = wsWos.ListObjects(TBL_WOS)

    RequireColumn loWos, "BuildID"
    RequireColumn loWos, "AssemblyID"
    RequireColumn loWos, "BuildQuantity"
    RequireColumn loWos, "ShipTo"

    dueDateCol = ResolveDueDateColumn(loWos)

    If Not AssemblyExistsInWorkbook(wb, assemblyId) Then
        Err.Raise vbObjectError + 7004, PROC_NAME, "AssemblyID '" & assemblyId & "' was not found in BOMS.TAID or Comps.CompID."
    End If

    buildId = GenerateNextBuildId(loWos)
    If Len(buildId) = 0 Then Err.Raise vbObjectError + 7005, PROC_NAME, "Could not generate BuildID."

    If Len(buildName) = 0 Then
        buildName = assemblyId & "_" & Format$(dueDate, "yyyymmdd")
    End If

    Set lr = loWos.ListRows.Add

    SetByHeader loWos, lr, "BuildID", buildId
    SetByHeader loWos, lr, "AssemblyID", assemblyId
    SetByHeader loWos, lr, "BuildQuantity", buildQty
    SetByHeader loWos, lr, "ShipTo", destination

    If ColumnExists(loWos, "DeliveryMethod") Then SetByHeader loWos, lr, "DeliveryMethod", deliveryMethod
    If ColumnExists(loWos, "DockDate") And Not IsMissing(dockDate) Then SetByHeader loWos, lr, "DockDate", CDate(dockDate)
    If ColumnExists(loWos, "TransitTime") And Not IsMissing(transitTime) Then SetByHeader loWos, lr, "TransitTime", CLng(transitTime)

    If Len(dueDateCol) > 0 Then SetByHeader loWos, lr, dueDateCol, dueDate
    If ColumnExists(loWos, "BuildName") Then SetByHeader loWos, lr, "BuildName", buildName
    If ColumnExists(loWos, "BuildStatus") Then SetByHeader loWos, lr, "BuildStatus", BUILD_STATUS_PLANNED
    If ColumnExists(loWos, "BuildNotes") Then SetByHeader loWos, lr, "BuildNotes", buildNotes

    nowStamp = Now
    actor = SafeActorName()
    StampAuditIfPresent loWos, lr, actor, nowStamp

    If Not M_Core_DataIntegrity.RunDataCheck(False) Then
        If Not lr Is Nothing Then lr.Delete
        Err.Raise vbObjectError + 7006, PROC_NAME, "Schema/data integrity requirements failed after build creation."
    End If

    M_Core_Logging.LogInfo PROC_NAME, "Created WOS build", _
        "BuildID=" & buildId & "; AssemblyID=" & assemblyId & "; Qty=" & CStr(buildQty) & _
        "; DeliveryMethod=" & deliveryMethod & "; DockDate=" & IIf(IsMissing(dockDate), "", CStr(dockDate)) & _
        "; TransitTime=" & IIf(IsMissing(transitTime), "", CStr(transitTime)) & "; DueCol=" & dueDateCol & "; Version=" & MODULE_VERSION

    If M_Core_UX.ShouldShowSuccessMessage("Add_WOS_Build_FromInputs") Then
        MsgBox "Build created successfully." & vbCrLf & _
               "BuildID: " & buildId, vbOKOnly, "New Build"
    End If
    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "Create WOS build failed", Err.Description, Err.Number
    GoToLogSheet
    M_Core_UX.ShowFailureMessageWithLogFocus PROC_NAME, "New Build", "Create build failed.", Err.Description, Err.Number
End Sub

Public Function Update_WOS_Build_Controlled(ByVal buildId As String, _
                                            Optional ByVal dueDate As Variant, _
                                            Optional ByVal buildQty As Variant, _
                                            Optional ByVal destination As Variant, _
                                            Optional ByVal buildStatus As Variant, _
                                            Optional ByVal buildNotes As Variant, _
                                            Optional ByVal allowClosedBuildEdit As Boolean = False) As Boolean
    Const PROC_NAME As String = "M_DATA_WOS_Add.Update_WOS_Build_Controlled"

    Dim wb As Workbook
    Dim wsWos As Worksheet
    Dim loWos As ListObject
    Dim rowIx As Long
    Dim dueDateCol As String
    Dim currentStatus As String

    On Error GoTo EH

    buildId = Trim$(buildId)
    If Len(buildId) = 0 Then Err.Raise vbObjectError + 7101, PROC_NAME, "BuildID is required."

    Set wb = ThisWorkbook
    Set wsWos = wb.Worksheets(SH_WOS)
    Set loWos = wsWos.ListObjects(TBL_WOS)

    RequireColumn loWos, "BuildID"

    rowIx = FindRowByColumnValue(loWos, "BuildID", buildId)
    If rowIx = 0 Then Err.Raise vbObjectError + 7102, PROC_NAME, "BuildID not found: " & buildId

    If ColumnExists(loWos, "BuildStatus") Then
        currentStatus = UCase$(Trim$(SafeCellText(loWos.ListColumns(GetColIndex(loWos, "BuildStatus")).DataBodyRange.Cells(rowIx, 1).Value)))
        If Not allowClosedBuildEdit Then
            If currentStatus = "SHIPPED" Or currentStatus = "CLOSED" Or currentStatus = "COMPLETE" Then
                Err.Raise vbObjectError + 7103, PROC_NAME, "Build is closed and cannot be edited without override."
            End If
        End If
    End If

    dueDateCol = ResolveDueDateColumn(loWos)

    If Not IsMissing(dueDate) Then
        If Len(dueDateCol) = 0 Then Err.Raise vbObjectError + 7104, PROC_NAME, "No due-date column available (ShipTargetDate/DockDate)."
        If Not IsDate(dueDate) Then Err.Raise vbObjectError + 7105, PROC_NAME, "Due date must be a valid date."
        SetCellByHeader loWos, rowIx, dueDateCol, CDate(dueDate)
    End If

    If Not IsMissing(buildQty) Then
        If Not IsNumeric(buildQty) Then Err.Raise vbObjectError + 7106, PROC_NAME, "BuildQuantity must be numeric."
        If CLng(buildQty) <= 0 Then Err.Raise vbObjectError + 7107, PROC_NAME, "BuildQuantity must be > 0."
        SetCellByHeader loWos, rowIx, "BuildQuantity", CLng(buildQty)
    End If

    If Not IsMissing(destination) Then
        If Len(Trim$(CStr(destination))) = 0 Then Err.Raise vbObjectError + 7108, PROC_NAME, "ShipTo cannot be blank."
        SetCellByHeader loWos, rowIx, "ShipTo", Trim$(CStr(destination))
    End If

    If Not IsMissing(buildStatus) Then
        If Not ColumnExists(loWos, "BuildStatus") Then Err.Raise vbObjectError + 7109, PROC_NAME, "BuildStatus column missing."
        If Len(Trim$(CStr(buildStatus))) = 0 Then Err.Raise vbObjectError + 7110, PROC_NAME, "BuildStatus cannot be blank."
        SetCellByHeader loWos, rowIx, "BuildStatus", UCase$(Trim$(CStr(buildStatus)))
    End If

    If Not IsMissing(buildNotes) Then
        If ColumnExists(loWos, "BuildNotes") Then SetCellByHeader loWos, rowIx, "BuildNotes", CStr(buildNotes)
    End If

    StampAuditOnExistingRow loWos, rowIx, SafeActorName(), Now

    M_Core_Logging.LogInfo PROC_NAME, "Updated WOS build", "BuildID=" & buildId & "; Version=" & MODULE_VERSION
    Update_WOS_Build_Controlled = True
    Exit Function

EH:
    M_Core_Logging.LogError PROC_NAME, "Controlled update failed", Err.Description, Err.Number
    MsgBox "Build update blocked." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Edit Build"
    Update_WOS_Build_Controlled = False
End Function

Private Function ResolveDueDateColumn(ByVal lo As ListObject) As String
    If ColumnExists(lo, "ShipTargetDate") Then
        ResolveDueDateColumn = "ShipTargetDate"
    ElseIf ColumnExists(lo, "DockDate") Then
        ResolveDueDateColumn = "DockDate"
    Else
        ResolveDueDateColumn = vbNullString
    End If
End Function

Private Function GenerateNextBuildId(ByVal loWos As ListObject) As String
    Dim thisYear As String
    Dim maxSeq As Long
    Dim i As Long

    thisYear = Right$(Format$(Date, "yyyy"), 2)
    maxSeq = 0

    If loWos.DataBodyRange Is Nothing Then
        GenerateNextBuildId = BUILD_ID_PREFIX & thisYear & "-001"
        Exit Function
    End If

    For i = 1 To loWos.ListRows.Count
        Dim existingId As String
        existingId = SafeCellText(loWos.ListColumns(GetColIndex(loWos, "BuildID")).DataBodyRange.Cells(i, 1).Value)
        maxSeq = Application.WorksheetFunction.Max(maxSeq, ExtractSeqForYear(existingId, thisYear))
    Next i

    GenerateNextBuildId = BUILD_ID_PREFIX & thisYear & "-" & Format$(maxSeq + 1, "000")
End Function

Private Function ExtractSeqForYear(ByVal buildId As String, ByVal yy As String) As Long
    Dim pfx As String
    pfx = UCase$(BUILD_ID_PREFIX & yy & "-")

    buildId = UCase$(Trim$(buildId))
    If Left$(buildId, Len(pfx)) <> pfx Then Exit Function

    Dim seqText As String
    seqText = Mid$(buildId, Len(pfx) + 1)
    If Not IsNumeric(seqText) Then Exit Function

    ExtractSeqForYear = CLng(seqText)
End Function

Private Function AssemblyExistsInWorkbook(ByVal wb As Workbook, ByVal assemblyId As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    assemblyId = Trim$(assemblyId)
    If Len(assemblyId) = 0 Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets(SH_BOMS)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set lo = M_Core_Utils.SafeGetListObject(ws, TBL_BOMS)
        If Not lo Is Nothing Then
            If ColumnExists(lo, "TAID") Then
                If FindRowByColumnValue(lo, "TAID", assemblyId) > 0 Then
                    AssemblyExistsInWorkbook = True
                    Exit Function
                End If
            End If
        End If
    End If

    Set ws = Nothing
    Set lo = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(SH_COMPS)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set lo = M_Core_Utils.SafeGetListObject(ws, TBL_COMPS)
        If Not lo Is Nothing Then
            If ColumnExists(lo, "CompID") Then
                If FindRowByColumnValue(lo, "CompID", assemblyId) > 0 Then
                    AssemblyExistsInWorkbook = True
                    Exit Function
                End If
            End If
        End If
    End If
End Function

Private Function PromptRequiredTextWithDefault(ByVal promptLabel As String, ByVal title As String, ByVal defaultValue As String) As String
    Dim resp As String
    Dim displayDefault As String

    displayDefault = Trim$(defaultValue)
    If Len(displayDefault) = 0 Then
        If InStr(1, UCase$(promptLabel), "DATE", vbTextCompare) > 0 Then
            displayDefault = Format$(Date + 14, "yyyy-mm-dd")
        ElseIf InStr(1, UCase$(promptLabel), "TRANSIT", vbTextCompare) > 0 Then
            displayDefault = "0"
        End If
    End If

    resp = InputBox(promptLabel & " (required):", title, displayDefault)
    resp = Trim$(resp)
    PromptRequiredTextWithDefault = resp
End Function

Private Function GetSchemaDefaultValue(ByVal tabName As String, ByVal tableName As String, ByVal columnHeader As String) As String
    Const SCHEMA_TABLE As String = "TBL_SCHEMA"
    Const H_TAB As String = "TAB_NAME"
    Const H_TBL As String = "TABLE_NAME"
    Const H_COL As String = "COLUMN_HEADER"
    Const H_DEF As String = "DefaultValue"

    Dim lo As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim idxTab As Long, idxTbl As Long, idxCol As Long, idxDef As Long
    Dim arr As Variant
    Dim r As Long

    GetSchemaDefaultValue = vbNullString
    Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(SCHEMA_TABLE)
        On Error GoTo 0
        If Not lo Is Nothing Then Exit For
    Next ws

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxTab = GetColIndex(lo, H_TAB)
    idxTbl = GetColIndex(lo, H_TBL)
    idxCol = GetColIndex(lo, H_COL)
    idxDef = GetColIndex(lo, H_DEF)
    If idxTab = 0 Or idxTbl = 0 Or idxCol = 0 Or idxDef = 0 Then Exit Function

    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(r, idxTab))), tabName, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(arr(r, idxTbl))), tableName, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(arr(r, idxCol))), columnHeader, vbTextCompare) = 0 Then
            GetSchemaDefaultValue = Trim$(CStr(arr(r, idxDef)))
            Exit Function
        End If
    Next r
End Function

Private Sub StampAuditIfPresent(ByVal lo As ListObject, ByVal lr As ListRow, ByVal actor As String, ByVal ts As Date)
    If ColumnExists(lo, "CreatedAt") Then SetByHeader lo, lr, "CreatedAt", ts
    If ColumnExists(lo, "CreatedBy") Then SetByHeader lo, lr, "CreatedBy", actor
    If ColumnExists(lo, "UpdatedAt") Then SetByHeader lo, lr, "UpdatedAt", ts
    If ColumnExists(lo, "UpdatedBy") Then SetByHeader lo, lr, "UpdatedBy", actor
End Sub

Private Sub StampAuditOnExistingRow(ByVal lo As ListObject, ByVal rowIx As Long, ByVal actor As String, ByVal ts As Date)
    If ColumnExists(lo, "UpdatedAt") Then SetCellByHeader lo, rowIx, "UpdatedAt", ts
    If ColumnExists(lo, "UpdatedBy") Then SetCellByHeader lo, rowIx, "UpdatedBy", actor
End Sub

Private Function SafeActorName() As String
    SafeActorName = Trim$(Environ$("Username"))
    If Len(SafeActorName) = 0 Then SafeActorName = "UNKNOWN"
End Function

Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If Not ColumnExists(lo, header) Then
        Err.Raise vbObjectError + 7190, "M_DATA_WOS_Add.RequireColumn", _
            "Missing required column '" & header & "' in table '" & lo.Name & "'."
    End If
End Sub

Private Function ColumnExists(ByVal lo As ListObject, ByVal header As String) As Boolean
    ColumnExists = (GetColIndex(lo, header) > 0)
End Function

Private Function GetColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(header), vbTextCompare) = 0 Then
            GetColIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx <= 0 Then Exit Sub
    lr.Range.Cells(1, idx).Value = v
End Sub

Private Sub SetCellByHeader(ByVal lo As ListObject, ByVal rowIx As Long, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx <= 0 Then Err.Raise vbObjectError + 7191, "M_DATA_WOS_Add.SetCellByHeader", "Missing column: " & header
    lo.ListColumns(idx).DataBodyRange.Cells(rowIx, 1).Value = v
End Sub

Private Function FindRowByColumnValue(ByVal lo As ListObject, ByVal header As String, ByVal valueText As String) As Long
    Dim idx As Long
    Dim i As Long

    idx = GetColIndex(lo, header)
    If idx <= 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    valueText = Trim$(valueText)

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(SafeCellText(lo.ListColumns(idx).DataBodyRange.Cells(i, 1).Value)), valueText, vbTextCompare) = 0 Then
            FindRowByColumnValue = i
            Exit Function
        End If
    Next i
End Function

Private Function SafeCellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeCellText = Trim$(CStr(v))
End Function

Private Sub FocusWosSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_WOS).Activate
    On Error GoTo 0
End Sub

Private Sub GoToLogSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets("Log").Activate
    On Error GoTo 0
End Sub
