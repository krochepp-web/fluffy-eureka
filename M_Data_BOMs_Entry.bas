Attribute VB_Name = "M_Data_BOMs_Entry"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_Entry
'
' Purpose:
'   Create a new BOM sheet from BOM_TEMPLATE for a top assembly (TA) WITHOUT
'   requiring Comps.IsBuildable. Collects TAID (unique), TAPN, TARev, TADesc,
'   and optional BOM_NOTES from optional arguments or manual prompts, then
'   registers the result in BOMS.
'   Creates sheet "BOM_<TAID>" (exact TAID suffix) and registers row
'   in BOMS.TBL_BOMS. Includes line-numbered diagnostics for debugging Error 13.
'
' Inputs (Tabs/Tables/Headers):
'   - BOM_TEMPLATE: TBL_BOM_TEMPLATE
'   - BOMS: TBL_BOMS [BOMID, BOMTab, TAID, BOM_NOTES] (+ optional TAPN, TARev, TADesc)
'   - Comps (optional): TBL_COMPS for best-effort TA validation
'
' Outputs:
'   - New BOM sheet copied from template
'   - New row in BOMS.TBL_BOMS
'
' Errors & Guards:
'   - SafeText prevents Type mismatch on Excel error values in cells
'   - Debug handler reports Erl line number and key context
'
' Version: v0.4.0
' Author: ChatGPT
' Date: 2026-02-07
'===============================================================================

Public Sub UI_Create_BOM_For_Assembly( _
    Optional ByVal taIdIn As String = "", _
    Optional ByVal taPnIn As String = "", _
    Optional ByVal taRevIn As String = "", _
    Optional ByVal taDescIn As String = "", _
    Optional ByVal bomNotesIn As String = "")
    Const PROC_NAME As String = "M_Data_BOMs_Entry.UI_Create_BOM_For_Assembly"

    Dim taId As String
    Dim taPn As String
    Dim taRev As String
    Dim taDesc As String
    Dim bomNotes As String

    On Error GoTo EH

    If Not GateReady_Safe(True) Then Exit Sub

    taId = Trim$(taIdIn)
    taPn = Trim$(taPnIn)
    taRev = Trim$(taRevIn)
    taDesc = Trim$(taDescIn)
    bomNotes = Trim$(bomNotesIn)

    If Len(taId) = 0 Then taId = PromptText("Enter TAID (unique):", "New BOM")
    If Len(taPn) = 0 Then taPn = PromptText("Enter top assembly part number (TAPN):", "New BOM")
    If Len(taRev) = 0 Then taRev = PromptText("Enter top assembly revision (TARev):", "New BOM")
    If Len(taDesc) = 0 Then taDesc = PromptText("Enter top assembly description (TADesc):", "New BOM")
    If Len(bomNotes) = 0 Then bomNotes = PromptText("Enter BOM notes (optional):", "New BOM")

    If Len(taId) = 0 Or Len(taPn) = 0 Or Len(taRev) = 0 Or Len(taDesc) = 0 Then
        MsgBox "BOM creation cancelled. All fields are required.", vbInformation, "New BOM"
        Exit Sub
    End If

    Create_BOM_For_Assembly_FromInputs taId, taPn, taRev, taDesc, bomNotes
    Exit Sub

EH:
    MsgBox "Manual BOM creation failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, PROC_NAME
End Sub

Public Sub Create_BOM_For_Assembly_FromInputs( _
    ByVal taId As String, _
    ByVal taPn As String, _
    ByVal taRev As String, _
    ByVal taDesc As String, _
    Optional ByVal bomNotes As String = "")
    Const PROC_NAME As String = "M_Data_BOMs_Entry.Create_BOM_For_Assembly_FromInputs"

    Const SH_TEMPLATE As String = "BOM_TEMPLATE"
    Const LO_TEMPLATE As String = "TBL_BOM_TEMPLATE"

    Const SH_BOMS As String = "BOMS"
    Const LO_BOMS As String = "TBL_BOMS"

    Const SH_COMPS As String = "Comps"
    Const LO_COMPS As String = "TBL_COMPS"

    Const BOM_TAB_PREFIX As String = "BOM_"
    Const BOM_ID_PREFIX As String = "BOM-"
    Const BOM_ID_PAD As Long = 4

    Const ACTIVE_REVSTATUS As String = "Active"

    Dim wb As Workbook
    Dim wsTemplate As Worksheet
    Dim wsBoms As Worksheet
    Dim wsComps As Worksheet
    Dim wsNew As Worksheet

    Dim loTemplate As ListObject
    Dim loBoms As ListObject
    Dim loComps As ListObject
    Dim loNew As ListObject

    Dim bomId As String
    Dim newSheetName As String
    Dim newTableName As String
    Dim createdAt As Date
    Dim createdBy As String

    On Error GoTo EH

If Not GateReady_Safe(True) Then GoTo CleanExit

Set wb = ThisWorkbook
Set wsTemplate = wb.Worksheets(SH_TEMPLATE)
Set wsBoms = wb.Worksheets(SH_BOMS)

Set loTemplate = wsTemplate.ListObjects(LO_TEMPLATE)
Set loBoms = wsBoms.ListObjects(LO_BOMS)

    ' Required headers
RequireColumn loTemplate, "CompID"
RequireColumn loTemplate, "OurPN"
RequireColumn loTemplate, "OurRev"
RequireColumn loTemplate, "Description"
RequireColumn loTemplate, "UOM"
RequireColumn loTemplate, "QtyPer"
RequireColumn loTemplate, "CompNotes"
RequireColumn loTemplate, "CreatedAt"
RequireColumn loTemplate, "CreatedBy"
RequireColumn loTemplate, "UpdatedAt"
RequireColumn loTemplate, "UpdatedBy"

RequireColumn loBoms, "BOMID"
RequireColumn loBoms, "BOMTab"
RequireColumn loBoms, "TAID"
RequireColumn loBoms, "BOM_NOTES"
' TARev and TADesc are optional in legacy workbooks

    ' Inputs
taId = Trim$(taId)
taPn = Trim$(taPn)
taRev = Trim$(taRev)
taDesc = Trim$(taDesc)
bomNotes = Trim$(bomNotes)

If Len(taId) = 0 Or Len(taPn) = 0 Or Len(taRev) = 0 Or Len(taDesc) = 0 Then
MsgBox "All fields are required (TAID, TAPN, TARev, TADesc).", vbExclamation, "New BOM"
GoTo CleanExit
End If

If TaId_Exists(loBoms, taId) Then
MsgBox "TAID '" & taId & "' already exists in BOMS (TAID). Choose a unique TAID.", vbExclamation, "New BOM"
GoTo CleanExit
End If

If PnRev_Exists_InBoms(loBoms, taPn, taRev) Then
MsgBox "PN/Revision combination already exists in BOMS." & vbCrLf & _
               "TAPN=" & taPn & ", TARev=" & taRev, vbExclamation, "New BOM"
GoTo CleanExit
End If

    ' Optional Comps validation (best-effort)
Set loComps = Nothing
On Error Resume Next
Set wsComps = wb.Worksheets(SH_COMPS)
If Not wsComps Is Nothing Then Set loComps = wsComps.ListObjects(LO_COMPS)
On Error GoTo EH

If Not loComps Is Nothing Then
If ColumnExists(loComps, "CompID") And ColumnExists(loComps, "OurPN") And ColumnExists(loComps, "OurRev") Then
Dim foundTop As Boolean
Dim topRevStatus As String
foundTop = Comps_FindByCompId(loComps, taId, taPn, taRev, topRevStatus)

If foundTop Then
If ColumnExists(loComps, "RevStatus") Then
If StrComp(Trim$(topRevStatus), ACTIVE_REVSTATUS, vbTextCompare) <> 0 Then
MsgBox "Top assembly exists in Comps but RevStatus is not '" & ACTIVE_REVSTATUS & "'." & vbCrLf & _
                           "TAID=" & taId & ", RevStatus=" & topRevStatus, vbExclamation, "New BOM"
GoTo CleanExit
End If
End If
End If
End If
End If

    ' Generate BOMID
bomId = GenerateNextId(loBoms, "BOMID", BOM_ID_PREFIX, BOM_ID_PAD)
If Len(bomId) = 0 Then Err.Raise vbObjectError + 6100, PROC_NAME, "Failed to generate BOMID."

    ' Copy template sheet
wsTemplate.Copy After:=wb.Sheets(wb.Sheets.Count)
Set wsNew = wb.Worksheets(wb.Worksheets.Count)

newSheetName = BOM_TAB_PREFIX & taId
If StrComp(NormalizeSheetName(newSheetName), newSheetName, vbBinaryCompare) <> 0 Then
MsgBox "TAID contains characters that cannot be used in BOM tab names." & vbCrLf & _
               "TAID=" & taId, vbExclamation, "New BOM"
GoTo CleanExit
End If
If WorksheetExists(wb, newSheetName) Then
MsgBox "A BOM tab named " & newSheetName & " already exists. TAID must map 1:1 to tab name." & vbCrLf & _
               "Choose a different TAID.", vbExclamation, "New BOM"
GoTo CleanExit
End If
wsNew.Name = newSheetName

    ' Rename the copied BOM table
If wsNew.ListObjects.Count < 1 Then Err.Raise vbObjectError + 6202, PROC_NAME, "No table found on copied BOM sheet."
Set loNew = wsNew.ListObjects(1)

newTableName = BuildUniqueTableName(wb, "TBL_BOM_" & NormalizeName(taId))
loNew.Name = newTableName

    ' Populate TA header fields on new BOM sheet
wsNew.Range("C1").Value = taId
wsNew.Range("C2").Value = taPn
wsNew.Range("C3").Value = taRev
wsNew.Range("C4").Value = taDesc

    ' Register in BOMS table
Dim lr As ListRow
Set lr = loBoms.ListRows.Add

createdAt = Now
createdBy = GetUserNameSafe()

SetByHeader loBoms, lr, "BOMID", bomId
SetByHeader loBoms, lr, "BOMTab", newSheetName
SetByHeader loBoms, lr, "TAID", taId
SetByHeader loBoms, lr, "BOM_NOTES", bomNotes
If ColumnExists(loBoms, "TARev") Then SetByHeader loBoms, lr, "TARev", taRev
If ColumnExists(loBoms, "TADesc") Then SetByHeader loBoms, lr, "TADesc", taDesc
If ColumnExists(loBoms, "TAPN") Then SetByHeader loBoms, lr, "TAPN", taPn

If ColumnExists(loBoms, "CreatedAt") Then SetByHeader loBoms, lr, "CreatedAt", createdAt
If ColumnExists(loBoms, "CreatedBy") Then SetByHeader loBoms, lr, "CreatedBy", createdBy
If ColumnExists(loBoms, "UpdatedAt") Then SetByHeader loBoms, lr, "UpdatedAt", createdAt
If ColumnExists(loBoms, "UpdatedBy") Then SetByHeader loBoms, lr, "UpdatedBy", createdBy

MsgBox "New BOM created: " & bomId & vbCrLf & _
          "Sheet: " & newSheetName & vbCrLf & _
          "TAID: " & taId & vbCrLf & _
          "PN/Rev: " & taPn & " / " & taRev, vbInformation, "New BOM"

CleanExit:
Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errLine As Long

    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl

    Debug_Report PROC_NAME, errNum, errDesc, errLine, _
        "TAID=" & taId & vbCrLf & _
        "TAPN=" & taPn & vbCrLf & _
        "TARev=" & taRev & vbCrLf & _
        "TADesc=" & taDesc & vbCrLf & _
        "BOM_NOTES=" & bomNotes & vbCrLf & _
        "TemplateSheet=" & SH_TEMPLATE & vbCrLf & _
        "BomsSheet=" & SH_BOMS & vbCrLf & _
        "ActiveSheet=" & SafeSheetNameSafe() & vbCrLf & _
        "Workbook=" & ThisWorkbook.Name
    Resume CleanExit
End Sub

'==========================
' Safe conversions
'==========================
Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = vbNullString
    ElseIf IsNull(v) Then
        SafeText = vbNullString
    Else
        SafeText = Trim$(CStr(v))
    End If
End Function

Private Function PromptText(ByVal prompt As String, ByVal title As String) As String
    Dim response As String
    response = InputBox(prompt, title)
    PromptText = Trim$(response)
End Function

Private Function SafeSheetNameSafe() As String
    On Error GoTo EH
    SafeSheetNameSafe = ActiveSheet.Name
    Exit Function
EH:
    SafeSheetNameSafe = "(unknown)"
End Function

'==========================
' Gate wrapper
'==========================
Private Function GateReady_Safe(Optional ByVal showUserMessage As Boolean = True) As Boolean
    On Error GoTo EH
    GateReady_Safe = M_Core_Gate.Gate_Ready(showUserMessage)
    Exit Function
EH:
    GateReady_Safe = False
End Function

'==========================
' Uniqueness checks
'==========================
Private Function TaId_Exists(ByVal loBoms As ListObject, ByVal taId As String) As Boolean
    Dim idx As Long, arr As Variant, i As Long, s As String, rowCount As Long
    TaId_Exists = False

If loBoms Is Nothing Then Exit Function
If loBoms.DataBodyRange Is Nothing Then Exit Function

idx = GetColIndex(loBoms, "TAID")
If idx = 0 Then Exit Function

arr = loBoms.ListColumns(idx).DataBodyRange.value
rowCount = ColumnDataRowCount(arr)
For i = 1 To rowCount
s = SafeText(ColumnDataItem(arr, i))
If Len(s) > 0 Then
If StrComp(s, taId, vbTextCompare) = 0 Then
TaId_Exists = True
Exit Function
End If
End If
Next i
End Function

Private Function PnRev_Exists_InBoms(ByVal loBoms As ListObject, ByVal pn As String, ByVal rev As String) As Boolean
    Dim idxPn As Long, idxRev As Long
    Dim arrPn As Variant, arrRev As Variant
    Dim idx As Long, arr As Variant, i As Long, notes As String
    Dim rowCount As Long
    Dim tokenPn As String, tokenRev As String

    PnRev_Exists_InBoms = False
If loBoms Is Nothing Then Exit Function
If loBoms.DataBodyRange Is Nothing Then Exit Function

    ' Preferred: structured columns if present
idxPn = GetColIndex(loBoms, "TAPN")
idxRev = GetColIndex(loBoms, "TARev")
If idxPn > 0 And idxRev > 0 Then
arrPn = loBoms.ListColumns(idxPn).DataBodyRange.Value
arrRev = loBoms.ListColumns(idxRev).DataBodyRange.Value
rowCount = ColumnDataRowCount(arrPn)
For i = 1 To rowCount
If StrComp(SafeText(ColumnDataItem(arrPn, i)), pn, vbTextCompare) = 0 And _
               StrComp(SafeText(ColumnDataItem(arrRev, i)), rev, vbTextCompare) = 0 Then
PnRev_Exists_InBoms = True
Exit Function
End If
Next i
Exit Function
End If

    ' Backward-compat fallback: legacy encoded values in BOM_NOTES
idx = GetColIndex(loBoms, "BOM_NOTES")
If idx = 0 Then Exit Function

tokenPn = "PN=" & pn & ";"
tokenRev = "Rev=" & rev & ";"

arr = loBoms.ListColumns(idx).DataBodyRange.Value
rowCount = ColumnDataRowCount(arr)
For i = 1 To rowCount
notes = SafeText(ColumnDataItem(arr, i))
If Len(notes) > 0 Then
If InStr(1, notes, tokenPn, vbTextCompare) > 0 And InStr(1, notes, tokenRev, vbTextCompare) > 0 Then
PnRev_Exists_InBoms = True
Exit Function
End If
End If
Next i
End Function

'==========================
' Optional Comps validation
'==========================
Private Function Comps_FindByCompId(ByVal loComps As ListObject, ByVal compId As String, ByVal ourPnIn As String, ByVal ourRevIn As String, ByRef revStatusOut As String) As Boolean
    Dim idxId As Long, idxPn As Long, idxRev As Long, idxRS As Long
    Dim arrId As Variant, arrPn As Variant, arrRev As Variant, arrRS As Variant
    Dim i As Long, rowCount As Long

    Comps_FindByCompId = False
    revStatusOut = vbNullString

If loComps Is Nothing Then Exit Function
If loComps.DataBodyRange Is Nothing Then Exit Function

idxId = GetColIndex(loComps, "CompID")
idxPn = GetColIndex(loComps, "OurPN")
idxRev = GetColIndex(loComps, "OurRev")
idxRS = GetColIndex(loComps, "RevStatus") 'may be 0

If idxId = 0 Or idxPn = 0 Or idxRev = 0 Then Exit Function

arrId = loComps.ListColumns(idxId).DataBodyRange.value
arrPn = loComps.ListColumns(idxPn).DataBodyRange.value
arrRev = loComps.ListColumns(idxRev).DataBodyRange.value
If idxRS > 0 Then arrRS = loComps.ListColumns(idxRS).DataBodyRange.value

rowCount = ColumnDataRowCount(arrId)
For i = 1 To rowCount
If StrComp(SafeText(ColumnDataItem(arrId, i)), compId, vbTextCompare) = 0 Then
If StrComp(SafeText(ColumnDataItem(arrPn, i)), ourPnIn, vbTextCompare) <> 0 Or _
               StrComp(SafeText(ColumnDataItem(arrRev, i)), ourRevIn, vbTextCompare) <> 0 Then
MsgBox "TAID exists in Comps but PN/Rev does not match your input." & vbCrLf & _
                       "Comps says: " & SafeText(ColumnDataItem(arrPn, i)) & " / " & SafeText(ColumnDataItem(arrRev, i)) & vbCrLf & _
                       "You entered: " & ourPnIn & " / " & ourRevIn, vbExclamation, "New BOM"
Exit Function
End If

If idxRS > 0 Then revStatusOut = SafeText(ColumnDataItem(arrRS, i))
Comps_FindByCompId = True
Exit Function
End If
Next i
End Function

'==========================
' Name builders
'==========================
Private Function BuildUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long

candidate = NormalizeSheetName(baseName)
If Not WorksheetExists(wb, candidate) Then
BuildUniqueSheetName = candidate
Exit Function
End If

suffix = 1
Do
candidate = NormalizeSheetName(baseName & "_" & CStr(suffix))
If Not WorksheetExists(wb, candidate) Then
BuildUniqueSheetName = candidate
Exit Function
End If
suffix = suffix + 1
Loop
End Function

Private Function NormalizeSheetName(ByVal nameIn As String) As String
    Dim outName As String
    outName = Trim$(nameIn)

    outName = Replace(outName, ":", "-")
    outName = Replace(outName, "\", "-")
    outName = Replace(outName, "/", "-")
    outName = Replace(outName, "?", "")
    outName = Replace(outName, "*", "")
    outName = Replace(outName, "[", "(")
    outName = Replace(outName, "]", ")")

    If Len(outName) = 0 Then outName = "BOM"
    If Len(outName) > 31 Then outName = Left$(outName, 31)
    NormalizeSheetName = outName
End Function

Private Function BuildUniqueTableName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long

candidate = NormalizeName(baseName)
If Not TableExists(wb, candidate) Then
BuildUniqueTableName = candidate
Exit Function
End If

suffix = 1
Do
candidate = NormalizeName(baseName & "_" & CStr(suffix))
If Not TableExists(wb, candidate) Then
BuildUniqueTableName = candidate
Exit Function
End If
suffix = suffix + 1
Loop
End Function

Private Function NormalizeName(ByVal nameIn As String) As String
    Dim outName As String
    outName = Trim$(nameIn)
    outName = Replace(outName, "-", "_")
    outName = Replace(outName, " ", "_")
    outName = Replace(outName, ".", "_")
    outName = Replace(outName, ":", "_")
    outName = Replace(outName, "/", "_")
    outName = Replace(outName, "\", "_")
    NormalizeName = outName
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

For Each ws In wb.Worksheets
For Each lo In ws.ListObjects
If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
TableExists = True
Exit Function
End If
Next lo
Next ws
TableExists = False
End Function

'==========================
' Table utilities
'==========================
Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 6200, "RequireColumn", "Missing column '" & header & "' in table '" & lo.Name & "'."
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

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
idx = GetColIndex(lo, header)
If idx = 0 Then Err.Raise vbObjectError + 6201, "SetByHeader", "Missing column '" & header & "' in table '" & lo.Name & "'."
lr.Range.Cells(1, idx).value = v
End Sub

'==========================
' ID generation
'==========================
Private Function GenerateNextId(ByVal lo As ListObject, ByVal header As String, ByVal prefix As String, ByVal padDigits As Long) As String
    Dim idx As Long, maxN As Long
    Dim arr As Variant
    Dim i As Long, s As String, n As Long, rowCount As Long

    GenerateNextId = vbNullString
idx = GetColIndex(lo, header)
If idx = 0 Then Exit Function

maxN = 0
If Not lo.DataBodyRange Is Nothing Then
arr = lo.ListColumns(idx).DataBodyRange.value
rowCount = ColumnDataRowCount(arr)
For i = 1 To rowCount
s = SafeText(ColumnDataItem(arr, i))
If Len(s) > 0 Then
n = TrailingNumber(s)
If n > maxN Then maxN = n
End If
Next i
End If

GenerateNextId = prefix & Right$(String$(padDigits, "0") & CStr(maxN + 1), padDigits)
End Function

Private Function ColumnDataRowCount(ByVal arr As Variant) As Long
    On Error GoTo EH
    If IsArray(arr) Then
        ColumnDataRowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    ElseIf IsEmpty(arr) Then
        ColumnDataRowCount = 0
    Else
        ColumnDataRowCount = 1
    End If
    Exit Function
EH:
    ColumnDataRowCount = 0
End Function

Private Function ColumnDataItem(ByVal arr As Variant, ByVal rowIndex As Long) As Variant
    If rowIndex < 1 Then
        ColumnDataItem = Empty
        Exit Function
    End If

    If IsArray(arr) Then
        ColumnDataItem = arr(LBound(arr, 1) + rowIndex - 1, LBound(arr, 2))
    ElseIf rowIndex = 1 Then
        ColumnDataItem = arr
    Else
        ColumnDataItem = Empty
    End If
End Function

Private Function TrailingNumber(ByVal s As String) As Long
    Dim i As Long, ch As String, digits As String
    digits = vbNullString

For i = Len(s) To 1 Step -1
ch = Mid$(s, i, 1)
If ch Like "#" Then
digits = ch & digits
Else
Exit For
End If
Next i

If Len(digits) = 0 Then
TrailingNumber = 0
Else
If IsNumeric(digits) Then
TrailingNumber = CLng(digits)
Else
TrailingNumber = 0
End If
End If
End Function

Private Function GetUserNameSafe() As String
    Dim u As String
u = Trim$(Environ$("Username"))
If Len(u) = 0 Then u = Application.userName
If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
GetUserNameSafe = u
End Function
