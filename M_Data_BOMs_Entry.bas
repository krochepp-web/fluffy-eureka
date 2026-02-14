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
'   Creates sheet "BOM_<TAID>" (normalized; unique if needed) and registers row
'   in BOMS.TBL_BOMS. Includes line-numbered diagnostics for debugging Error 13.
'
' Inputs (Tabs/Tables/Headers):
'   - BOM_TEMPLATE: TBL_BOM_TEMPLATE
'   - BOMS: TBL_BOMS [BOMID, BOMTab, TAID, BOM_NOTES, TARev, TADesc] (+ optional TAPN)
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

10  If Not GateReady_Safe(True) Then GoTo CleanExit

20  Set wb = ThisWorkbook
30  Set wsTemplate = wb.Worksheets(SH_TEMPLATE)
40  Set wsBoms = wb.Worksheets(SH_BOMS)

50  Set loTemplate = wsTemplate.ListObjects(LO_TEMPLATE)
60  Set loBoms = wsBoms.ListObjects(LO_BOMS)

    ' Required headers
70  RequireColumn loTemplate, "CompID"
80  RequireColumn loTemplate, "OurPN"
90  RequireColumn loTemplate, "OurRev"
100 RequireColumn loTemplate, "Description"
110 RequireColumn loTemplate, "UOM"
120 RequireColumn loTemplate, "QtyPer"
130 RequireColumn loTemplate, "CompNotes"
140 RequireColumn loTemplate, "CreatedAt"
150 RequireColumn loTemplate, "CreatedBy"
160 RequireColumn loTemplate, "UpdatedAt"
170 RequireColumn loTemplate, "UpdatedBy"

180 RequireColumn loBoms, "BOMID"
190 RequireColumn loBoms, "BOMTab"
200 RequireColumn loBoms, "TAID"
210 RequireColumn loBoms, "BOM_NOTES"
220 RequireColumn loBoms, "TARev"
230 RequireColumn loBoms, "TADesc"

    ' Inputs
235 taId = Trim$(taId)
236 taPn = Trim$(taPn)
240 taRev = Trim$(taRev)
250 taDesc = Trim$(taDesc)
255 bomNotes = Trim$(bomNotes)

260 If Len(taId) = 0 Or Len(taPn) = 0 Or Len(taRev) = 0 Or Len(taDesc) = 0 Then
270     MsgBox "All fields are required (TAID, TAPN, TARev, TADesc).", vbExclamation, "New BOM"
280     GoTo CleanExit
290 End If

300 If TaId_Exists(loBoms, taId) Then
310     MsgBox "TAID '" & taId & "' already exists in BOMS (TAID). Choose a unique TAID.", vbExclamation, "New BOM"
320     GoTo CleanExit
330 End If

340 If PnRev_Exists_InBoms(loBoms, taPn, taRev) Then
350     MsgBox "PN/Revision combination already exists in BOMS." & vbCrLf & _
               "TAPN=" & taPn & ", TARev=" & taRev, vbExclamation, "New BOM"
360     GoTo CleanExit
370 End If

    ' Optional Comps validation (best-effort)
380 Set loComps = Nothing
390 On Error Resume Next
400 Set wsComps = wb.Worksheets(SH_COMPS)
410 If Not wsComps Is Nothing Then Set loComps = wsComps.ListObjects(LO_COMPS)
420 On Error GoTo EH

430 If Not loComps Is Nothing Then
440     If ColumnExists(loComps, "CompID") And ColumnExists(loComps, "OurPN") And ColumnExists(loComps, "OurRev") Then
450         Dim foundTop As Boolean
460         Dim topRevStatus As String
470         foundTop = Comps_FindByCompId(loComps, taId, taPn, taRev, topRevStatus)

480         If foundTop Then
490             If ColumnExists(loComps, "RevStatus") Then
500                 If StrComp(Trim$(topRevStatus), ACTIVE_REVSTATUS, vbTextCompare) <> 0 Then
510                     MsgBox "Top assembly exists in Comps but RevStatus is not '" & ACTIVE_REVSTATUS & "'." & vbCrLf & _
                           "TAID=" & taId & ", RevStatus=" & topRevStatus, vbExclamation, "New BOM"
520                     GoTo CleanExit
530                 End If
540             End If
550         End If
560     End If
570 End If

    ' Generate BOMID
580 bomId = GenerateNextId(loBoms, "BOMID", BOM_ID_PREFIX, BOM_ID_PAD)
590 If Len(bomId) = 0 Then Err.Raise vbObjectError + 6100, PROC_NAME, "Failed to generate BOMID."

    ' Copy template sheet
600 wsTemplate.Copy After:=wb.Sheets(wb.Sheets.Count)
610 Set wsNew = wb.Worksheets(wb.Worksheets.Count)

620 newSheetName = BuildUniqueSheetName(wb, BOM_TAB_PREFIX & taId)
630 wsNew.Name = newSheetName

    ' Populate TA description cell on the new BOM sheet
635 wsNew.Range("C4").Value = taDesc

    ' Rename the copied BOM table
640 If wsNew.ListObjects.Count < 1 Then Err.Raise vbObjectError + 6202, PROC_NAME, "No table found on copied BOM sheet."
650 Set loNew = wsNew.ListObjects(1)

660 newTableName = BuildUniqueTableName(wb, "TBL_BOM_" & NormalizeName(taId))
670 loNew.Name = newTableName

    ' Register in BOMS table
680 Dim lr As ListRow
690 Set lr = loBoms.ListRows.Add

700 createdAt = Now
710 createdBy = GetUserNameSafe()

720 SetByHeader loBoms, lr, "BOMID", bomId
730 SetByHeader loBoms, lr, "BOMTab", newSheetName
740 SetByHeader loBoms, lr, "TAID", taId
750 SetByHeader loBoms, lr, "BOM_NOTES", bomNotes
760 SetByHeader loBoms, lr, "TARev", taRev
770 SetByHeader loBoms, lr, "TADesc", taDesc
775 If ColumnExists(loBoms, "TAPN") Then SetByHeader loBoms, lr, "TAPN", taPn

780 If ColumnExists(loBoms, "CreatedAt") Then SetByHeader loBoms, lr, "CreatedAt", createdAt
790 If ColumnExists(loBoms, "CreatedBy") Then SetByHeader loBoms, lr, "CreatedBy", createdBy
800 If ColumnExists(loBoms, "UpdatedAt") Then SetByHeader loBoms, lr, "UpdatedAt", createdAt
810 If ColumnExists(loBoms, "UpdatedBy") Then SetByHeader loBoms, lr, "UpdatedBy", createdBy

820 MsgBox "New BOM created: " & bomId & vbCrLf & _
          "Sheet: " & newSheetName & vbCrLf & _
          "TAID: " & taId & vbCrLf & _
          "PN/Rev: " & taPn & " / " & taRev, vbInformation, "New BOM"

CleanExit:
830 Exit Sub

EH:
    Debug_Report PROC_NAME, Err, _
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
    Dim idx As Long, arr As Variant, i As Long, s As String
    TaId_Exists = False

900 If loBoms Is Nothing Then Exit Function
910 If loBoms.DataBodyRange Is Nothing Then Exit Function

920 idx = GetColIndex(loBoms, "TAID")
930 If idx = 0 Then Exit Function

940 arr = loBoms.ListColumns(idx).DataBodyRange.value
950 For i = 1 To UBound(arr, 1)
960     s = SafeText(arr(i, 1))
970     If Len(s) > 0 Then
980         If StrComp(s, taId, vbTextCompare) = 0 Then
990             TaId_Exists = True
1000            Exit Function
1010        End If
1020    End If
1030 Next i
End Function

Private Function PnRev_Exists_InBoms(ByVal loBoms As ListObject, ByVal pn As String, ByVal rev As String) As Boolean
    Dim idxPn As Long, idxRev As Long
    Dim arrPn As Variant, arrRev As Variant
    Dim idx As Long, arr As Variant, i As Long, notes As String
    Dim tokenPn As String, tokenRev As String

    PnRev_Exists_InBoms = False
1100 If loBoms Is Nothing Then Exit Function
1110 If loBoms.DataBodyRange Is Nothing Then Exit Function

    ' Preferred: structured columns if present
1120 idxPn = GetColIndex(loBoms, "TAPN")
1130 idxRev = GetColIndex(loBoms, "TARev")
1140 If idxPn > 0 And idxRev > 0 Then
1150     arrPn = loBoms.ListColumns(idxPn).DataBodyRange.Value
1160     arrRev = loBoms.ListColumns(idxRev).DataBodyRange.Value
1170     For i = 1 To UBound(arrPn, 1)
1180         If StrComp(SafeText(arrPn(i, 1)), pn, vbTextCompare) = 0 And _
               StrComp(SafeText(arrRev(i, 1)), rev, vbTextCompare) = 0 Then
1190             PnRev_Exists_InBoms = True
1200             Exit Function
1210         End If
1220     Next i
1230     Exit Function
1240 End If

    ' Backward-compat fallback: legacy encoded values in BOM_NOTES
1250 idx = GetColIndex(loBoms, "BOM_NOTES")
1260 If idx = 0 Then Exit Function

1270 tokenPn = "PN=" & pn & ";"
1280 tokenRev = "Rev=" & rev & ";"

1290 arr = loBoms.ListColumns(idx).DataBodyRange.Value
1300 For i = 1 To UBound(arr, 1)
1310     notes = SafeText(arr(i, 1))
1320     If Len(notes) > 0 Then
1330         If InStr(1, notes, tokenPn, vbTextCompare) > 0 And InStr(1, notes, tokenRev, vbTextCompare) > 0 Then
1340             PnRev_Exists_InBoms = True
1350             Exit Function
1360         End If
1370     End If
1380 Next i
End Function

'==========================
' Optional Comps validation
'==========================
Private Function Comps_FindByCompId(ByVal loComps As ListObject, ByVal compId As String, ByVal ourPnIn As String, ByVal ourRevIn As String, ByRef revStatusOut As String) As Boolean
    Dim idxId As Long, idxPn As Long, idxRev As Long, idxRS As Long
    Dim arrId As Variant, arrPn As Variant, arrRev As Variant, arrRS As Variant
    Dim i As Long

    Comps_FindByCompId = False
    revStatusOut = vbNullString

1300 If loComps Is Nothing Then Exit Function
1310 If loComps.DataBodyRange Is Nothing Then Exit Function

1320 idxId = GetColIndex(loComps, "CompID")
1330 idxPn = GetColIndex(loComps, "OurPN")
1340 idxRev = GetColIndex(loComps, "OurRev")
1350 idxRS = GetColIndex(loComps, "RevStatus") 'may be 0

1360 If idxId = 0 Or idxPn = 0 Or idxRev = 0 Then Exit Function

1370 arrId = loComps.ListColumns(idxId).DataBodyRange.value
1380 arrPn = loComps.ListColumns(idxPn).DataBodyRange.value
1390 arrRev = loComps.ListColumns(idxRev).DataBodyRange.value
1400 If idxRS > 0 Then arrRS = loComps.ListColumns(idxRS).DataBodyRange.value

1410 For i = 1 To UBound(arrId, 1)
1420     If StrComp(SafeText(arrId(i, 1)), compId, vbTextCompare) = 0 Then
1430         If StrComp(SafeText(arrPn(i, 1)), ourPnIn, vbTextCompare) <> 0 Or _
               StrComp(SafeText(arrRev(i, 1)), ourRevIn, vbTextCompare) <> 0 Then
1440             MsgBox "TAID exists in Comps but PN/Rev does not match your input." & vbCrLf & _
                       "Comps says: " & SafeText(arrPn(i, 1)) & " / " & SafeText(arrRev(i, 1)) & vbCrLf & _
                       "You entered: " & ourPnIn & " / " & ourRevIn, vbExclamation, "New BOM"
1450             Exit Function
1460         End If

1470         If idxRS > 0 Then revStatusOut = SafeText(arrRS(i, 1))
1480         Comps_FindByCompId = True
1490         Exit Function
1500     End If
1510 Next i
End Function

'==========================
' Name builders
'==========================
Private Function BuildUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long

1600 candidate = NormalizeSheetName(baseName)
1610 If Not WorksheetExists(wb, candidate) Then
1620     BuildUniqueSheetName = candidate
1630     Exit Function
1640 End If

1650 suffix = 1
1660 Do
1670     candidate = NormalizeSheetName(baseName & "_" & CStr(suffix))
1680     If Not WorksheetExists(wb, candidate) Then
1690         BuildUniqueSheetName = candidate
1700         Exit Function
1710     End If
1720     suffix = suffix + 1
1730 Loop
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

1800 candidate = NormalizeName(baseName)
1810 If Not TableExists(wb, candidate) Then
1820     BuildUniqueTableName = candidate
1830     Exit Function
1840 End If

1850 suffix = 1
1860 Do
1870     candidate = NormalizeName(baseName & "_" & CStr(suffix))
1880     If Not TableExists(wb, candidate) Then
1890         BuildUniqueTableName = candidate
1900         Exit Function
1910     End If
1920     suffix = suffix + 1
1930 Loop
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

2000 For Each ws In wb.Worksheets
2010     For Each lo In ws.ListObjects
2020         If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
2030             TableExists = True
2040             Exit Function
2050         End If
2060     Next lo
2070 Next ws
2080 TableExists = False
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
2100 For Each lc In lo.ListColumns
2110     If StrComp(lc.Name, header, vbTextCompare) = 0 Then
2120         GetColIndex = lc.Index
2130         Exit Function
2140     End If
2150 Next lc
2160 GetColIndex = 0
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
2200 idx = GetColIndex(lo, header)
2210 If idx = 0 Then Err.Raise vbObjectError + 6201, "SetByHeader", "Missing column '" & header & "' in table '" & lo.Name & "'."
2220 lr.Range.Cells(1, idx).value = v
End Sub

'==========================
' ID generation
'==========================
Private Function GenerateNextId(ByVal lo As ListObject, ByVal header As String, ByVal prefix As String, ByVal padDigits As Long) As String
    Dim idx As Long, maxN As Long
    Dim arr As Variant
    Dim i As Long, s As String, n As Long

    GenerateNextId = vbNullString
2300 idx = GetColIndex(lo, header)
2310 If idx = 0 Then Exit Function

2320 maxN = 0
2330 If Not lo.DataBodyRange Is Nothing Then
2340     arr = lo.ListColumns(idx).DataBodyRange.value
2350     For i = 1 To UBound(arr, 1)
2360         s = SafeText(arr(i, 1))
2370         If Len(s) > 0 Then
2380             n = TrailingNumber(s)
2390             If n > maxN Then maxN = n
2400         End If
2410     Next i
2420 End If

2430 GenerateNextId = prefix & Right$(String$(padDigits, "0") & CStr(maxN + 1), padDigits)
End Function

Private Function TrailingNumber(ByVal s As String) As Long
    Dim i As Long, ch As String, digits As String
    digits = vbNullString

2500 For i = Len(s) To 1 Step -1
2510     ch = Mid$(s, i, 1)
2520     If ch Like "#" Then
2530         digits = ch & digits
2540     Else
2550         Exit For
2560     End If
2570 Next i

2580 If Len(digits) = 0 Then
2590     TrailingNumber = 0
2600 Else
2610     If IsNumeric(digits) Then
2620         TrailingNumber = CLng(digits)
2630     Else
2640         TrailingNumber = 0
2650     End If
2660 End If
End Function

Private Function GetUserNameSafe() As String
    Dim u As String
2700 u = Trim$(Environ$("Username"))
2710 If Len(u) = 0 Then u = Application.userName
2720 If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
2730 GetUserNameSafe = u
End Function
