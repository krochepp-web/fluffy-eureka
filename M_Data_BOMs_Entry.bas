Attribute VB_Name = "M_Data_BOMs_Entry"
Option Explicit

'===============================================================================
' Module: M_Data_BOMs_Entry
'
' Purpose:
'   Create a new BOM sheet from BOM_TEMPLATE for a buildable top assembly (TA)
'   and register it in BOMS.TBL_BOMS.
'
' Inputs (Tabs/Tables/Headers):
'   - BOM_TEMPLATE sheet: TBL_BOM_TEMPLATE
'       Required headers:
'         CompID, OurPN, OurRev, Description, UOM, QtyPer, CompNotes,
'         CreatedAt, CreatedBy, UpdatedAt, UpdatedBy
'   - BOMS sheet: TBL_BOMS
'       Required headers:
'         BOMID, BOMTab, AssemblyID, BOM_NOTES
'   - Comps sheet: TBL_COMPS
'       Required headers:
'         CompID, IsBuildable
'
' Outputs / Side effects:
'   - Copies BOM_TEMPLATE to a new sheet
'   - Renames the copied sheet using the buildable BOM naming syntax
'   - Renames the BOM table on the copied sheet to a unique name
'   - Adds a row to BOMS.TBL_BOMS
'
' Preconditions / Postconditions:
'   -
'
' Errors & Guards:
'   - Fails fast on missing sheets/tables/headers
'   - Blocks creation if AssemblyID is not marked buildable in Comps
'
' Version: v0.1.0
' Author: ChatGPT (assistant)
' Date: 2025-01-04
'===============================================================================

'==========================
' PUBLIC ENTRY POINT
'==========================
Public Sub UI_Create_BOM_For_Assembly()
    Const PROC_NAME As String = "M_Data_BOMs_Entry.UI_Create_BOM_For_Assembly"

    Const SH_TEMPLATE As String = "BOM_TEMPLATE"
    Const LO_TEMPLATE As String = "TBL_BOM_TEMPLATE"

    Const SH_BOMS As String = "BOMS"
    Const LO_BOMS As String = "TBL_BOMS"

    Const SH_COMPS As String = "Comps"
    Const LO_COMPS As String = "TBL_COMPS"

    Const BOM_TAB_PREFIX As String = "BOM_BUILD_"
    Const BOM_ID_PREFIX As String = "BOM-"
    Const BOM_ID_PAD As Long = 4

    Dim wb As Workbook
    Dim wsTemplate As Worksheet
    Dim wsBoms As Worksheet
    Dim wsComps As Worksheet
    Dim wsNew As Worksheet

    Dim loTemplate As ListObject
    Dim loBoms As ListObject
    Dim loComps As ListObject
    Dim loNew As ListObject

    Dim assemblyId As String
    Dim bomNotes As String
    Dim bomId As String
    Dim newSheetName As String
    Dim newTableName As String
    Dim createdAt As Date
    Dim createdBy As String

    On Error GoTo EH

    If Not GateReady_Safe(True) Then Exit Sub

    Set wb = ThisWorkbook
    Set wsTemplate = wb.Worksheets(SH_TEMPLATE)
    Set wsBoms = wb.Worksheets(SH_BOMS)
    Set wsComps = wb.Worksheets(SH_COMPS)

    Set loTemplate = wsTemplate.ListObjects(LO_TEMPLATE)
    Set loBoms = wsBoms.ListObjects(LO_BOMS)
    Set loComps = wsComps.ListObjects(LO_COMPS)

    ' Guard required headers
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
    RequireColumn loBoms, "AssemblyID"
    RequireColumn loBoms, "BOM_NOTES"

    RequireColumn loComps, "CompID"
    RequireColumn loComps, "IsBuildable"

    assemblyId = Trim$(InputBox("Enter AssemblyID (CompID) for the new buildable BOM.", "New BOM"))
    If Len(assemblyId) = 0 Then Exit Sub

    If Not Assembly_IsBuildable(loComps, assemblyId) Then
        MsgBox "AssemblyID '" & assemblyId & "' is not marked buildable in Comps.", vbExclamation, "New BOM"
        Exit Sub
    End If

    bomNotes = Trim$(InputBox("Optional BOM notes (blank is ok).", "New BOM (" & assemblyId & ")"))

    ' Generate BOMID
    bomId = GenerateNextId(loBoms, "BOMID", BOM_ID_PREFIX, BOM_ID_PAD)
    If Len(bomId) = 0 Then Err.Raise vbObjectError + 6100, PROC_NAME, "Failed to generate BOMID."

    ' Copy template sheet
    wsTemplate.Copy After:=wb.Sheets(wb.Sheets.Count)
    Set wsNew = ActiveSheet

    newSheetName = BuildUniqueSheetName(wb, BOM_TAB_PREFIX & assemblyId)
    wsNew.Name = newSheetName

    Set loNew = wsNew.ListObjects(1)
    newTableName = BuildUniqueTableName(wb, "TBL_BOM_" & NormalizeName(assemblyId))
    loNew.Name = newTableName

    ' Register in BOMS table
    Dim lr As ListRow
    Set lr = loBoms.ListRows.Add

    createdAt = Now
    createdBy = GetUserNameSafe()

    SetByHeader loBoms, lr, "BOMID", bomId
    SetByHeader loBoms, lr, "BOMTab", newSheetName
    SetByHeader loBoms, lr, "AssemblyID", assemblyId
    SetByHeader loBoms, lr, "BOM_NOTES", bomNotes

    If ColumnExists(loBoms, "CreatedAt") Then SetByHeader loBoms, lr, "CreatedAt", createdAt
    If ColumnExists(loBoms, "CreatedBy") Then SetByHeader loBoms, lr, "CreatedBy", createdBy
    If ColumnExists(loBoms, "UpdatedAt") Then SetByHeader loBoms, lr, "UpdatedAt", createdAt
    If ColumnExists(loBoms, "UpdatedBy") Then SetByHeader loBoms, lr, "UpdatedBy", createdBy

    MsgBox "New BOM created: " & bomId & vbCrLf & _
           "Sheet: " & newSheetName, vbInformation, "New BOM"
    Exit Sub

EH:
    MsgBox "New BOM creation failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New BOM"
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
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New BOM"
    GateReady_Safe = False
End Function

Private Function Assembly_IsBuildable(ByVal loComps As ListObject, ByVal assemblyId As String) As Boolean
    Dim idxId As Long, idxBuild As Long
    Dim arrId As Variant, arrBuild As Variant
    Dim i As Long

    Assembly_IsBuildable = False
    If loComps Is Nothing Then Exit Function
    If loComps.DataBodyRange Is Nothing Then Exit Function

    idxId = GetColIndex(loComps, "CompID")
    idxBuild = GetColIndex(loComps, "IsBuildable")
    If idxId = 0 Or idxBuild = 0 Then Exit Function

    arrId = loComps.ListColumns(idxId).DataBodyRange.value
    arrBuild = loComps.ListColumns(idxBuild).DataBodyRange.value

    For i = 1 To UBound(arrId, 1)
        If StrComp(Trim$(CStr(arrId(i, 1))), assemblyId, vbTextCompare) = 0 Then
            Assembly_IsBuildable = IsTrueish(arrBuild(i, 1))
            Exit Function
        End If
    Next i
End Function

Private Function IsTrueish(ByVal v As Variant) As Boolean
    Dim s As String
    If IsNumeric(v) Then
        IsTrueish = (CLng(v) <> 0)
        Exit Function
    End If
    s = UCase$(Trim$(CStr(v)))
    IsTrueish = (s = "Y" Or s = "YES" Or s = "TRUE" Or s = "T" Or s = "1")
End Function

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

Private Function GenerateNextId(ByVal lo As ListObject, ByVal header As String, ByVal prefix As String, ByVal padDigits As Long) As String
    Dim idx As Long, maxN As Long
    Dim arr As Variant
    Dim i As Long, s As String, n As Long

    GenerateNextId = vbNullString
    idx = GetColIndex(lo, header)
    If idx = 0 Then Exit Function

    maxN = 0
    If Not lo.DataBodyRange Is Nothing Then
        arr = lo.ListColumns(idx).DataBodyRange.value
        For i = 1 To UBound(arr, 1)
            s = Trim$(CStr(arr(i, 1)))
            n = TrailingNumber(s)
            If n > maxN Then maxN = n
        Next i
    End If

    GenerateNextId = prefix & Right$(String$(padDigits, "0") & CStr(maxN + 1), padDigits)
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
        TrailingNumber = CLng(digits)
    End If
End Function

Private Function GetUserNameSafe() As String
    Dim u As String
    u = Trim$(Environ$("Username"))
    If Len(u) = 0 Then u = Application.userName
    If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
    GetUserNameSafe = u
End Function
