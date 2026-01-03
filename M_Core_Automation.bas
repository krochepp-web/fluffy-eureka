Attribute VB_Name = "M_Core_Automation"

Option Explicit

'===============================================================================
' Purpose:
'   Refresh the Automation Registry (Auto.TBL_AUTO) by scanning the VBA project
'   for ALL public procedures (Public Sub/Function) in standard modules.
'
'   This macro is the anti-drift mechanism:
'     - Upserts one row per discovered public procedure (key = Public Entry Point)
'     - Overwrites scanner-owned columns for consistency (by design)
'     - Optionally flags stale rows when a proc no longer exists in code
'
' Inputs (Tabs/Tables/Headers):
'   - Sheet: Auto
'   - Table: TBL_AUTO
'   - Required column (one of these must exist):
'       "Public Entry Point"  (preferred)
'       "PublicEntryPoint"
'       "EntryPoint"
'       "Macro"
'   - Optional columns populated if present:
'       "Module"
'       "Status"
'       "Trigger"
'       "Feature"
'       "FeatureName"
'       "Notes/Version"
'       "CreatedAt", "CreatedBy", "UpdatedAt", "UpdatedBy"
'
' Outputs / Side effects:
'   - Writes/updates rows in Auto.TBL_AUTO
'   - Optionally flags stale entries (Status="STALE") without deleting rows
'   - Logs via M_Core_Logging.LogEvent if present
'
' Preconditions:
'   - Trust Center enabled:
'       Trust access to the VBA project object model
'   - Auto sheet and TBL_AUTO exist
'
' Postconditions:
'   - TBL_AUTO contains a complete, current inventory of all public procedures
'
' Errors & Guards:
'   - Fails gracefully if VBProject access is blocked
'   - Fails fast if Auto/TBL_AUTO missing
'
' Version: v1.2.0
' Author: Keith + ChatGPT
' Date: 2025-12-20
'===============================================================================

'-----------------------------
' Public entry points
'-----------------------------

' Macro-dialog / button friendly entry point
Public Sub UI_RefreshAutomationRegistry()
    Dev_RefreshAutomationRegistry False, True, True
End Sub

' Developer entry point with controls
Public Sub Dev_RefreshAutomationRegistry( _
    Optional ByVal dryRun As Boolean = False, _
    Optional ByVal showUserMessage As Boolean = True, _
    Optional ByVal flagStale As Boolean = True)

    Const PROC_NAME As String = "Dev_RefreshAutomationRegistry"

    Dim wb As Workbook
    Dim lo As ListObject

    Dim colEntry As Long
    Dim colModule As Long, colStatus As Long, colTrigger As Long
    Dim colFeature As Long, colFeatureName As Long, colNotesVer As Long
    Dim colCreatedAt As Long, colCreatedBy As Long, colUpdatedAt As Long, colUpdatedBy As Long

    Dim dicFound As Object      ' key=ProcName, val=ModuleName
    Dim dicTrig As Object       ' key=ProcName, val=Trigger classification
    Dim dicExisting As Object   ' key=ProcName, val=ListRow index (1-based)

    Dim nowStamp As Date
    Dim userName As String

    Dim insertedCount As Long, updatedCount As Long, staleCount As Long

    On Error GoTo EH

    Set wb = ThisWorkbook
    nowStamp = Now
    userName = Environ$("Username")
    If Len(Trim$(userName)) = 0 Then userName = "UNKNOWN"

    ' Get table (no SafeGetListObject dependency)
    Set lo = Lo_GetByName(wb, "Auto", "TBL_AUTO")
    If lo Is Nothing Then
        Err.Raise vbObjectError + 510, PROC_NAME, "Missing table TBL_AUTO on sheet 'Auto'."
    End If

    ' Resolve columns
    colEntry = Lo_ResolveCol(lo, Array("Public Entry Point", "PublicEntryPoint", "EntryPoint", "Macro"), True)
    colModule = Lo_ResolveCol(lo, Array("Module", "ModuleName"), False)
    colStatus = Lo_ResolveCol(lo, Array("Status"), False)
    colTrigger = Lo_ResolveCol(lo, Array("Trigger", "Triggers"), False)
    colFeature = Lo_ResolveCol(lo, Array("Feature"), False)
    colFeatureName = Lo_ResolveCol(lo, Array("FeatureName"), False)
    colNotesVer = Lo_ResolveCol(lo, Array("Notes/Version", "Notes / Version", "Notes"), False)

    colCreatedAt = Lo_ResolveCol(lo, Array("CreatedAt", "Created At"), False)
    colCreatedBy = Lo_ResolveCol(lo, Array("CreatedBy", "Created By"), False)
    colUpdatedAt = Lo_ResolveCol(lo, Array("UpdatedAt", "Updated At"), False)
    colUpdatedBy = Lo_ResolveCol(lo, Array("UpdatedBy", "Updated By"), False)

    ' Build dictionaries from code
    Set dicFound = CreateObject("Scripting.Dictionary")
    dicFound.compareMode = vbTextCompare

    Set dicTrig = CreateObject("Scripting.Dictionary")
    dicTrig.compareMode = vbTextCompare

    BuildPublicProcDictionary wb, dicFound, dicTrig

    ' Index existing rows
    Set dicExisting = CreateObject("Scripting.Dictionary")
    dicExisting.compareMode = vbTextCompare
    BuildExistingIndex lo, colEntry, dicExisting

    ' Upsert
    insertedCount = 0
    updatedCount = 0
    UpsertRows lo, dicFound, dicTrig, dicExisting, _
               colEntry, colModule, colStatus, colTrigger, colFeature, colFeatureName, colNotesVer, _
               colCreatedAt, colCreatedBy, colUpdatedAt, colUpdatedBy, _
               nowStamp, userName, dryRun, insertedCount, updatedCount

    ' Stale flagging
    staleCount = 0
    If flagStale Then
        staleCount = FlagStaleRows(lo, dicFound, colEntry, colStatus, colNotesVer, nowStamp, userName, colUpdatedAt, colUpdatedBy, dryRun)
    End If

    TryLog PROC_NAME, 0, "Automation registry refresh complete", _
           "Found=" & dicFound.Count & "; Inserted=" & insertedCount & "; Updated=" & updatedCount & _
           "; FlagStale=" & CStr(flagStale) & "; StaleFlagged=" & staleCount & "; DryRun=" & CStr(dryRun)

    If showUserMessage Then
        MsgBox "Automation Registry Refresh" & vbCrLf & _
               String(30, "-") & vbCrLf & _
               "Public procedures found: " & dicFound.Count & vbCrLf & _
               "Inserted: " & insertedCount & vbCrLf & _
               "Updated: " & updatedCount & vbCrLf & _
               IIf(flagStale, "Stale flagged: " & staleCount & vbCrLf, vbNullString) & _
               "DryRun: " & CStr(dryRun), _
               vbInformation, "Registry Refresh"
    End If

CleanExit:
    Exit Sub

EH:
    TryLog PROC_NAME, Err.Number, Err.Description, "Registry refresh failed."
    If showUserMessage Then
        MsgBox "Registry refresh failed." & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "If this mentions VBProject access, enable:" & vbCrLf & _
               "Trust Center > Macro Settings > Trust access to the VBA project object model.", _
               vbExclamation, "Registry Refresh"
    End If
    Resume CleanExit
End Sub

'-----------------------------
' Core logic
'-----------------------------

Private Sub UpsertRows( _
    ByVal lo As ListObject, _
    ByVal dicFound As Object, _
    ByVal dicTrig As Object, _
    ByVal dicExisting As Object, _
    ByVal colEntry As Long, _
    ByVal colModule As Long, _
    ByVal colStatus As Long, _
    ByVal colTrigger As Long, _
    ByVal colFeature As Long, _
    ByVal colFeatureName As Long, _
    ByVal colNotesVer As Long, _
    ByVal colCreatedAt As Long, _
    ByVal colCreatedBy As Long, _
    ByVal colUpdatedAt As Long, _
    ByVal colUpdatedBy As Long, _
    ByVal nowStamp As Date, _
    ByVal userName As String, _
    ByVal dryRun As Boolean, _
    ByRef insertedCount As Long, _
    ByRef updatedCount As Long)

    Dim key As Variant
    Dim procName As String
    Dim moduleName As String
    Dim trig As String

    Dim lr As ListRow
    Dim isNew As Boolean
    Dim existingIdx As Long

    For Each key In dicFound.Keys
        procName = CStr(key)
        moduleName = CStr(dicFound(procName))
        trig = CStr(dicTrig(procName))

        isNew = False
        If dicExisting.Exists(procName) Then
            existingIdx = CLng(dicExisting(procName))
            Set lr = lo.ListRows(existingIdx)
        Else
            isNew = True
            If Not dryRun Then Set lr = lo.ListRows.Add
            insertedCount = insertedCount + 1
        End If

        If Not dryRun Then
            If isNew Then
                lr.Range.Cells(1, colEntry).value = procName
                If colCreatedAt > 0 Then lr.Range.Cells(1, colCreatedAt).value = nowStamp
                If colCreatedBy > 0 Then lr.Range.Cells(1, colCreatedBy).value = userName
            End If

            ' Scanner-owned fields: overwrite by design
            If colModule > 0 Then lr.Range.Cells(1, colModule).value = moduleName
            If colTrigger > 0 Then lr.Range.Cells(1, colTrigger).value = trig
            If colStatus > 0 Then lr.Range.Cells(1, colStatus).value = "ACTIVE"

            ' Feature + FeatureName: keep simple and deterministic for now
            If colFeature > 0 Then lr.Range.Cells(1, colFeature).value = procName
            If colFeatureName > 0 Then lr.Range.Cells(1, colFeatureName).value = procName

            If colNotesVer > 0 Then
                ' Do not destroy any richer, hand-authored notes if present.
                ' If blank, populate a minimal auto note.
                If Len(Trim$(CStr(lr.Range.Cells(1, colNotesVer).value))) = 0 Then
                    lr.Range.Cells(1, colNotesVer).value = "AUTO: scanned public proc (inventory baseline)"
                End If
            End If

            If colUpdatedAt > 0 Then lr.Range.Cells(1, colUpdatedAt).value = nowStamp
            If colUpdatedBy > 0 Then lr.Range.Cells(1, colUpdatedBy).value = userName
        End If

        If Not isNew Then updatedCount = updatedCount + 1

        ' Keep dicExisting consistent in-session if we inserted
        If isNew And Not dryRun Then
            dicExisting(procName) = lo.ListRows.Count
        End If
    Next key
End Sub

Private Function FlagStaleRows( _
    ByVal lo As ListObject, _
    ByVal dicFound As Object, _
    ByVal colEntry As Long, _
    ByVal colStatus As Long, _
    ByVal colNotesVer As Long, _
    ByVal nowStamp As Date, _
    ByVal userName As String, _
    ByVal colUpdatedAt As Long, _
    ByVal colUpdatedBy As Long, _
    ByVal dryRun As Boolean) As Long

    Dim i As Long
    Dim procName As String
    Dim flagged As Long

    flagged = 0
    If lo.DataBodyRange Is Nothing Then
        FlagStaleRows = 0
        Exit Function
    End If

    For i = 1 To lo.DataBodyRange.Rows.Count
        procName = Trim$(CStr(lo.DataBodyRange.Cells(i, colEntry).value))
        If Len(procName) > 0 Then
            If Not dicFound.Exists(procName) Then
                flagged = flagged + 1
                If Not dryRun Then
                    If colStatus > 0 Then lo.DataBodyRange.Cells(i, colStatus).value = "STALE"
                    If colNotesVer > 0 Then
                        lo.DataBodyRange.Cells(i, colNotesVer).value = _
                            "STALE: no longer found in code as of " & Format$(nowStamp, "yyyy-mm-dd hh:nn")
                    End If
                    If colUpdatedAt > 0 Then lo.DataBodyRange.Cells(i, colUpdatedAt).value = nowStamp
                    If colUpdatedBy > 0 Then lo.DataBodyRange.Cells(i, colUpdatedBy).value = userName
                End If
            End If
        End If
    Next i

    FlagStaleRows = flagged
End Function

Private Sub BuildExistingIndex(ByVal lo As ListObject, ByVal colEntry As Long, ByVal dicExisting As Object)
    Dim i As Long
    Dim key As String

    If lo.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To lo.DataBodyRange.Rows.Count
        key = Trim$(CStr(lo.DataBodyRange.Cells(i, colEntry).value))
        If Len(key) > 0 Then
            If Not dicExisting.Exists(key) Then dicExisting.Add key, i
        End If
    Next i
End Sub

'-----------------------------
' VBProject scanning
'-----------------------------

Private Sub BuildPublicProcDictionary(ByVal wb As Workbook, ByVal dicFound As Object, ByVal dicTrig As Object)
    Const PROC_NAME As String = "BuildPublicProcDictionary"

    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object

    On Error GoTo VBAccessBlocked
    Set vbProj = wb.VBProject

    For Each vbComp In vbProj.VBComponents
        ' Standard modules only (Type=1)
        If vbComp.Type = 1 Then
            Set codeMod = vbComp.CodeModule
            ScanCodeModule vbComp.Name, codeMod, dicFound, dicTrig
        End If
    Next vbComp

    Exit Sub

VBAccessBlocked:
    Err.Raise vbObjectError + 520, PROC_NAME, _
              "VBProject access blocked. Enable Trust Center setting: 'Trust access to the VBA project object model'."
End Sub

Private Sub ScanCodeModule(ByVal moduleName As String, ByVal codeMod As Object, ByVal dicFound As Object, ByVal dicTrig As Object)
    Dim nLines As Long
    Dim i As Long
    Dim lineText As String
    Dim decl As String
    Dim procName As String

    nLines = codeMod.CountOfLines
    i = 1

    Do While i <= nLines
        lineText = Trim$(codeMod.lines(i, 1))

        If Len(lineText) = 0 Then GoTo NextLine
        If Left$(lineText, 1) = "'" Then GoTo NextLine
        If Left$(LCase$(lineText), 9) = "attribute" Then GoTo NextLine

        If StartsWithPublicProc(lineText) Then
            decl = lineText

            ' Continuation lines "_"
            Do While (Right$(Trim$(decl), 1) = "_") And (i < nLines)
                decl = Left$(Trim$(decl), Len(Trim$(decl)) - 1) & " " & Trim$(codeMod.lines(i + 1, 1))
                i = i + 1
            Loop

            procName = ExtractProcNameFromDecl(decl)
            If Len(procName) > 0 Then
                If Not dicFound.Exists(procName) Then
                    dicFound.Add procName, moduleName
                    dicTrig.Add procName, ClassifyTrigger(procName)
                End If
            End If
        End If

NextLine:
        i = i + 1
    Loop
End Sub

Private Function StartsWithPublicProc(ByVal lineText As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(lineText))

    ' Accept:
    '   Public Sub ...
    '   Public Function ...
    ' Ignore:
    '   Public Property ...
    '   Private / Friend
    If Left$(t, 11) = "public sub " Then
        StartsWithPublicProc = True
    ElseIf Left$(t, 16) = "public function " Then
        StartsWithPublicProc = True
    Else
        StartsWithPublicProc = False
    End If
End Function

Private Function ExtractProcNameFromDecl(ByVal decl As String) As String
    ' decl examples:
    '   Public Sub Foo()
    '   Public Function Bar(ByVal x As Long) As Boolean
    Dim t As String
    Dim parts() As String
    Dim namePart As String

    t = Replace(Replace(Trim$(decl), vbTab, " "), "  ", " ")
    t = Replace(t, "Public Sub ", "", 1, 1, vbTextCompare)
    t = Replace(t, "Public Function ", "", 1, 1, vbTextCompare)

    parts = Split(t, " ")
    If UBound(parts) >= 0 Then
        namePart = parts(0)
        ' Strip "(" if present
        If InStr(1, namePart, "(", vbTextCompare) > 0 Then
            namePart = Left$(namePart, InStr(1, namePart, "(", vbTextCompare) - 1)
        End If
        ExtractProcNameFromDecl = Trim$(namePart)
    Else
        ExtractProcNameFromDecl = vbNullString
    End If
End Function

Private Function ClassifyTrigger(ByVal procName As String) As String
    ' Simple deterministic policy:
    '   UI_*   => User
    '   Dev_*  => Developer
    '   Test_* => Test
    '   Otherwise => Internal
    If Left$(procName, 3) = "UI_" Then
        ClassifyTrigger = "User"
    ElseIf Left$(procName, 4) = "Dev_" Then
        ClassifyTrigger = "Developer"
    ElseIf Left$(procName, 5) = "Test_" Then
        ClassifyTrigger = "Test"
    Else
        ClassifyTrigger = "Internal"
    End If
End Function

'-----------------------------
' Excel object helpers (no Select/Activate)
'-----------------------------

Private Function Lo_GetByName(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo CleanFail
    Set ws = wb.Worksheets(sheetName)
    Set lo = ws.ListObjects(tableName)
    Set Lo_GetByName = lo
    Exit Function

CleanFail:
    Set Lo_GetByName = Nothing
End Function

Private Function Lo_ResolveCol(ByVal lo As ListObject, ByVal candidates As Variant, ByVal required As Boolean) As Long
    Dim i As Long
    Dim nameTry As String
    Dim lc As ListColumn

    For i = LBound(candidates) To UBound(candidates)
        nameTry = CStr(candidates(i))
        For Each lc In lo.ListColumns
            If StrComp(lc.Name, nameTry, vbTextCompare) = 0 Then
                Lo_ResolveCol = lc.Index
                Exit Function
            End If
        Next lc
    Next i

    If required Then
        Err.Raise vbObjectError + 530, "Lo_ResolveCol", "Required column not found in " & lo.Name
    End If

    Lo_ResolveCol = 0
End Function

'-----------------------------
' Logging shim (avoid hard dependency)
'-----------------------------

Private Sub TryLog(ByVal procName As String, ByVal errNum As Long, ByVal msg As String, Optional ByVal details As String = "")
    On Error GoTo CleanExit
    ' If M_Core_Logging exists, this will work; otherwise falls through safely.
    Application.Run "LogEvent", procName, errNum, msg, details
CleanExit:
    On Error GoTo 0
End Sub


