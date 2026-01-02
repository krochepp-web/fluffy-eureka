Attribute VB_Name = "M_Core_Lockdown"
Option Explicit

'===========================================================
' Purpose:
'   Workbook Lockdown (Excel 2016+):
'     - Lock all sheets by default
'     - Unlock only schema-authorized user entry cells in Input tables
'     - Protect sheets with UserInterfaceOnly:=True (VBA can write)
'     - Optionally protect workbook structure
'     - Produce Lockdown_Preview report
'
' Public entry points (Alt+F8):
'   - Lockdown_DryRun   (no changes; generates preview; NO PASSWORD PROMPT)
'   - Lockdown_Apply    (enforces protection + hides dev sheets; prompts password)
'   - Lockdown_Remove   (developer convenience; prompts password)
'
' Inputs:
'   - Sheet: SCHEMA
'   - Table: TBL_SCHEMA
'   Required headers:
'     TAB_NAME, TABLE_NAME, COLUMN_HEADER
'
' Unlock logic (priority):
'   1) If "UserEditable" exists and has at least one Y/TRUE: unlock where UserEditable=Trueish
'   2) Else if "EntryMethod" exists: unlock where EntryMethod contains "USER" or equals "MANUAL"
'
' Optional role filter:
'   - If "EditRole" exists and SCHEMA!B2 has CurrentRole,
'     unlocked fields restricted to allowed roles.
'
' Outputs:
'   - Sheet "Lockdown_Preview" is created/overwritten
'
' Version: v1.2.1
' Author: ChatGPT
' Date: 2025-12-21
'===========================================================

Private gStep As String

Public Sub Lockdown_DryRun()
    Lockdown_Run True
End Sub

Public Sub Lockdown_Apply()
    Lockdown_Run False
End Sub

Public Sub Lockdown_Remove()
    Const DEFAULT_PASSWORD As String = "CHANGE_ME"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pwd As String

    On Error GoTo EH
    Set wb = ThisWorkbook

    pwd = GetPassword(DEFAULT_PASSWORD)
    If Len(Trim$(pwd)) = 0 Then Err.Raise vbObjectError + 102, , "Password is blank or Cancelled."

    If wb.ProtectStructure Then wb.Unprotect Password:=pwd

    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=pwd
        ws.Visible = xlSheetVisible
        On Error GoTo EH
    Next ws

    MsgBox "Lockdown removed (developer mode).", vbInformation, "Lockdown"
    Exit Sub

EH:
    MsgBox "Lockdown_Remove failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Lockdown"
End Sub

Private Sub Lockdown_Run(ByVal dryRun As Boolean)
    Const DEFAULT_PASSWORD As String = "CHANGE_ME"
    Const WS_SCHEMA As String = "SCHEMA"
    Const LO_SCHEMA As String = "TBL_SCHEMA"

    Dim wb As Workbook
    Dim wsSchema As Worksheet
    Dim loSchema As ListObject
    Dim pwd As String
    Dim currentRole As String
    Dim plan As Object

    On Error GoTo EH
    Set wb = ThisWorkbook

    '-------------------------------------------------------
    ' IMPORTANT: DryRun does NOT prompt for password.
    '-------------------------------------------------------
    If dryRun Then
        pwd = vbNullString
    Else
        gStep = "GetPassword"
        pwd = GetPassword(DEFAULT_PASSWORD)
        If Len(Trim$(pwd)) = 0 Then Err.Raise vbObjectError + 101, , "Password is blank or Cancelled."
    End If

    gStep = "Find SCHEMA sheet"
    If Not WorksheetExists(wb, WS_SCHEMA) Then
        Err.Raise vbObjectError + 200, , "Missing required sheet: " & WS_SCHEMA
    End If
    Set wsSchema = wb.Worksheets(WS_SCHEMA)

    gStep = "Find TBL_SCHEMA"
    If Not ListObjectExists(wsSchema, LO_SCHEMA) Then
        Err.Raise vbObjectError + 201, , "Missing required table: " & WS_SCHEMA & "!" & LO_SCHEMA
    End If
    Set loSchema = wsSchema.ListObjects(LO_SCHEMA)

    gStep = "Require minimal headers"
    RequireHeader loSchema, "TAB_NAME"
    RequireHeader loSchema, "TABLE_NAME"
    RequireHeader loSchema, "COLUMN_HEADER"

    gStep = "Read CurrentRole (optional: SCHEMA!B2)"
    currentRole = Trim$(CStr(wsSchema.Range("B2").value))

    gStep = "Build Unlock Plan"
    Set plan = BuildUnlockPlan(wb, loSchema, currentRole)

    gStep = "Write Preview"
    WritePreview wb, plan, currentRole

    If dryRun Then
        MsgBox "Lockdown dry run complete. Review 'Lockdown_Preview'.", vbInformation, "Lockdown"
        Exit Sub
    End If

    gStep = "Apply protections"
    ApplySheetProtections wb, plan, pwd
    HideDevSheets wb, pwd
    ApplyWorkbookStructure wb, pwd, True


    MsgBox "Lockdown applied.", vbInformation, "Lockdown"
    Exit Sub

EH:
    MsgBox "Lockdown failed at step: " & gStep & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Lockdown"
End Sub

'----------------------------
' Plan building
'----------------------------

Private Function BuildUnlockPlan(ByVal wb As Workbook, ByVal loSchema As ListObject, ByVal currentRole As String) As Object
    Dim plan As Object
    Set plan = CreateObject("Scripting.Dictionary")
    plan.compareMode = vbTextCompare

    If loSchema.DataBodyRange Is Nothing Then
        Set BuildUnlockPlan = plan
        Exit Function
    End If

    Dim hasUserEditable As Boolean, hasEntryMethod As Boolean, hasEditRole As Boolean
    hasUserEditable = HasHeader(loSchema, "UserEditable")
    hasEntryMethod = HasHeader(loSchema, "EntryMethod")
    hasEditRole = HasHeader(loSchema, "EditRole")

    Dim idxTab As Long, idxTbl As Long, idxCol As Long
    idxTab = loSchema.ListColumns("TAB_NAME").Index
    idxTbl = loSchema.ListColumns("TABLE_NAME").Index
    idxCol = loSchema.ListColumns("COLUMN_HEADER").Index

    Dim idxUE As Long, idxEM As Long, idxER As Long
    idxUE = 0: idxEM = 0: idxER = 0
    If hasUserEditable Then idxUE = loSchema.ListColumns("UserEditable").Index
    If hasEntryMethod Then idxEM = loSchema.ListColumns("EntryMethod").Index
    If hasEditRole Then idxER = loSchema.ListColumns("EditRole").Index

    Dim ueYCount As Long
    ueYCount = 0
    If hasUserEditable Then
        Dim rr As Long
        For rr = 1 To loSchema.ListRows.Count
            If IsTrueishText(CStr(loSchema.DataBodyRange.Cells(rr, idxUE).value)) Then ueYCount = ueYCount + 1
        Next rr
    End If

    Dim useUE As Boolean
    useUE = (hasUserEditable And ueYCount > 0)

    Dim r As Long
    For r = 1 To loSchema.ListRows.Count
        Dim tabN As String, tblN As String, colH As String
        tabN = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxTab).value))
        tblN = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxTbl).value))
        colH = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxCol).value))

        If Len(tabN) = 0 Or Len(tblN) = 0 Or Len(colH) = 0 Then GoTo NextR

        Dim eligible As Boolean
        eligible = False

        If useUE Then
            eligible = IsTrueishText(CStr(loSchema.DataBodyRange.Cells(r, idxUE).value))
        ElseIf hasEntryMethod Then
            eligible = EntryMethodIsUser(CStr(loSchema.DataBodyRange.Cells(r, idxEM).value))
        Else
            eligible = False
        End If

        If Not eligible Then GoTo NextR

        If Len(currentRole) > 0 And hasEditRole Then
            Dim allowed As String
            allowed = Trim$(CStr(loSchema.DataBodyRange.Cells(r, idxER).value))
            If Len(allowed) > 0 Then
                If Not RoleIsAllowed(currentRole, allowed) Then GoTo NextR
            End If
        End If

        Dim key As String
        key = tabN & "|" & tblN

        If Not plan.Exists(key) Then
            Dim cols As Collection
            Set cols = New Collection
            plan.Add key, cols
        End If

        If Not CollectionContains(plan(key), colH) Then plan(key).Add colH

NextR:
    Next r

    Set BuildUnlockPlan = plan
End Function

Private Function EntryMethodIsUser(ByVal s As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(CStr(s)))
    EntryMethodIsUser = (InStr(1, u, "USER", vbTextCompare) > 0 Or u = "MANUAL")
End Function

Private Function RoleIsAllowed(ByVal currentRole As String, ByVal allowedRoles As String) As Boolean
    Dim normCurrent As String
    normCurrent = UCase$(Trim$(currentRole))

    Dim tmp As String
    tmp = allowedRoles
    tmp = Replace(tmp, ",", "|")
    tmp = Replace(tmp, ";", "|")
    tmp = Replace(tmp, "/", "|")

    Dim parts() As String
    parts = Split(tmp, "|")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If UCase$(Trim$(parts(i))) = normCurrent Then
            RoleIsAllowed = True
            Exit Function
        End If
    Next i
    RoleIsAllowed = False
End Function

'----------------------------
' Apply protections
'----------------------------

Private Sub ApplySheetProtections(ByVal wb As Workbook, ByVal plan As Object, ByVal pwd As String)
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        ws.Unprotect Password:=pwd
        ws.Cells.Locked = True
        ws.Cells.FormulaHidden = True
    Next ws

    Dim k As Variant
    For Each k In plan.Keys
        Dim tabN As String, tblN As String
        tabN = Split(CStr(k), "|")(0)
        tblN = Split(CStr(k), "|")(1)

        If WorksheetExists(wb, tabN) Then
            Set ws = wb.Worksheets(tabN)
            If ListObjectExists(ws, tblN) Then
                Set lo = ws.ListObjects(tblN)
                UnlockTableColumns lo, plan(CStr(k))
            End If
        End If
    Next k

    For Each ws In wb.Worksheets
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Next ws
End Sub

Private Sub UnlockTableColumns(ByVal lo As ListObject, ByVal cols As Collection)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To cols.Count
        Dim colH As String
        colH = CStr(cols(i))
        If ColumnExists(lo, colH) Then
            lo.ListColumns(colH).DataBodyRange.Locked = False
            lo.ListColumns(colH).DataBodyRange.FormulaHidden = False
        End If
    Next i
End Sub

Private Sub ApplyWorkbookStructure(ByVal wb As Workbook, ByVal pwd As String, ByVal protectOn As Boolean)
    If protectOn Then
        wb.Protect Password:=pwd, Structure:=True, Windows:=False
    Else
        wb.Unprotect Password:=pwd
    End If
End Sub

Private Sub HideDevSheets(ByVal wb As Workbook, ByVal pwd As String)
    ' Ensures we never attempt to hide the last visible sheet,
    ' and temporarily unprotects workbook structure if needed.

    Dim names As Variant
    names = Array("SCHEMA", "AUTO", "Helpers", "Log", "Data_Check", "Schema_Check", _
                  "Lockdown_Preview", "Lockdown_Diag", "Dev_ModuleInventory", "Core_Tests", _
                  "Workbook_Schema", "SupData", "SCRIPTSPlan")

    Dim mustRemainVisible As String
    mustRemainVisible = "Landing" ' adjust if your user-facing landing sheet has a different name

    Dim ws As Worksheet
    Dim i As Long

    ' Ensure at least one sheet is visible (Landing preferred)
    If WorksheetExists(wb, mustRemainVisible) Then
        wb.Worksheets(mustRemainVisible).Visible = xlSheetVisible
    End If

    ' If workbook structure is protected, unprotect temporarily to allow Visible changes
    Dim wasStructureProtected As Boolean
    wasStructureProtected = wb.ProtectStructure
    If wasStructureProtected Then wb.Unprotect Password:=pwd

    For i = LBound(names) To UBound(names)
        If WorksheetExists(wb, CStr(names(i))) Then
            Set ws = wb.Worksheets(CStr(names(i)))

            ' Don't hide the must-remain-visible sheet even if listed accidentally
            If StrComp(ws.Name, mustRemainVisible, vbTextCompare) <> 0 Then

                ' Never hide the last visible sheet
                If CountVisibleWorksheets(wb) <= 1 Then
                    ' stop hiding to avoid Excel error 1004
                    Exit For
                End If

                ws.Unprotect Password:=pwd
                ws.Visible = xlSheetVeryHidden
                ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
            End If
        End If
    Next i

    ' Restore structure protection if it was on
    If wasStructureProtected Then wb.Protect Password:=pwd, Structure:=True, Windows:=False
End Sub

Private Function CountVisibleWorksheets(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim c As Long
    c = 0
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then c = c + 1
    Next ws
    CountVisibleWorksheets = c
End Function


'----------------------------
' Preview report
'----------------------------

Private Sub WritePreview(ByVal wb As Workbook, ByVal plan As Object, ByVal currentRole As String)
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(wb, "Lockdown_Preview")

    ws.Cells.ClearContents
    ws.Range("A1:F1").value = Array("TabName", "TableName", "UnlockedColumn", "FoundInWorkbook", "CurrentRole", "Notes")
    ws.Range("H1").value = "CurrentRole"
    ws.Range("H2").value = currentRole

    Dim outRow As Long
    outRow = 2

    Dim k As Variant
    For Each k In plan.Keys
        Dim tabN As String, tblN As String
        tabN = Split(CStr(k), "|")(0)
        tblN = Split(CStr(k), "|")(1)

        Dim existsSheet As Boolean, existsTable As Boolean
        existsSheet = WorksheetExists(wb, tabN)
        existsTable = False
        If existsSheet Then existsTable = ListObjectExists(wb.Worksheets(tabN), tblN)

        Dim cols As Collection
        Set cols = plan(CStr(k))

        Dim i As Long
        For i = 1 To cols.Count
            ws.Cells(outRow, 1).value = tabN
            ws.Cells(outRow, 2).value = tblN
            ws.Cells(outRow, 3).value = CStr(cols(i))
            ws.Cells(outRow, 4).value = IIf(existsSheet And existsTable, "Y", "")
            ws.Cells(outRow, 5).value = currentRole
            ws.Cells(outRow, 6).value = ""
            outRow = outRow + 1
        Next i
    Next k

    ws.Columns("A:F").AutoFit
End Sub

'----------------------------
' Utilities
'----------------------------

Private Function GetPassword(ByVal defaultPwd As String) As String
    If UCase$(defaultPwd) <> "CHANGE_ME" Then
        GetPassword = defaultPwd
    Else
        GetPassword = InputBox("Enter developer protection password (used for Protect/Unprotect):", "Lockdown Password")
    End If
End Function

Private Sub RequireHeader(ByVal lo As ListObject, ByVal headerName As String)
    If Not HasHeader(lo, headerName) Then Err.Raise vbObjectError + 220, , "Missing required schema header: " & headerName
End Sub

Private Function HasHeader(ByVal lo As ListObject, ByVal headerName As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(headerName)
    On Error GoTo 0
    HasHeader = Not lc Is Nothing
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

Private Function ListObjectExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    ListObjectExists = Not lo Is Nothing
End Function

Private Function ColumnExists(ByVal lo As ListObject, ByVal headerName As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(headerName)
    On Error GoTo 0
    ColumnExists = Not lc Is Nothing
End Function

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set EnsureWorksheet = ws
End Function

Private Function IsTrueishText(ByVal s As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(CStr(s)))
    IsTrueishText = (u = "Y" Or u = "YES" Or u = "TRUE" Or u = "1")
End Function

Private Function CollectionContains(ByVal col As Collection, ByVal value As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(CStr(col(i)), value, vbTextCompare) = 0 Then
            CollectionContains = True
            Exit Function
        End If
    Next i
    CollectionContains = False
End Function


