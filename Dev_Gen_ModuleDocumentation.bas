Attribute VB_Name = "Dev_Gen_ModuleDocumentation"
Option Explicit

'===========================================================
' Purpose:
'   Generate an auditable catalog of VBA modules and procedures,
'   including best-effort extraction of standardized header doc
'   blocks (Purpose/Inputs/Outputs/Preconditions/Errors/Version/etc.).
'
' Inputs:
'   - VBAProject source code (modules, class/document modules)
'
' Outputs / Side effects:
'   - Creates/overwrites sheets:
'       Dev_ModuleCatalog
'       Dev_ProcedureCatalog
'   - Populates them with:
'       * module inventory
'       * procedure inventory
'       * extracted doc-block fields (when present)
'
' Preconditions:
'   - Excel option enabled:
'       File > Options > Trust Center > Trust Center Settings >
'       Macro Settings > "Trust access to the VBA project object model"
'   - Reference recommended (but late-binding is used to reduce friction):
'       Microsoft Visual Basic for Applications Extensibility 5.3
'
' Postconditions:
'   - Two sheets exist with refreshed catalogs
'
' Errors & Guards:
'   - Fails fast with a clear message if VBA project access is blocked
'
' Version: v1.0.0
' Author: ChatGPT
' Date: 2025-12-21
'===========================================================

Public Sub Dev_Generate_ModuleDocumentation()
    Const PROC_NAME As String = "Dev_Generate_ModuleDocumentation"

    Dim wb As Workbook
    Dim vbProj As Object ' VBIDE.VBProject
    Dim vbComp As Object ' VBIDE.VBComponent

    Dim wsMod As Worksheet, wsProc As Worksheet
    Dim modRow As Long, procRow As Long

    Dim calcMode As XlCalculation
    Dim scrUpdate As Boolean, evts As Boolean

    On Error GoTo EH

    Set wb = ThisWorkbook

    '--- toggles
    scrUpdate = Application.ScreenUpdating
    evts = Application.EnableEvents
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- access VBA project (will error if trust setting is off)
    Set vbProj = wb.VBProject

    '--- prep output sheets
    Set wsMod = EnsureSheet(wb, "Dev_ModuleCatalog")
    Set wsProc = EnsureSheet(wb, "Dev_ProcedureCatalog")

    ClearSheet wsMod
    ClearSheet wsProc

    WriteModuleHeaders wsMod
    WriteProcHeaders wsProc

    modRow = 2
    procRow = 2

    For Each vbComp In vbProj.VBComponents
        WriteModuleRow wsMod, modRow, vbComp
        modRow = modRow + 1

        WriteProceduresForComponent wsProc, procRow, vbComp
    Next vbComp

    AutofitAll wsMod
    AutofitAll wsProc

CleanExit:
    Application.ScreenUpdating = scrUpdate
    Application.EnableEvents = evts
    Application.Calculation = calcMode
    Exit Sub

EH:
    MsgBox "Module documentation generation failed." & vbCrLf & _
           "Most common cause: Trust Center setting 'Trust access to the VBA project object model' is OFF." & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, PROC_NAME
    Resume CleanExit
End Sub

'-------------------------
' Module catalog writers
'-------------------------

Private Sub WriteModuleHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).value = "ModuleName"
    ws.Cells(1, 2).value = "ComponentType"
    ws.Cells(1, 3).value = "CodeLines"
    ws.Cells(1, 4).value = "HasCode"
    ws.Cells(1, 5).value = "Notes"
    ws.rows(1).Font.Bold = True
End Sub

Private Sub WriteModuleRow(ByVal ws As Worksheet, ByVal r As Long, ByVal vbComp As Object)
    Dim codeLines As Long
    Dim hasCode As String

    codeLines = 0
    On Error Resume Next
    codeLines = vbComp.CodeModule.CountOfLines
    On Error GoTo 0

    hasCode = IIf(codeLines > 0, "Y", "N")

    ws.Cells(r, 1).value = vbComp.Name
    ws.Cells(r, 2).value = ComponentTypeName(vbComp.Type)
    ws.Cells(r, 3).value = codeLines
    ws.Cells(r, 4).value = hasCode
    ws.Cells(r, 5).value = ""
End Sub

Private Function ComponentTypeName(ByVal compType As Long) As String
    ' VBIDE.vbext_ComponentType values:
    ' 1=StdModule, 2=ClassModule, 3=MSForm, 100=Document
    Select Case compType
        Case 1: ComponentTypeName = "StdModule"
        Case 2: ComponentTypeName = "ClassModule"
        Case 3: ComponentTypeName = "MSForm"
        Case 100: ComponentTypeName = "Document"
        Case Else: ComponentTypeName = "Unknown(" & CStr(compType) & ")"
    End Select
End Function

'-------------------------
' Procedure catalog writers
'-------------------------

Private Sub WriteProcHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).value = "ModuleName"
    ws.Cells(1, 2).value = "ComponentType"
    ws.Cells(1, 3).value = "ProcName"
    ws.Cells(1, 4).value = "ProcKind"
    ws.Cells(1, 5).value = "Scope"
    ws.Cells(1, 6).value = "SignatureLine"
    ws.Cells(1, 7).value = "StartLine"
    ws.Cells(1, 8).value = "Doc_Purpose"
    ws.Cells(1, 9).value = "Doc_Inputs"
    ws.Cells(1, 10).value = "Doc_Outputs"
    ws.Cells(1, 11).value = "Doc_Preconditions"
    ws.Cells(1, 12).value = "Doc_ErrorsGuards"
    ws.Cells(1, 13).value = "Doc_Trigger"
    ws.Cells(1, 14).value = "Doc_EntryPoint"
    ws.Cells(1, 15).value = "Doc_Version"
    ws.Cells(1, 16).value = "Doc_Author"
    ws.Cells(1, 17).value = "Doc_Date"
    ws.Cells(1, 18).value = "Doc_RawBlock"
    ws.rows(1).Font.Bold = True
End Sub

Private Sub WriteProceduresForComponent(ByVal ws As Worksheet, ByRef procRow As Long, ByVal vbComp As Object)
    Dim cm As Object ' CodeModule
    Dim lineNum As Long
    Dim procName As String
    Dim procKind As Long

    Dim startLine As Long
    Dim scopeName As String
    Dim kindName As String
    Dim sigLine As String

    Dim doc As Object ' Scripting.Dictionary (late-bound)
    Dim rawBlock As String

    On Error GoTo SafeExit

    Set cm = vbComp.CodeModule
    If cm.CountOfLines = 0 Then Exit Sub

    lineNum = 1
    Do While lineNum <= cm.CountOfLines
        procName = cm.ProcOfLine(lineNum, procKind)
        If Len(procName) > 0 Then
            startLine = cm.ProcStartLine(procName, procKind)

            sigLine = GetProcedureSignatureLine(cm, startLine)
            scopeName = GetProcedureScope(sigLine)
            kindName = ProcKindName(sigLine)

            rawBlock = GetDocBlockAbove(cm, startLine)
            Set doc = ParseDocBlock(rawBlock)

            ws.Cells(procRow, 1).value = vbComp.Name
            ws.Cells(procRow, 2).value = ComponentTypeName(vbComp.Type)
            ws.Cells(procRow, 3).value = procName
            ws.Cells(procRow, 4).value = kindName
            ws.Cells(procRow, 5).value = scopeName
            ws.Cells(procRow, 6).value = sigLine
            ws.Cells(procRow, 7).value = startLine

            ws.Cells(procRow, 8).value = NzDict(doc, "Purpose")
            ws.Cells(procRow, 9).value = NzDict(doc, "Inputs")
            ws.Cells(procRow, 10).value = NzDict(doc, "Outputs")
            ws.Cells(procRow, 11).value = NzDict(doc, "Preconditions")
            ws.Cells(procRow, 12).value = NzDict(doc, "Errors")
            ws.Cells(procRow, 13).value = NzDict(doc, "Trigger")
            ws.Cells(procRow, 14).value = NzDict(doc, "EntryPoint")
            ws.Cells(procRow, 15).value = NzDict(doc, "Version")
            ws.Cells(procRow, 16).value = NzDict(doc, "Author")
            ws.Cells(procRow, 17).value = NzDict(doc, "Date")
            ws.Cells(procRow, 18).value = rawBlock

            procRow = procRow + 1

            lineNum = startLine + cm.ProcCountLines(procName, procKind)
        Else
            lineNum = lineNum + 1
        End If
    Loop

SafeExit:
    Exit Sub
End Sub

Private Function GetProcedureSignatureLine(ByVal cm As Object, ByVal startLine As Long) As String
    ' Best-effort: return the first non-empty line at proc start.
    Dim s As String
    Dim i As Long

    For i = 0 To 3
        s = Trim$(cm.lines(startLine + i, 1))
        If Len(s) > 0 Then
            GetProcedureSignatureLine = s
            Exit Function
        End If
    Next i

    GetProcedureSignatureLine = Trim$(cm.lines(startLine, 1))
End Function

Private Function GetProcedureScope(ByVal sigLine As String) As String
    Dim t As String
    t = LCase$(sigLine)
    If InStr(1, t, "public ", vbTextCompare) > 0 Then
        GetProcedureScope = "Public"
    ElseIf InStr(1, t, "private ", vbTextCompare) > 0 Then
        GetProcedureScope = "Private"
    ElseIf InStr(1, t, "friend ", vbTextCompare) > 0 Then
        GetProcedureScope = "Friend"
    Else
        GetProcedureScope = ""
    End If
End Function

Private Function ProcKindName(ByVal sigLine As String) As String
    Dim t As String
    t = LCase$(sigLine)

    If InStr(1, t, "function ", vbTextCompare) > 0 Then
        ProcKindName = "Function"
    ElseIf InStr(1, t, "sub ", vbTextCompare) > 0 Then
        ProcKindName = "Sub"
    ElseIf InStr(1, t, "property get", vbTextCompare) > 0 Then
        ProcKindName = "Property Get"
    ElseIf InStr(1, t, "property let", vbTextCompare) > 0 Then
        ProcKindName = "Property Let"
    ElseIf InStr(1, t, "property set", vbTextCompare) > 0 Then
        ProcKindName = "Property Set"
    Else
        ProcKindName = "Unknown"
    End If
End Function

Private Function GetDocBlockAbove(ByVal cm As Object, ByVal startLine As Long) As String
    ' Collect contiguous comment lines immediately above the procedure.
    ' Stops at first blank line or first non-comment line.
    Dim i As Long
    Dim s As String
    Dim Block As String

    Block = ""
    For i = startLine - 1 To 1 Step -1
        s = cm.lines(i, 1)

        If Len(Trim$(s)) = 0 Then
            Exit For
        End If

        If Left$(Trim$(s), 1) = "'" Then
            Block = s & vbCrLf & Block
        Else
            Exit For
        End If
    Next i

    GetDocBlockAbove = Trim$(Block)
End Function

Private Function ParseDocBlock(ByVal rawBlock As String) As Object
    ' Returns a late-bound Scripting.Dictionary with keys:
    ' Purpose, Inputs, Outputs, Preconditions, Errors, Trigger, EntryPoint, Version, Author, Date
    Dim d As Object
    Dim lines() As String
    Dim i As Long, s As String
    Dim key As String, val As String

    Set d = CreateObject("Scripting.Dictionary")
    d.compareMode = 1 ' vbTextCompare

    If Len(rawBlock) = 0 Then
        Set ParseDocBlock = d
        Exit Function
    End If

    lines = Split(rawBlock, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If Left$(s, 1) = "'" Then s = Trim$(Mid$(s, 2))

        key = ""
        val = ""

        If StartsWithKey(s, "Purpose:") Then
            key = "Purpose": val = Trim$(Mid$(s, Len("Purpose:") + 1))
        ElseIf StartsWithKey(s, "Inputs:") Then
            key = "Inputs": val = Trim$(Mid$(s, Len("Inputs:") + 1))
        ElseIf StartsWithKey(s, "Outputs") Then
            key = "Outputs": val = Trim$(AfterColon(s))
        ElseIf StartsWithKey(s, "Preconditions:") Then
            key = "Preconditions": val = Trim$(Mid$(s, Len("Preconditions:") + 1))
        ElseIf StartsWithKey(s, "Errors") Then
            key = "Errors": val = Trim$(AfterColon(s))
        ElseIf StartsWithKey(s, "Trigger") Then
            key = "Trigger": val = Trim$(AfterColon(s))
        ElseIf StartsWithKey(s, "Entry") Then
            key = "EntryPoint": val = Trim$(AfterColon(s))
        ElseIf StartsWithKey(s, "Version:") Then
            key = "Version": val = Trim$(Mid$(s, Len("Version:") + 1))
        ElseIf StartsWithKey(s, "Author:") Then
            key = "Author": val = Trim$(Mid$(s, Len("Author:") + 1))
        ElseIf StartsWithKey(s, "Date:") Then
            key = "Date": val = Trim$(Mid$(s, Len("Date:") + 1))
        End If

        If Len(key) > 0 Then
            If d.Exists(key) Then
                d(key) = Trim$(d(key) & " " & val)
            Else
                d.Add key, val
            End If
        End If
    Next i

    Set ParseDocBlock = d
End Function

Private Function StartsWithKey(ByVal s As String, ByVal keyToken As String) As Boolean
    StartsWithKey = (LCase$(Left$(s, Len(keyToken))) = LCase$(keyToken))
End Function

Private Function AfterColon(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, ":", vbTextCompare)
    If p > 0 Then
        AfterColon = Mid$(s, p + 1)
    Else
        AfterColon = ""
    End If
End Function

Private Function NzDict(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then
        NzDict = ""
    ElseIf d.Exists(key) Then
        NzDict = CStr(d(key))
    Else
        NzDict = ""
    End If
End Function

'-------------------------
' Sheet helpers
'-------------------------

Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set EnsureSheet = ws
End Function

Private Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Sub AutofitAll(ByVal ws As Worksheet)
    ws.Columns.AutoFit
    ws.rows(1).WrapText = False
End Sub


