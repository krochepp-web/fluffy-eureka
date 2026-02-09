Attribute VB_Name = "Dev_ImportExportVBAModules1"
Option Explicit

'===========================================================
' Purpose:
'   Export all VBA modules (Std/Class/UserForms) to a user-
'   selected folder using a Save As dialog (robust when the
'   workbook path is an Office/OneDrive URL such as
'   https://d.docs.live.net/...), plus import standard modules
'   from .bas files in a selected folder.
'
' Inputs:
'   - ThisWorkbook.VBProject.VBComponents
'
' Outputs / Side effects:
'   - Exports .bas / .cls / .frm (+ .frx)
'   - Imports .bas (standard modules only)
'   - Creates _EXPORT_INFO.txt in the destination folder
'
' Preconditions:
'   - Trust access to the VBA project model enabled
'
' Postconditions:
'   - Destination folder contains exported source files
'
' Errors & Guards:
'   - Continues on per-component export failures
'   - Reports failures in Immediate Window (Ctrl+G)
'
' Version: v1.1.0
' Author:  ChatGPT (CUSTOM Tracker)
' Date:    2025-12-19
'===========================================================

Private Const VB_COMP_STD_MODULE As Long = 1
Private Const VB_COMP_CLASS_MODULE As Long = 2
Private Const VB_COMP_USERFORM As Long = 3
Private Const DIALOG_FOLDER_PICKER As Long = 4

'========================
' Public entry points
'========================

Public Sub Export_All_VBAModules_SaveAsFolder()
    Export_All_VBAModules False
End Sub

Public Sub Export_All_VBAModules_ToTxt()
    Export_All_VBAModules True
End Sub

Public Sub Export_All_VBAModules(Optional ByVal exportAsText As Boolean = False)
    Const PROC_NAME As String = "Export_All_VBAModules"

    Dim fPath As Variant
    Dim exportFolder As String
    Dim okCount As Long, failCount As Long
    Dim failReport As String
    Dim hadErr52 As Boolean

    On Error GoTo EH

    If Not CanAccessVBAProject() Then
        MsgBox _
            "VBA export blocked." & vbCrLf & _
            "Enable: File ? Options ? Trust Center ? Trust Center Settings ?" & vbCrLf & _
            "Macro Settings ? 'Trust access to the VBA project model'.", _
            vbCritical, "Export Failed"
        Exit Sub
    End If

    If exportAsText Then
        exportFolder = Get_ExportFolder()
        If exportFolder = vbNullString Then
            MsgBox "Export cancelled by user.", vbInformation
            Exit Sub
        End If
    Else
        ' Use Save As to select a REAL filesystem destination.
        ' We ask for a file name, then export to its folder.
        fPath = Application.GetSaveAsFilename( _
                    InitialFileName:=GetDefaultExportStub(), _
                    FileFilter:="Text Files (*.txt), *.txt", _
                    title:="Pick export folder (file name is ignored; folder is used)")

        If VarType(fPath) = vbBoolean And fPath = False Then Exit Sub ' user cancelled

        exportFolder = GetFolderFromPath(CStr(fPath))
        If Len(exportFolder) = 0 Then
            MsgBox "Could not resolve export folder from selected path.", vbCritical, "Export Failed"
            Exit Sub
        End If
    End If

    ' Ensure folder exists (should, but guard anyway)
    If Not EnsureFolderExistsSafe(exportFolder) Then
        MsgBox "Cannot create or access export folder: " & exportFolder, vbCritical, "Export Failed"
        Exit Sub
    End If

    LogHeader PROC_NAME, exportFolder

    okCount = 0: failCount = 0: hadErr52 = False: failReport = vbNullString
    If exportAsText Then
        ExportComponentsToFolderTxt exportFolder, okCount, failCount, failReport
    Else
        ExportComponentsToFolder exportFolder, okCount, failCount, hadErr52, failReport
    End If

    WriteExportInfo exportFolder, okCount, failCount, hadErr52

    If failCount > 0 Then Debug.Print failReport

    MsgBox _
        "VBA export complete." & vbCrLf & _
        "Folder: " & exportFolder & vbCrLf & _
        "OK: " & okCount & "   Failed: " & failCount & _
        IIf(hadErr52, vbCrLf & vbCrLf & "Note: Err 52 occurred for at least one component (see Immediate Window).", "") & _
        IIf(failCount > 0, vbCrLf & vbCrLf & "See Immediate Window (Ctrl+G) for details.", ""), _
        IIf(failCount > 0 Or hadErr52, vbExclamation, vbInformation), _
        "Export Results"

    Exit Sub

EH:
    MsgBox _
        "Error in " & PROC_NAME & vbCrLf & _
        "Err " & Err.Number & ": " & Err.Description, _
        vbCritical, "Export Failed"
End Sub

Public Sub Import_All_BAS_Modules_FromFolder()
    Const PROC_NAME As String = "Import_All_BAS_Modules_FromFolder"

    Dim importFolder As String
    Dim fileName As String
    Dim filePath As String
    Dim moduleName As String
    Dim vbComp As Object
    Dim okCount As Long, failCount As Long
    Dim foundCount As Long

    On Error GoTo EH

    If Not CanAccessVBAProject() Then
        MsgBox _
            "VBA import blocked." & vbCrLf & _
            "Enable: File ? Options ? Trust Center ? Trust Center Settings ?" & vbCrLf & _
            "Macro Settings ? 'Trust access to the VBA project model'.", _
            vbCritical, "Import Failed"
        Exit Sub
    End If

    importFolder = PickImportFolder()
    If Len(importFolder) = 0 Then Exit Sub

    LogHeader PROC_NAME, importFolder

    okCount = 0: failCount = 0: foundCount = 0

    fileName = Dir(importFolder & "\*.bas")
    Do While Len(fileName) > 0
        foundCount = foundCount + 1
        filePath = importFolder & "\" & fileName
        moduleName = Left$(fileName, Len(fileName) - 4)

        On Error GoTo ImportFail

        RemoveStandardModuleIfExists moduleName, vbComp

        ThisWorkbook.VBProject.VBComponents.Import filePath
        Debug.Print "OK: " & moduleName & " <- " & filePath
        okCount = okCount + 1
        On Error GoTo 0

ContinueNext:
        fileName = Dir()
    Loop

    If foundCount = 0 Then
        MsgBox _
            "No .bas files found in the selected folder." & vbCrLf & _
            "Folder: " & importFolder, _
            vbInformation, "Import Results"
        Exit Sub
    End If

    MsgBox _
        "VBA import complete." & vbCrLf & _
        "Folder: " & importFolder & vbCrLf & _
        "OK: " & okCount & "   Failed: " & failCount & _
        IIf(failCount > 0, vbCrLf & vbCrLf & "See Immediate Window (Ctrl+G) for details.", ""), _
        IIf(failCount > 0, vbExclamation, vbInformation), _
        "Import Results"

    Exit Sub

ImportFail:
    failCount = failCount + 1

    Debug.Print "FAIL: " & moduleName & " -> " & filePath
    Debug.Print "      Err " & Err.Number & ": " & Err.Description

    Err.Clear
    On Error GoTo 0
    GoTo ContinueNext

EH:
    MsgBox _
        "Error in " & PROC_NAME & vbCrLf & _
        "Err " & Err.Number & ": " & Err.Description, _
        vbCritical, "Import Failed"
End Sub

'========================
' Export implementation
'========================

Private Sub ExportComponentsToFolder(ByVal exportFolder As String, _
                                    ByRef okCount As Long, _
                                    ByRef failCount As Long, _
                                    ByRef hadErr52 As Boolean, _
                                    ByRef failReport As String)
    Dim vbComp As Object
    Dim outFile As String, ext As String, compNameSafe As String

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ext = ComponentExtension(vbComp.Type)
        If Len(ext) = 0 Then
            ' Skip document modules (ThisWorkbook / Sheet modules)
        Else
            compNameSafe = SafeFileName(CStr(vbComp.Name))
            outFile = exportFolder & "\" & compNameSafe & ext

            On Error GoTo ExportFail
            vbComp.Export outFile
            okCount = okCount + 1
            On Error GoTo 0
        End If

ContinueNext:
    Next vbComp
    Exit Sub

ExportFail:
    failCount = failCount + 1
    If Err.Number = 52 Then hadErr52 = True

    failReport = failReport & vbCrLf & _
        "- " & vbComp.Name & " -> " & outFile & " | Err " & Err.Number & ": " & Err.Description

    Debug.Print "FAIL: " & vbComp.Name & " -> " & outFile
    Debug.Print "      Err " & Err.Number & ": " & Err.Description

    Err.Clear
    On Error GoTo 0
    GoTo ContinueNext
End Sub

Private Sub ExportComponentsToFolderTxt(ByVal exportFolder As String, _
                                        ByRef okCount As Long, _
                                        ByRef failCount As Long, _
                                        ByRef failReport As String)
    Dim vbComp As Object
    Dim codeMod As Object
    Dim outFile As String
    Dim fileNum As Long
    Dim codeText As String
    Dim lineCount As Long
    Dim compNameSafe As String

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        compNameSafe = SafeFileName(CStr(Get_ComponentFileName(vbComp)))
        outFile = exportFolder & "\" & compNameSafe & ".txt"

        On Error GoTo ExportFail
        Set codeMod = vbComp.CodeModule
        lineCount = codeMod.CountOfLines

        If lineCount > 0 Then
            codeText = codeMod.lines(1, lineCount)
        Else
            codeText = vbNullString
        End If

        fileNum = FreeFile
        Open outFile For Output As #fileNum
        Print #fileNum, codeText
        Close #fileNum
        okCount = okCount + 1
        On Error GoTo 0

ContinueNext:
    Next vbComp
    Exit Sub

ExportFail:
    failCount = failCount + 1
    failReport = failReport & vbCrLf & _
        "- " & vbComp.Name & " -> " & outFile & " | Err " & Err.Number & ": " & Err.Description

    Debug.Print "FAIL: " & vbComp.Name & " -> " & outFile
    Debug.Print "      Err " & Err.Number & ": " & Err.Description

    Err.Clear
    On Error GoTo 0
    GoTo ContinueNext
End Sub

'========================
' Provenance file
'========================

Private Sub WriteExportInfo(ByVal exportFolder As String, ByVal okCount As Long, ByVal failCount As Long, ByVal hadErr52 As Boolean)
    Dim fso As Object, ts As Object
    Dim infoFile As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    infoFile = exportFolder & "\_EXPORT_INFO.txt"
    Set ts = fso.CreateTextFile(infoFile, True)

    ts.WriteLine "CUSTOM Tracker – VBA Export"
    ts.WriteLine String(55, "-")
    ts.WriteLine "WorkbookName: " & ThisWorkbook.Name
    ts.WriteLine "WorkbookPath: " & ThisWorkbook.FullName
    ts.WriteLine "ExportFolder: " & exportFolder
    ts.WriteLine "ExportedAt: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ts.WriteLine "User: " & Environ$("USERNAME")
    ts.WriteLine "ExcelVersion: " & Application.Version
    ts.WriteLine "OK: " & okCount
    ts.WriteLine "Failed: " & failCount
    ts.WriteLine "HadErr52: " & IIf(hadErr52, "YES", "NO")
    ts.WriteLine "Import: Run Import_All_BAS_Modules_FromFolder to import .bas modules from this folder."
    ts.Close
End Sub

'========================
' Path helpers
'========================

Private Function GetFolderFromPath(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetFolderFromPath = Left$(fullPath, p - 1)
    Else
        GetFolderFromPath = vbNullString
    End If
End Function

Private Function GetDefaultExportStub() As String
    ' This is just a stub to drive the dialog into a sensible place.
    ' File name is ignored; only folder is used.
    Dim wbBase As String
    wbBase = SafeFileName(GetWorkbookBaseName(ThisWorkbook.Name))
    GetDefaultExportStub = wbBase & "_VBA_EXPORT_INFO.txt"
End Function

Private Function PickImportFolder() As String
    Dim dlg As Object

    Set dlg = Application.FileDialog(DIALOG_FOLDER_PICKER)
    With dlg
        .title = "Select folder containing .bas modules to import"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickImportFolder = vbNullString
            Exit Function
        End If
        PickImportFolder = .SelectedItems(1)
    End With
End Function

Private Function GetWorkbookBaseName(ByVal wbName As String) As String
    Dim p As Long
    p = InStrRev(wbName, ".")
    If p > 0 Then
        GetWorkbookBaseName = Left$(wbName, p - 1)
    Else
        GetWorkbookBaseName = wbName
    End If
End Function

Private Function EnsureFolderExistsSafe(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    EnsureFolderExistsSafe = True
    Exit Function
EH:
    MsgBox "EnsureFolderExistsSafe failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Export Failed"
    EnsureFolderExistsSafe = False
End Function

'========================
' Component + filename utilities
'========================

Private Function Get_ExportFolder() As String
    Dim fDialog As FileDialog

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

    With fDialog
        .title = "Select folder to export VBA modules"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Get_ExportFolder = vbNullString
        Else
            Get_ExportFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Function Get_ComponentFileName(ByVal vbComp As Object) As String
    ' Prefix document modules so they are obvious in source control
    Select Case CLng(vbComp.Type)
        Case 100 ' vbext_ct_Document
            Get_ComponentFileName = "Doc_" & SafeFileName(CStr(vbComp.Name))
        Case Else
            Get_ComponentFileName = SafeFileName(CStr(vbComp.Name))
    End Select
End Function

Private Function ComponentExtension(ByVal compType As Long) As String
    Select Case compType
        Case VB_COMP_STD_MODULE: ComponentExtension = ".bas" ' Std module
        Case VB_COMP_CLASS_MODULE: ComponentExtension = ".cls" ' Class module
        Case VB_COMP_USERFORM: ComponentExtension = ".frm" ' UserForm
        Case Else: ComponentExtension = vbNullString ' Document modules
    End Select
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    s = Trim$(s)
    If Len(s) = 0 Then s = "Unnamed"
    SafeFileName = s
End Function

Private Function CanAccessVBAProject() As Boolean
    On Error Resume Next
    Dim test As Object
    Set test = ThisWorkbook.VBProject.VBComponents
    CanAccessVBAProject = (Err.Number = 0)
    Err.Clear
End Function

Private Sub LogHeader(ByVal procName As String, ByVal targetFolder As String)
    Debug.Print "=== " & procName & " ==="
    Debug.Print "Workbook: " & ThisWorkbook.Name
    Debug.Print "Folder: " & targetFolder
End Sub

Private Sub RemoveStandardModuleIfExists(ByVal moduleName As String, ByRef vbComp As Object)
    If StandardModuleExists(moduleName, vbComp) Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
        Debug.Print "REMOVED: " & moduleName & " (existing standard module)"
    End If
End Sub

Private Function StandardModuleExists(ByVal moduleName As String, ByRef vbComp As Object) As Boolean
    Dim comp As Object

    Set vbComp = Nothing
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = VB_COMP_STD_MODULE Then
            If StrComp(comp.Name, moduleName, vbTextCompare) = 0 Then
                Set vbComp = comp
                StandardModuleExists = True
                Exit Function
            End If
        End If
    Next comp
End Function
