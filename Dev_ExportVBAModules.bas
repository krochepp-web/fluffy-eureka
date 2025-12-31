Attribute VB_Name = "Dev_ExportVBAModules"
Option Explicit

'===========================================================
' Purpose:
'   Export all VBA modules (Std/Class/UserForms) to a user-
'   selected folder using a Save As dialog (robust when the
'   workbook path is an Office/OneDrive URL such as
'   https://d.docs.live.net/...).
'
' Inputs:
'   - ThisWorkbook.VBProject.VBComponents
'
' Outputs / Side effects:
'   - Exports .bas / .cls / .frm (+ .frx)
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
' Version: v1.0.0
' Author:  ChatGPT (CUSTOM Tracker)
' Date:    2025-12-19
'===========================================================

Public Sub Export_All_VBAModules_SaveAsFolder()
    Const PROC_NAME As String = "Export_All_VBAModules_SaveAsFolder"

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

    ' Use Save As to select a REAL filesystem destination.
    ' We ask for a file name, then export to its folder.
    fPath = Application.GetSaveAsFilename( _
                InitialFileName:=GetDefaultExportStub(), _
                FileFilter:="Text Files (*.txt), *.txt", _
                TITLE:="Pick export folder (file name is ignored; folder is used)")

    If VarType(fPath) = vbBoolean And fPath = False Then Exit Sub ' user cancelled

    exportFolder = GetFolderFromPath(CStr(fPath))
    If Len(exportFolder) = 0 Then
        MsgBox "Could not resolve export folder from selected path.", vbCritical, "Export Failed"
        Exit Sub
    End If

    ' Ensure folder exists (should, but guard anyway)
    If Not EnsureFolderExistsSafe(exportFolder) Then
        MsgBox "Cannot create or access export folder: " & exportFolder, vbCritical, "Export Failed"
        Exit Sub
    End If

    Debug.Print "=== " & PROC_NAME & " ==="
    Debug.Print "Workbook: " & ThisWorkbook.Name
    Debug.Print "ExportFolder: " & exportFolder

    okCount = 0: failCount = 0: hadErr52 = False: failReport = vbNullString
    ExportComponentsToFolder exportFolder, okCount, failCount, hadErr52, failReport

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
    EnsureFolderExistsSafe = False
End Function

'========================
' Component + filename utilities
'========================

Private Function ComponentExtension(ByVal compType As Long) As String
    Select Case compType
        Case 1: ComponentExtension = ".bas" ' Std module
        Case 2: ComponentExtension = ".cls" ' Class module
        Case 3: ComponentExtension = ".frm" ' UserForm
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


