Attribute VB_Name = "Dev_Export_ALL_ToTxt"
Option Explicit

'***************************************************************************
' Module:       Dev_Export_ALL_ToTxt
' Procedure:    Dev_Export_AllVBAModules_ToTxt
'
' Purpose:
'   Export the source code of all VBA components (standard modules, class
'   modules, userforms, and document modules) to .txt files in a user-
'   selected directory.
'
' Inputs:
'   - ActiveWorkbook.VBProject.VBComponents
'
' Outputs / Side Effects:
'   - Creates or overwrites *.txt files in the selected folder
'
' Preconditions:
'   - Workbook must be macro-enabled (.xlsm)
'   - Trust Center: "Trust access to the VBA project object model" enabled
'
' Postconditions:
'   - One .txt file exists per VBA component
'
' Errors & Guards:
'   - Fails fast if VBProject access is denied
'   - Safe cancel if user aborts folder picker
'
' Version:      v1.0.0
' Author:       Keith + ChatGPT
' Date:         2026-02-04
'***************************************************************************

Public Sub Dev_Export_AllVBAModules_ToTxt()
    Dim vbComp As Object
    Dim codeMod As Object
    Dim exportPath As String
    Dim outFile As String
    Dim fileNum As Long
    Dim codeText As String
    Dim lineCount As Long

    On Error GoTo EH

    exportPath = Get_ExportFolder()
    If exportPath = vbNullString Then
        MsgBox "Export cancelled by user.", vbInformation
        GoTo CleanExit
    End If

    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        Set codeMod = vbComp.CodeModule
        lineCount = codeMod.CountOfLines

        If lineCount > 0 Then
            codeText = codeMod.lines(1, lineCount)
        Else
            codeText = vbNullString
        End If

        outFile = exportPath & Application.PathSeparator & _
                  Get_ComponentFileName(vbComp) & ".txt"

        fileNum = FreeFile
        Open outFile For Output As #fileNum
        Print #fileNum, codeText
        Close #fileNum
    Next vbComp

    MsgBox "VBA modules exported as .txt files to:" & vbCrLf & exportPath, vbInformation

CleanExit:
    Exit Sub

EH:
    MsgBox _
        "Error exporting VBA modules." & vbCrLf & _
        "Err " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
        "Check Trust Center settings for VBProject access.", _
        vbCritical
    Resume CleanExit
End Sub

'========================
' Helpers
'========================

Private Function Get_ExportFolder() As String
    Dim fDialog As FileDialog

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

    With fDialog
        .title = "Select folder to export VBA modules as .txt"
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
            Get_ComponentFileName = "Doc_" & Sanitize_FileName(vbComp.Name)
        Case Else
            Get_ComponentFileName = Sanitize_FileName(vbComp.Name)
    End Select
End Function

Private Function Sanitize_FileName(ByVal rawName As String) As String
    Dim s As String
    s = rawName

    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")

    Sanitize_FileName = s
End Function


