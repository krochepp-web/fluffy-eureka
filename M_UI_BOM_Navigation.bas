Attribute VB_Name = "M_UI_BOM_Navigation"
Option Explicit

Public Sub UI_GoTo_Selected_BOM_FromBOMS()
    Const PROC_NAME As String = "M_UI_BOM_Navigation.UI_GoTo_Selected_BOM_FromBOMS"
    Const SH_BOMS As String = "BOMS"
    Const LO_BOMS As String = "TBL_BOMS"
    Const COL_BOMTAB As String = "BOMTab"

    Dim wb As Workbook
    Dim wsBoms As Worksheet
    Dim loBoms As ListObject
    Dim selectedCell As Range
    Dim selectedRow As Range
    Dim rowOffset As Long
    Dim idxBomTab As Long
    Dim bomTabName As String
    Dim wsTarget As Worksheet
    Dim lc As ListColumn

    On Error GoTo EH

    Set wb = ThisWorkbook
    Set wsBoms = wb.Worksheets(SH_BOMS)
    Set loBoms = wsBoms.ListObjects(LO_BOMS)

    If ActiveSheet.Name <> SH_BOMS Then
        MsgBox "Please run this from the BOMS sheet after selecting a BOM row.", vbInformation, "Go To Specific BOM"
        Exit Sub
    End If

    If loBoms.DataBodyRange Is Nothing Then
        MsgBox "BOMS table has no data rows.", vbExclamation, "Go To Specific BOM"
        Exit Sub
    End If

    Set selectedCell = ActiveCell
    If selectedCell Is Nothing Then
        MsgBox "Select any cell in the BOM row you want to open.", vbInformation, "Go To Specific BOM"
        Exit Sub
    End If

    Set selectedRow = Intersect(selectedCell.EntireRow, loBoms.DataBodyRange)
    If selectedRow Is Nothing Then
        MsgBox "Select a cell inside TBL_BOMS data rows, then run again.", vbExclamation, "Go To Specific BOM"
        Exit Sub
    End If

    idxBomTab = 0
    For Each lc In loBoms.ListColumns
        If StrComp(lc.Name, COL_BOMTAB, vbTextCompare) = 0 Then
            idxBomTab = lc.Index
            Exit For
        End If
    Next lc

    If idxBomTab = 0 Then
        MsgBox "TBL_BOMS is missing the BOMTab column.", vbCritical, "Go To Specific BOM"
        Exit Sub
    End If

    rowOffset = selectedRow.Row - loBoms.DataBodyRange.Row + 1
    bomTabName = Trim$(CStr(loBoms.ListColumns(idxBomTab).DataBodyRange.Cells(rowOffset, 1).Value))

    If Len(bomTabName) = 0 Then
        MsgBox "Selected BOMS row has a blank BOMTab value.", vbExclamation, "Go To Specific BOM"
        Exit Sub
    End If

    Set wsTarget = Nothing
    On Error Resume Next
    Set wsTarget = wb.Worksheets(bomTabName)
    On Error GoTo EH

    If wsTarget Is Nothing Then
        MsgBox "Could not find worksheet '" & bomTabName & "' from the selected BOMS row.", vbExclamation, "Go To Specific BOM"
        Exit Sub
    End If

    wsTarget.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Exit Sub

EH:
    MsgBox "Go To Specific BOM failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, PROC_NAME
End Sub
