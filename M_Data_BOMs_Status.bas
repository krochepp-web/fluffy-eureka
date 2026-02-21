Attribute VB_Name = "M_Data_BOMs_Status"
Option Explicit

Public Function GetBomStatus(ByVal bomId As String) As String
    Dim wb As Workbook
    Dim wsBoms As Worksheet
    Dim loBoms As ListObject
    Dim rowIndex As Long
    Dim idxStatus As Long
    Dim statusVal As String

    GetBomStatus = vbNullString
    bomId = Trim$(bomId)
    If Len(bomId) = 0 Then Exit Function

    On Error GoTo CleanFail

    Set wb = ThisWorkbook
    Set wsBoms = wb.Worksheets(M_Core_Constants.SH_BOMS)
    Set loBoms = wsBoms.ListObjects(M_Core_Constants.TBL_BOMS)

    rowIndex = FindBomRowIndex(loBoms, bomId)
    If rowIndex = 0 Then Exit Function

    idxStatus = ResolveBomStatusColumnIndex(loBoms)
    If idxStatus = 0 Then
        GetBomStatus = M_Core_Constants.BOM_STATUS_DRAFT
        Exit Function
    End If

    statusVal = SafeText(loBoms.ListColumns(idxStatus).DataBodyRange.Cells(rowIndex, 1).Value)
    If Len(statusVal) = 0 Then statusVal = M_Core_Constants.BOM_STATUS_DRAFT

    GetBomStatus = NormalizeBomStatus(statusVal)
    Exit Function

CleanFail:
    GetBomStatus = vbNullString
End Function

Public Function CanEditBom(ByVal bomId As String) As Boolean
    Dim statusVal As String
    statusVal = GetBomStatus(bomId)
    CanEditBom = (StrComp(statusVal, M_Core_Constants.BOM_STATUS_DRAFT, vbTextCompare) = 0)
End Function

Public Function CanBuildBom(ByVal bomId As String) As Boolean
    Dim statusVal As String
    statusVal = GetBomStatus(bomId)

    CanBuildBom = (StrComp(statusVal, M_Core_Constants.BOM_STATUS_DRAFT, vbTextCompare) = 0 Or _
                   StrComp(statusVal, M_Core_Constants.BOM_STATUS_LOCK, vbTextCompare) = 0)
End Function

Public Function ValidateBomStatusTransition(ByVal oldStatus As String, ByVal newStatus As String, ByVal role As String) As Boolean
    Dim oldNorm As String
    Dim newNorm As String
    Dim roleNorm As String

    oldNorm = NormalizeBomStatus(oldStatus)
    newNorm = NormalizeBomStatus(newStatus)
    roleNorm = UCase$(Trim$(role))

    ValidateBomStatusTransition = False

    If Len(oldNorm) = 0 Or Len(newNorm) = 0 Then
        MsgBox "Invalid BOM status transition: unknown status value.", vbOKOnly, "BOM Status"
        Exit Function
    End If

    If StrComp(oldNorm, newNorm, vbTextCompare) = 0 Then
        ValidateBomStatusTransition = True
        Exit Function
    End If

    Select Case oldNorm
        Case M_Core_Constants.BOM_STATUS_DRAFT
            ValidateBomStatusTransition = (StrComp(newNorm, M_Core_Constants.BOM_STATUS_LOCK, vbTextCompare) = 0 Or _
                                           StrComp(newNorm, M_Core_Constants.BOM_STATUS_OBSOLETE, vbTextCompare) = 0)

        Case M_Core_Constants.BOM_STATUS_LOCK
            If IsBomAdminRole(roleNorm) Then
                ValidateBomStatusTransition = (StrComp(newNorm, M_Core_Constants.BOM_STATUS_DRAFT, vbTextCompare) = 0 Or _
                                               StrComp(newNorm, M_Core_Constants.BOM_STATUS_OBSOLETE, vbTextCompare) = 0)
            Else
                ValidateBomStatusTransition = (StrComp(newNorm, M_Core_Constants.BOM_STATUS_OBSOLETE, vbTextCompare) = 0)
            End If

        Case M_Core_Constants.BOM_STATUS_OBSOLETE
            If IsBomAdminRole(roleNorm) Then
                ValidateBomStatusTransition = (StrComp(newNorm, M_Core_Constants.BOM_STATUS_DRAFT, vbTextCompare) = 0 Or _
                                               StrComp(newNorm, M_Core_Constants.BOM_STATUS_LOCK, vbTextCompare) = 0)
            Else
                ValidateBomStatusTransition = False
            End If
    End Select

    If Not ValidateBomStatusTransition Then
        MsgBox "Status transition blocked: " & oldNorm & " -> " & newNorm & _
               ". Role '" & role & "' is not permitted.", vbOKOnly, "BOM Status"
    End If
End Function

Public Function GetBomIdByTabName(ByVal bomTabName As String) As String
    Dim wb As Workbook
    Dim wsBoms As Worksheet
    Dim loBoms As ListObject
    Dim idxBomId As Long
    Dim idxBomTab As Long
    Dim arrBomId As Variant
    Dim arrBomTab As Variant
    Dim i As Long

    GetBomIdByTabName = vbNullString
    bomTabName = Trim$(bomTabName)
    If Len(bomTabName) = 0 Then Exit Function

    On Error GoTo CleanFail

    Set wb = ThisWorkbook
    Set wsBoms = wb.Worksheets(M_Core_Constants.SH_BOMS)
    Set loBoms = wsBoms.ListObjects(M_Core_Constants.TBL_BOMS)

    If loBoms.DataBodyRange Is Nothing Then Exit Function

    idxBomId = GetColIndex(loBoms, M_Core_Constants.COL_BOM_ID)
    idxBomTab = GetColIndex(loBoms, "BOMTab")
    If idxBomId = 0 Or idxBomTab = 0 Then Exit Function

    arrBomId = loBoms.ListColumns(idxBomId).DataBodyRange.Value
    arrBomTab = loBoms.ListColumns(idxBomTab).DataBodyRange.Value

    For i = 1 To UBound(arrBomTab, 1)
        If StrComp(SafeText(arrBomTab(i, 1)), bomTabName, vbTextCompare) = 0 Then
            GetBomIdByTabName = SafeText(arrBomId(i, 1))
            Exit Function
        End If
    Next i
    Exit Function

CleanFail:
    GetBomIdByTabName = vbNullString
End Function

Public Function GetBomEditDisabledMessage(ByVal bomId As String) As String
    Dim statusVal As String

    statusVal = GetBomStatus(bomId)
    If Len(statusVal) = 0 Then
        GetBomEditDisabledMessage = "Unable to determine BOM status; edits are disabled."
        Exit Function
    End If

    If StrComp(statusVal, M_Core_Constants.BOM_STATUS_LOCK, vbTextCompare) = 0 Then
        GetBomEditDisabledMessage = "BOM is LOCK; edits are disabled"
    ElseIf StrComp(statusVal, M_Core_Constants.BOM_STATUS_OBSOLETE, vbTextCompare) = 0 Then
        GetBomEditDisabledMessage = "BOM is OBSOLETE; edits are disabled"
    Else
        GetBomEditDisabledMessage = "BOM status '" & statusVal & "' does not allow edits."
    End If
End Function

Private Function FindBomRowIndex(ByVal loBoms As ListObject, ByVal bomId As String) As Long
    Dim idxBomId As Long
    Dim arrBomId As Variant
    Dim i As Long

    FindBomRowIndex = 0
    If loBoms Is Nothing Then Exit Function
    If loBoms.DataBodyRange Is Nothing Then Exit Function

    idxBomId = GetColIndex(loBoms, M_Core_Constants.COL_BOM_ID)
    If idxBomId = 0 Then Exit Function

    arrBomId = loBoms.ListColumns(idxBomId).DataBodyRange.Value
    For i = 1 To UBound(arrBomId, 1)
        If StrComp(SafeText(arrBomId(i, 1)), bomId, vbTextCompare) = 0 Then
            FindBomRowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function ResolveBomStatusColumnIndex(ByVal loBoms As ListObject) As Long
    ResolveBomStatusColumnIndex = GetColIndex(loBoms, M_Core_Constants.COL_BOM_STATUS)
    If ResolveBomStatusColumnIndex = 0 Then
        ResolveBomStatusColumnIndex = GetColIndex(loBoms, "Status")
    End If
End Function

Private Function NormalizeBomStatus(ByVal statusVal As String) As String
    Dim s As String
    s = UCase$(Trim$(statusVal))

    Select Case s
        Case UCase$(M_Core_Constants.BOM_STATUS_DRAFT)
            NormalizeBomStatus = M_Core_Constants.BOM_STATUS_DRAFT
        Case UCase$(M_Core_Constants.BOM_STATUS_LOCK)
            NormalizeBomStatus = M_Core_Constants.BOM_STATUS_LOCK
        Case UCase$(M_Core_Constants.BOM_STATUS_OBSOLETE)
            NormalizeBomStatus = M_Core_Constants.BOM_STATUS_OBSOLETE
        Case Else
            NormalizeBomStatus = vbNullString
    End Select
End Function

Private Function IsBomAdminRole(ByVal roleNorm As String) As Boolean
    IsBomAdminRole = (roleNorm = "ADMIN" Or roleNorm = "ENGINEER")
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

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = vbNullString
    ElseIf IsNull(v) Then
        SafeText = vbNullString
    Else
        SafeText = Trim$(CStr(v))
    End If
End Function
