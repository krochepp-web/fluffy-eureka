Attribute VB_Name = "Dev_NamedRange_Diagnostics"
Option Explicit

'===============================================================================
' Module: Dev_NamedRange_Diagnostics
'
' Purpose:
'   Developer diagnostics for named-range issues that can break list prompts
'   (e.g., NR_RevStatus errors during new component creation).
'
' Entry point:
'   DEV_AuditNamedRanges
'
' Output:
'   Writes findings to sheet: DEV_NR_AUDIT
'===============================================================================

Public Sub DEV_AuditNamedRanges()
    Const PROC_NAME As String = "Dev_NamedRange_Diagnostics.DEV_AuditNamedRanges"
    Const OUT_SHEET As String = "DEV_NR_AUDIT"

    Dim wb As Workbook
    Dim wsOut As Worksheet
    Dim targets As Variant
    Dim target As Variant
    Dim rowOut As Long

    On Error GoTo EH

    Set wb = ThisWorkbook
    targets = Array("NR_RevStatus", "NR_UOM", "NR_IMSStatus")

    Set wsOut = PrepareOutputSheet(wb, OUT_SHEET)
    rowOut = 2

    Dim i As Long
    For i = LBound(targets) To UBound(targets)
        target = CStr(targets(i))
        rowOut = AuditOneNamedRange(wb, wsOut, rowOut, CStr(target))
    Next i

    wsOut.Columns("A:N").EntireColumn.AutoFit
    wsOut.Activate

    MsgBox "Named-range audit complete." & vbCrLf & _
           "Sheet: " & OUT_SHEET & vbCrLf & _
           "Check rows flagged as DUPLICATE, INVALID_REF, or NO_MATCH.", vbOKOnly, "DEV_AuditNamedRanges"
    Exit Sub

EH:
    MsgBox PROC_NAME & " failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "DEV_AuditNamedRanges"
End Sub

Private Function PrepareOutputSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "TargetName"
    ws.Range("B1").Value = "Status"
    ws.Range("C1").Value = "ScopeType"
    ws.Range("D1").Value = "ScopeName"
    ws.Range("E1").Value = "NameObject"
    ws.Range("F1").Value = "Visible"
    ws.Range("G1").Value = "RefersTo"
    ws.Range("H1").Value = "ResolvesToRange"
    ws.Range("I1").Value = "RangeAddress"
    ws.Range("J1").Value = "Rows"
    ws.Range("K1").Value = "Cols"
    ws.Range("L1").Value = "TopLeftValue"
    ws.Range("M1").Value = "RefersToRangeError"
    ws.Range("N1").Value = "Notes"

    ws.Range("A1:N1").Font.Bold = True

    Set PrepareOutputSheet = ws
End Function

Private Function AuditOneNamedRange(ByVal wb As Workbook, ByVal wsOut As Worksheet, ByVal rowStart As Long, ByVal targetName As String) As Long
    Dim rowOut As Long
    Dim matches As Long
    Dim nm As Name
    Dim ws As Worksheet

    rowOut = rowStart
    matches = 0

    For Each nm In wb.Names
        If NameMatches(nm.Name, targetName) Then
            matches = matches + 1
            WriteNameRow wsOut, rowOut, targetName, "Workbook", wb.Name, nm
            rowOut = rowOut + 1
        End If
    Next nm

    For Each ws In wb.Worksheets
        For Each nm In ws.Names
            If NameMatches(nm.Name, targetName) Then
                matches = matches + 1
                WriteNameRow wsOut, rowOut, targetName, "Worksheet", ws.Name, nm
                rowOut = rowOut + 1
            End If
        Next nm
    Next ws

    If matches = 0 Then
        wsOut.Cells(rowOut, 1).Value = targetName
        wsOut.Cells(rowOut, 2).Value = "NO_MATCH"
        wsOut.Cells(rowOut, 14).Value = "No name object found at workbook or worksheet scope."
        rowOut = rowOut + 1
    ElseIf matches > 1 Then
        wsOut.Cells(rowStart, 14).Value = AppendNote(CStr(wsOut.Cells(rowStart, 14).Value), "DUPLICATE_NAME_COUNT=" & CStr(matches))
        wsOut.Cells(rowStart, 2).Value = "DUPLICATE"
    End If

    AuditOneNamedRange = rowOut
End Function

Private Sub WriteNameRow(ByVal wsOut As Worksheet, ByVal rowOut As Long, ByVal targetName As String, ByVal scopeType As String, ByVal scopeName As String, ByVal nm As Name)
    Dim rng As Range
    Dim resolveOk As Boolean
    Dim errNum As Long
    Dim errDesc As String

    wsOut.Cells(rowOut, 1).Value = targetName
    wsOut.Cells(rowOut, 2).Value = "OK"
    wsOut.Cells(rowOut, 3).Value = scopeType
    wsOut.Cells(rowOut, 4).Value = scopeName
    wsOut.Cells(rowOut, 5).Value = nm.Name
    wsOut.Cells(rowOut, 6).Value = CStr(nm.Visible)
    wsOut.Cells(rowOut, 7).Value = SafeNameRefersTo(nm)

    resolveOk = TryGetRangeFromName(nm, rng, errNum, errDesc)
    wsOut.Cells(rowOut, 8).Value = IIf(resolveOk, "TRUE", "FALSE")

    If resolveOk Then
        wsOut.Cells(rowOut, 9).Value = rng.Address(External:=True)
        wsOut.Cells(rowOut, 10).Value = rng.Rows.Count
        wsOut.Cells(rowOut, 11).Value = rng.Columns.Count
        wsOut.Cells(rowOut, 12).Value = SafeTopLeftValue(rng)
    Else
        wsOut.Cells(rowOut, 2).Value = "INVALID_REF"
        wsOut.Cells(rowOut, 13).Value = "Err " & CStr(errNum) & ": " & errDesc
    End If

    If IsSuspiciousRefersTo(CStr(wsOut.Cells(rowOut, 7).Value)) Then
        wsOut.Cells(rowOut, 14).Value = AppendNote(CStr(wsOut.Cells(rowOut, 14).Value), "SUSPICIOUS_REFERS_TO")
    End If
End Sub

Private Function TryGetRangeFromName(ByVal nm As Name, ByRef rng As Range, ByRef errNum As Long, ByRef errDesc As String) As Boolean
    On Error Resume Next
    Set rng = nm.RefersToRange
    errNum = Err.Number
    errDesc = Err.Description
    TryGetRangeFromName = (errNum = 0 And Not rng Is Nothing)
    On Error GoTo 0
End Function

Private Function NameMatches(ByVal qualifiedName As String, ByVal targetName As String) As Boolean
    Dim bangPos As Long
    Dim baseName As String

    baseName = Trim$(qualifiedName)
    bangPos = InStrRev(baseName, "!")
    If bangPos > 0 Then baseName = Mid$(baseName, bangPos + 1)

    If Left$(baseName, 1) = "=" Then baseName = Mid$(baseName, 2)
    If Left$(baseName, 1) = "'" And Right$(baseName, 1) = "'" And Len(baseName) >= 2 Then
        baseName = Mid$(baseName, 2, Len(baseName) - 2)
    End If

    NameMatches = (StrComp(baseName, targetName, vbTextCompare) = 0)
End Function

Private Function IsSuspiciousRefersTo(ByVal refersToText As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(refersToText))

    If InStr(1, t, "#REF!", vbTextCompare) > 0 Then
        IsSuspiciousRefersTo = True
        Exit Function
    End If

    If InStr(1, t, ".XLS", vbTextCompare) > 0 Then
        IsSuspiciousRefersTo = True
        Exit Function
    End If

    IsSuspiciousRefersTo = False
End Function

Private Function SafeNameRefersTo(ByVal nm As Name) As String
    On Error GoTo EH
    SafeNameRefersTo = nm.RefersTo
    Exit Function
EH:
    SafeNameRefersTo = "<error " & Err.Number & ": " & Err.Description & ">"
End Function

Private Function SafeTopLeftValue(ByVal rng As Range) As String
    On Error GoTo EH
    If rng Is Nothing Then
        SafeTopLeftValue = vbNullString
    Else
        SafeTopLeftValue = CStr(rng.Cells(1, 1).Value)
    End If
    Exit Function
EH:
    SafeTopLeftValue = "<error " & Err.Number & ": " & Err.Description & ">"
End Function

Private Function AppendNote(ByVal currentNote As String, ByVal newNote As String) As String
    If Len(Trim$(currentNote)) = 0 Then
        AppendNote = newNote
    Else
        AppendNote = currentNote & "; " & newNote
    End If
End Function
