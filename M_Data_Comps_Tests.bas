Attribute VB_Name = "M_Data_Comps_Tests"
Option Explicit

'===============================================================================
' Module: M_Data_Comps_Tests
'
' Purpose:
'   Behavioral smoke tests for the "Create New Component" workflow.
'   These tests are intentionally NOT schema-structure tests; they validate behavior
'   that can regress even when schema validation passes.
'
'   Aligns to 050 Testing Strategy:
'     - Smoke Tests (e.g., Unique PN/Rev) + per-procedure Test_() harnesses.
'     - Complements (does not duplicate) Gate/Schema validation.
'
' Tests included:
'   1) Test_CompID_NextValueAndFormat
'   2) Test_Unique_OurPN_OurRev
'   3) Test_Rollback_Safety
'   4) Test_Supplier_Normalization_Match
'
' Inputs (Tabs/Tables/Headers):
'   - Comps sheet:      TBL_COMPS
'       Required headers: CompID, OurPN, OurRev, SupplierID, SupplierName
'   - Suppliers sheet:  TBL_SUPPLIERS
'       Required headers: SupplierID, SupplierName, SupplierDefaultLT
'
' Outputs / Side effects:
'   - Logs pass/fail to M_Core_Logging if available (otherwise Debug.Print)
'   - Shows a summary MsgBox
'   - Rollback test creates a temp row then deletes it (no permanent data change)
'
' Preconditions / Postconditions:
'   - Tables exist and have required headers (this module fails fast if not)
'
' Errors & Guards:
'   - Each test traps its own errors; suite continues to run remaining tests
'   - No Select/Activate
'
' Version: v3.5.0
' Author: Keith + GPT
' Date: 2025-12-30
'===============================================================================

Private Const SH_COMPS As String = "Comps"
Private Const LO_COMPS As String = "TBL_COMPS"

Private Const SH_SUPPLIERS As String = "Suppliers"
Private Const LO_SUPPLIERS As String = "TBL_SUPPLIERS"

Private Const COMP_ID_PREFIX As String = "COMP-"
Private Const COMP_ID_PAD As Long = 4

'==========================
' Public entry point
'==========================
Public Sub UI_Run_Comps_Tests()
    Const PROC_NAME As String = "M_Data_Comps_Tests.UI_Run_Comps_Tests"

    Dim passCount As Long, failCount As Long
    Dim summary As String

    On Error GoTo EH

    LogInfoSafe PROC_NAME, "Starting Comps test suite", vbNullString

    passCount = 0
    failCount = 0

    RunOneTest "Test_CompID_NextValueAndFormat", passCount, failCount
    RunOneTest "Test_Unique_OurPN_OurRev", passCount, failCount
    RunOneTest "Test_Rollback_Safety", passCount, failCount
    RunOneTest "Test_Supplier_Normalization_Match", passCount, failCount

    summary = "Comps Tests Complete" & vbCrLf & _
              "Passed: " & CStr(passCount) & vbCrLf & _
              "Failed: " & CStr(failCount)

    If failCount = 0 Then
        LogInfoSafe PROC_NAME, "All Comps tests passed", "Passed=" & CStr(passCount)
        MsgBox summary, vbInformation, "Comps Tests"
    Else
        LogWarnSafe PROC_NAME, "Some Comps tests failed", "Passed=" & CStr(passCount) & "; Failed=" & CStr(failCount)
        MsgBox summary & vbCrLf & vbCrLf & "See Log for details.", vbExclamation, "Comps Tests"
    End If

    Exit Sub

EH:
    LogErrorSafe PROC_NAME, "Unhandled error in test runner", Err.Number, Err.Description
    MsgBox "Comps test runner failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
End Sub

Private Sub RunOneTest(ByVal testName As String, ByRef passCount As Long, ByRef failCount As Long)
    Const PROC_NAME As String = "M_Data_Comps_Tests.RunOneTest"
    On Error GoTo EH

    Select Case testName
        Case "Test_CompID_NextValueAndFormat"
            Test_CompID_NextValueAndFormat
        Case "Test_Unique_OurPN_OurRev"
            Test_Unique_OurPN_OurRev
        Case "Test_Rollback_Safety"
            Test_Rollback_Safety
        Case "Test_Supplier_Normalization_Match"
            Test_Supplier_Normalization_Match
        Case Else
            Err.Raise vbObjectError + 7000, PROC_NAME, "Unknown test: " & testName
    End Select

    passCount = passCount + 1
    LogInfoSafe PROC_NAME, "PASS: " & testName, vbNullString
    Exit Sub

EH:
    failCount = failCount + 1
    LogErrorSafe PROC_NAME, "FAIL: " & testName, Err.Number, Err.Description
    MsgBox "Test failed: " & testName & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
End Sub

'===============================================================================
' Test 1: CompID format + next value is max+1
'===============================================================================
Private Sub Test_CompID_NextValueAndFormat()
    Const PROC_NAME As String = "M_Data_Comps_Tests.Test_CompID_NextValueAndFormat"

    Dim loComps As ListObject
    Dim nextId As String
    Dim maxN As Long
    Dim nextN As Long

    On Error GoTo EH

    Set loComps = ThisWorkbook.Worksheets(SH_COMPS).ListObjects(LO_COMPS)

    RequireColumn loComps, "CompID"

    maxN = GetMaxTrailingNumber(loComps, "CompID", COMP_ID_PREFIX)
    nextId = GenerateNextId(loComps, "CompID", COMP_ID_PREFIX, COMP_ID_PAD)

    If Len(nextId) = 0 Then Err.Raise vbObjectError + 7101, PROC_NAME, "GenerateNextId returned blank."

    If Not (UCase$(Left$(nextId, Len(COMP_ID_PREFIX))) = UCase$(COMP_ID_PREFIX)) Then
        Err.Raise vbObjectError + 7102, PROC_NAME, "CompID prefix wrong: " & nextId
    End If

    If Not (nextId Like "COMP-####") Then
        Err.Raise vbObjectError + 7103, PROC_NAME, "CompID format must be COMP-#### but got: " & nextId
    End If

    nextN = TrailingNumber(nextId)
    If nextN <> (maxN + 1) Then
        Err.Raise vbObjectError + 7104, PROC_NAME, "Expected next trailing number " & CStr(maxN + 1) & " but got " & CStr(nextN) & " (" & nextId & ")"
    End If

    Exit Sub

EH:
    LogErrorSafe PROC_NAME, "Error", Err.Number, Err.Description
    MsgBox "Test_CompID_NextValueAndFormat failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

'===============================================================================
' Test 2: OurPN + OurRev uniqueness detection
'===============================================================================
Private Sub Test_Unique_OurPN_OurRev()
    Const PROC_NAME As String = "M_Data_Comps_Tests.Test_Unique_OurPN_OurRev"

    Dim loComps As ListObject
    Dim pn As String, rev As String
    Dim existsBefore As Boolean

    On Error GoTo EH

    Set loComps = ThisWorkbook.Worksheets(SH_COMPS).ListObjects(LO_COMPS)

    RequireColumn loComps, "OurPN"
    RequireColumn loComps, "OurRev"

    ' Pick a combo that is extremely likely not to exist
    pn = "ZZZ-TEST-PN-" & Format$(Now, "yyyymmddhhnnss")
    rev = "A"

    existsBefore = PNRevComboExists(loComps, pn, rev)
    If existsBefore Then
        Err.Raise vbObjectError + 7201, PROC_NAME, "Unexpected: test PN/Rev already exists (should not)."
    End If

    ' Add then verify exists
    AppendTempPNRevRow loComps, pn, rev
    If Not PNRevComboExists(loComps, pn, rev) Then
        Err.Raise vbObjectError + 7202, PROC_NAME, "PNRevComboExists did not detect inserted test row."
    End If

    ' Cleanup (remove the temp row)
    DeleteRowsByPNRev loComps, pn, rev

    Exit Sub

EH:
    ' Try to cleanup on failure
    On Error Resume Next
    If Not loComps Is Nothing Then DeleteRowsByPNRev loComps, pn, rev
    On Error GoTo 0

    LogErrorSafe PROC_NAME, "Error", Err.Number, Err.Description
    MsgBox "Test_Unique_OurPN_OurRev failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

'===============================================================================
' Test 3: Rollback safety (create temp row then force error and ensure deleted)
'===============================================================================
Private Sub Test_Rollback_Safety()
    Const PROC_NAME As String = "M_Data_Comps_Tests.Test_Rollback_Safety"

    Dim loComps As ListObject
    Dim startCount As Long, endCount As Long
    Dim testCompId As String

    On Error GoTo EH

    Set loComps = ThisWorkbook.Worksheets(SH_COMPS).ListObjects(LO_COMPS)
    RequireColumn loComps, "CompID"

    startCount = loComps.ListRows.Count
    testCompId = "COMP-9999" ' intentionally obvious; test will delete it if created

    ' Ensure not already present (if it is, pick a unique one)
    If ValueExistsInColumn(loComps, "CompID", testCompId) Then
        testCompId = "COMP-TEST-" & Format$(Now, "hhnnss")
    End If

    ' Simulate: add row then throw error, then rollback in handler
    CreateRowThenForceErrorAndRollback loComps, testCompId

    endCount = loComps.ListRows.Count

    If endCount <> startCount Then
        Err.Raise vbObjectError + 7301, PROC_NAME, "Row count changed after rollback. Start=" & CStr(startCount) & ", End=" & CStr(endCount)
    End If

    ' Also ensure CompID not present
    If ValueExistsInColumn(loComps, "CompID", testCompId) Then
        Err.Raise vbObjectError + 7302, PROC_NAME, "Rollback failed; CompID still present: " & testCompId
    End If

    Exit Sub

EH:
    ' Best-effort cleanup
    On Error Resume Next
    If Not loComps Is Nothing Then DeleteRowsByCompId loComps, testCompId
    On Error GoTo 0

    LogErrorSafe PROC_NAME, "Error", Err.Number, Err.Description
    MsgBox "Test_Rollback_Safety failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

'===============================================================================
' Test 4: Supplier normalization/match behavior (logic-level)
'===============================================================================
Private Sub Test_Supplier_Normalization_Match()
    Const PROC_NAME As String = "M_Data_Comps_Tests.Test_Supplier_Normalization_Match"

    Dim a As String, b As String
    Dim na As String, nb As String

    On Error GoTo EH

    ' Core case you called out: B&B should match "B and B ..." after normalization
    a = "B&B"
    b = "B and B Thread"

    na = NormalizeForMatch(a)
    nb = NormalizeForMatch(b)

    If Len(na) = 0 Or Len(nb) = 0 Then
        Err.Raise vbObjectError + 7401, PROC_NAME, "Normalization produced blank output unexpectedly."
    End If

    If InStr(1, nb, na, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 7402, PROC_NAME, _
            "Expected normalized '" & a & "' (" & na & ") to be contained in normalized '" & b & "' (" & nb & ")."
    End If

    Exit Sub

EH:
    LogErrorSafe PROC_NAME, "Error", Err.Number, Err.Description
    MsgBox "Test_Supplier_Normalization_Match failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

'===============================================================================
' Helpers: non-interactive utilities
'===============================================================================

Private Sub AppendTempPNRevRow(ByVal loComps As ListObject, ByVal pn As String, ByVal rev As String)
    Dim lr As ListRow
    Set lr = loComps.ListRows.Add
    If GetColIndex(loComps, "OurPN") > 0 Then lr.Range.Cells(1, GetColIndex(loComps, "OurPN")).value = pn
    If GetColIndex(loComps, "OurRev") > 0 Then lr.Range.Cells(1, GetColIndex(loComps, "OurRev")).value = rev
End Sub

Private Sub DeleteRowsByPNRev(ByVal loComps As ListObject, ByVal pn As String, ByVal rev As String)
    Dim idxPN As Long, idxRev As Long
    Dim i As Long

    idxPN = GetColIndex(loComps, "OurPN")
    idxRev = GetColIndex(loComps, "OurRev")
    If idxPN = 0 Or idxRev = 0 Then Exit Sub

    For i = loComps.ListRows.Count To 1 Step -1
        If StrComp(Trim$(CStr(loComps.ListRows(i).Range.Cells(1, idxPN).value)), Trim$(pn), vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(loComps.ListRows(i).Range.Cells(1, idxRev).value)), Trim$(rev), vbTextCompare) = 0 Then
            loComps.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub DeleteRowsByCompId(ByVal loComps As ListObject, ByVal compId As String)
    Dim idx As Long, i As Long
    idx = GetColIndex(loComps, "CompID")
    If idx = 0 Then Exit Sub

    For i = loComps.ListRows.Count To 1 Step -1
        If StrComp(Trim$(CStr(loComps.ListRows(i).Range.Cells(1, idx).value)), Trim$(compId), vbTextCompare) = 0 Then
            loComps.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub CreateRowThenForceErrorAndRollback(ByVal loComps As ListObject, ByVal compId As String)
    Const PROC_NAME As String = "M_Data_Comps_Tests.CreateRowThenForceErrorAndRollback"
    Dim lr As ListRow
    On Error GoTo EH

    Set lr = loComps.ListRows.Add
    lr.Range.Cells(1, GetColIndexOrRaise(loComps, "CompID")).value = compId

    ' Force a failure
    Err.Raise vbObjectError + 7999, PROC_NAME, "Intentional failure to validate rollback."

EH:
    ' Rollback
    On Error Resume Next
    If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    MsgBox "CreateRowThenForceErrorAndRollback cleanup failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comps Tests"
End Sub

'===============================================================================
' Core logic helpers (mirrors what your entry module uses)
'===============================================================================

Private Function PNRevComboExists(ByVal lo As ListObject, ByVal ourPN As String, ByVal ourRev As String) As Boolean
    Dim idxPN As Long, idxRev As Long
    Dim arrPN As Variant, arrRev As Variant
    Dim i As Long

    PNRevComboExists = False
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxPN = GetColIndex(lo, "OurPN")
    idxRev = GetColIndex(lo, "OurRev")
    If idxPN = 0 Or idxRev = 0 Then Exit Function

    arrPN = lo.ListColumns(idxPN).DataBodyRange.value
    arrRev = lo.ListColumns(idxRev).DataBodyRange.value

    For i = 1 To UBound(arrPN, 1)
        If StrComp(Trim$(CStr(arrPN(i, 1))), Trim$(ourPN), vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(arrRev(i, 1))), Trim$(ourRev), vbTextCompare) = 0 Then
            PNRevComboExists = True
            Exit Function
        End If
    Next i
End Function

Private Function GenerateNextId(ByVal lo As ListObject, ByVal header As String, ByVal prefix As String, ByVal padDigits As Long) As String
    Dim idx As Long
    Dim maxN As Long
    Dim arr As Variant
    Dim i As Long
    Dim s As String
    Dim n As Long

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

Private Function GetMaxTrailingNumber(ByVal lo As ListObject, ByVal header As String, ByVal prefix As String) As Long
    Dim idx As Long
    Dim arr As Variant
    Dim i As Long
    Dim s As String
    Dim n As Long
    Dim maxN As Long

    GetMaxTrailingNumber = 0
    idx = GetColIndex(lo, header)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    maxN = 0
    arr = lo.ListColumns(idx).DataBodyRange.value

    For i = 1 To UBound(arr, 1)
        s = Trim$(CStr(arr(i, 1)))
        If Len(prefix) = 0 Or UCase$(Left$(s, Len(prefix))) = UCase$(prefix) Then
            n = TrailingNumber(s)
            If n > maxN Then maxN = n
        End If
    Next i

    GetMaxTrailingNumber = maxN
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

Private Function NormalizeForMatch(ByVal s As String) As String
    Dim t As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    t = UCase$(Trim$(Replace(CStr(s), ChrW(160), " ")))
    t = Replace(t, "&", " AND ")

    out = vbNullString
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        Else
            out = out & " "
        End If
    Next i

    NormalizeForMatch = CollapseSpaces(out)
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = t
End Function

'===============================================================================
' Table helpers
'===============================================================================

Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 7600, "RequireColumn", "Missing column '" & header & "' in table '" & lo.Name & "'."
    End If
End Sub

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

Private Function GetColIndexOrRaise(ByVal lo As ListObject, ByVal header As String) As Long
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Err.Raise vbObjectError + 7601, "GetColIndexOrRaise", "Missing required header: " & header
    GetColIndexOrRaise = idx
End Function

Private Function ValueExistsInColumn(ByVal lo As ListObject, ByVal header As String, ByVal valueText As String) As Boolean
    Dim idx As Long
    Dim rng As Range

    ValueExistsInColumn = False
    idx = GetColIndex(lo, header)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set rng = lo.ListColumns(idx).DataBodyRange
    ValueExistsInColumn = (Application.WorksheetFunction.CountIf(rng, valueText) > 0)
End Function

'===============================================================================
' Logging wrappers (use M_Core_Logging if present, otherwise Debug.Print)
'===============================================================================

Private Sub LogInfoSafe(ByVal procName As String, ByVal message As String, ByVal details As String)
    On Error GoTo Fallback
    M_Core_Logging.LogInfo procName, message, details
    Exit Sub
Fallback:
    Debug.Print Now & " [INFO] " & procName & " - " & message & IIf(Len(details) > 0, " | " & details, vbNullString)
End Sub

Private Sub LogWarnSafe(ByVal procName As String, ByVal message As String, ByVal details As String)
    On Error GoTo Fallback
    M_Core_Logging.LogWarn procName, message, details
    Exit Sub
Fallback:
    Debug.Print Now & " [WARN] " & procName & " - " & message & IIf(Len(details) > 0, " | " & details, vbNullString)
End Sub

Private Sub LogErrorSafe(ByVal procName As String, ByVal message As String, ByVal errNum As Long, ByVal errDesc As String)
    On Error GoTo Fallback
    M_Core_Logging.LogError procName, message, "Err " & CStr(errNum) & ": " & errDesc, errNum
    Exit Sub
Fallback:
    Debug.Print Now & " [ERR ] " & procName & " - " & message & " | Err " & CStr(errNum) & ": " & errDesc
End Sub


