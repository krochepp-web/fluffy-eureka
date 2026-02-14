Attribute VB_Name = "M_Data_Comps_Entry"
Option Explicit

'===============================================================================
' Module: M_Data_Comps_Entry
'
' Purpose:
'   Create a new Component record in Comps.TBL_COMPS with:
'     - Gate check (blocks if workbook not ready)
'     - Required-field prompting (ComponentDescription REQUIRED)
'     - Forced-valid Supplier selection (search by SupplierName; forgiving matching)
'     - Supplier picker shows numbered options when multiple matches
'     - If exactly 1 match, prompts user to confirm (Yes=accept, No=re-search)
'     - Auto-population of audit fields (CreatedAt/By, UpdatedAt/By)
'     - Basic numeric validation (MOQ1, CostPerUOMMOQ1, ComponentLT)
'     - List prompts show all options (NR_UOM, NR_RevStatus, NR_IMSStatus)
'     - IMSStatus default is sourced from TBL_SCHEMA.DefaultValue when available
'     - IsBuildable computed at creation
'     - If macro ends without creating a record, user gets a brief message
'
' Inputs (Tabs/Tables/Headers):
'   - Comps sheet:      TBL_COMPS
'       Required headers:
'         CompID, OurPN, OurRev, ComponentDescription,
'         SupplierID, SupplierName,
'         UOM, RevStatus, IMSStatus, MOQ1, CostPerUOMMOQ1,
'         CreatedAt, CreatedBy, UpdatedAt, UpdatedBy, IsBuildable
'       Optional headers:
'         SupplierLeadTime, ComponentLT
'   - Suppliers sheet:  TBL_SUPPLIERS
'       Required headers:
'         SupplierID, SupplierName, SupplierDefaultLT
'   - Schema table:     TBL_SCHEMA
'       Used fields: TAB_NAME, TABLE_NAME, COLUMN_HEADER, DefaultValue
'
' Named Ranges (list prompts):
'   - NR_UOM
'   - NR_RevStatus
'   - NR_IMSStatus
'
' Outputs / Side effects:
'   - Appends one row to Comps.TBL_COMPS
'   - Rolls back (deletes) the inserted row if user cancels or errors
'
' Preconditions / Postconditions:
'   - Gate passes (schema + integrity) before any write occurs
'
' Errors & Guards:
'   - Fails fast on missing tables/headers/named ranges with clear messages
'
' Version: v3.6.2
' Author: Keith + GPT
' Date: 2025-12-30
'===============================================================================

'==========================
' PUBLIC ENTRY POINTS
'==========================
Public Sub NewComponent()
    RunNewComponent
End Sub

Private Sub RunNewComponent()
    Const PROC_NAME As String = "M_Data_Comps_Entry.RunNewComponent"

    Const SH_COMPS As String = "Comps"
    Const LO_COMPS As String = "TBL_COMPS"

    Const SH_SUPPLIERS As String = "Suppliers"
    Const LO_SUPPLIERS As String = "TBL_SUPPLIERS"

    Const COMP_ID_PREFIX As String = "COMP-"
    Const COMP_ID_PAD As Long = 4

    ' Defaults (hard-coded fallback; schema-driving later)
    Const DEFAULT_UOM As String = "each"
    Const DEFAULT_REVSTATUS As String = "Active"
    Const DEFAULT_IMSSTATUS_FALLBACK As String = "Released"  'Used only if schema default is blank/missing
    Const DEFAULT_COST_MOQ1 As Double = 0.01
    Const DEFAULT_MOQ1 As Long = 1
    Const DEFAULT_DESC As String = "Description is Missing"

    Dim wb As Workbook
    Dim wsComps As Worksheet, wsSupp As Worksheet
    Dim loComps As ListObject, loSupp As ListObject
    Dim lr As ListRow

    Dim compId As String
    Dim ourPN As String, ourRev As String
    Dim desc As String

    Dim pickId As String, pickName As String, pickDfltLT As Variant
    Dim uom As String, revStatus As String, imsStatus As String
    Dim imsDefault As String
    Dim moq1 As Long
    Dim costMOQ1 As Double
    Dim createdAt As Date, createdBy As String

    Dim createdOk As Boolean
    Dim abortedReason As String

    createdOk = False
    abortedReason = vbNullString

    On Error GoTo EH

    '-----------------------------
    ' Gate check (consistency)
    '-----------------------------
    If Not GateReady_Safe(True) Then
        abortedReason = "Gate not ready."
        GoTo Aborted
    End If

    Set wb = ThisWorkbook
    Set wsComps = wb.Worksheets(SH_COMPS)
    Set wsSupp = wb.Worksheets(SH_SUPPLIERS)

    Set loComps = wsComps.ListObjects(LO_COMPS)
    Set loSupp = wsSupp.ListObjects(LO_SUPPLIERS)

    ' Guards: required columns
    RequireColumn loComps, "CompID"
    RequireColumn loComps, "OurPN"
    RequireColumn loComps, "OurRev"
    RequireColumn loComps, "ComponentDescription"
    RequireColumn loComps, "SupplierID"
    RequireColumn loComps, "SupplierName"

    RequireColumn loComps, "UOM"
    RequireColumn loComps, "RevStatus"
    RequireColumn loComps, "IMSStatus"
    RequireColumn loComps, "MOQ1"
    RequireColumn loComps, "CostPerUOMMOQ1"
    RequireColumn loComps, "CreatedAt"
    RequireColumn loComps, "CreatedBy"
    RequireColumn loComps, "UpdatedAt"
    RequireColumn loComps, "UpdatedBy"
    RequireColumn loComps, "IsBuildable"

    RequireColumn loSupp, "SupplierID"
    RequireColumn loSupp, "SupplierName"
    RequireColumn loSupp, "SupplierDefaultLT"

    ' Guards: required named ranges
    RequireNamedRange "NR_UOM"
    RequireNamedRange "NR_RevStatus"
    RequireNamedRange "NR_IMSStatus"

    ' IMS default: attempt schema lookup; fallback if blank
    imsDefault = GetSchemaDefaultValue("Comps", "TBL_COMPS", "IMSStatus")
    imsDefault = Trim$(imsDefault)
    If Len(imsDefault) = 0 Then imsDefault = DEFAULT_IMSSTATUS_FALLBACK

    ' Generate CompID
    compId = GenerateNextId(loComps, "CompID", COMP_ID_PREFIX, COMP_ID_PAD)
    If Len(compId) = 0 Then Err.Raise vbObjectError + 5100, PROC_NAME, "Failed to generate CompID."
    If ValueExistsInColumn(loComps, "CompID", compId) Then Err.Raise vbObjectError + 5101, PROC_NAME, "Generated CompID already exists: " & compId

    ' Create row now; rollback on cancel/error
    Set lr = loComps.ListRows.Add
    SetByHeader loComps, lr, "CompID", compId

    ' Audit fields
    createdAt = Now
    createdBy = GetUserNameSafe()

    SetByHeader loComps, lr, "CreatedAt", createdAt
    SetByHeader loComps, lr, "CreatedBy", createdBy
    SetByHeader loComps, lr, "UpdatedAt", createdAt
    SetByHeader loComps, lr, "UpdatedBy", createdBy

    ' Required: OurPN / OurRev
    ourPN = Trim$(InputBox("Enter OurPN (required).", "New Component (" & compId & ")"))
    If Len(ourPN) = 0 Then
        abortedReason = "OurPN not provided."
        GoTo FailRollback
    End If

    ourRev = Trim$(InputBox("Enter OurRev (required).", "New Component (" & compId & ")"))
    If Len(ourRev) = 0 Then
        abortedReason = "OurRev not provided."
        GoTo FailRollback
    End If

    If PNRevComboExists(loComps, ourPN, ourRev) Then
        MsgBox "OurPN + OurRev must be unique." & vbCrLf & _
               "That combination already exists:" & vbCrLf & _
               "OurPN=" & ourPN & vbCrLf & _
               "OurRev=" & ourRev, vbExclamation, "New Component"
        abortedReason = "Duplicate OurPN+OurRev."
        GoTo FailRollback
    End If

    SetByHeader loComps, lr, "OurPN", ourPN
    SetByHeader loComps, lr, "OurRev", ourRev

    ' Required: ComponentDescription
    desc = Prompt_RequiredText("Enter ComponentDescription (required).", "New Component (" & compId & ")", DEFAULT_DESC)
    If Len(desc) = 0 Then
        abortedReason = "ComponentDescription not provided."
        GoTo FailRollback
    End If
    SetByHeader loComps, lr, "ComponentDescription", desc

    ' Supplier selection (forced)
    If Not SupplierPick_ByName(loSupp, pickId, pickName, pickDfltLT) Then
        abortedReason = "Supplier selection cancelled."
        GoTo FailRollback
    End If

    SetByHeader loComps, lr, "SupplierID", pickId
    SetByHeader loComps, lr, "SupplierName", pickName

    If ColumnExists(loComps, "SupplierLeadTime") Then
        SetByHeader loComps, lr, "SupplierLeadTime", pickDfltLT
    End If

    ' Required list fields
    uom = Prompt_ListValue("NR_UOM", "Select UOM (required).", "New Component (" & compId & ")", DEFAULT_UOM)
    If Len(uom) = 0 Then
        abortedReason = "UOM not selected."
        GoTo FailRollback
    End If
    SetByHeader loComps, lr, "UOM", uom

    revStatus = Prompt_ListValue("NR_RevStatus", "Select RevStatus (required).", "New Component (" & compId & ")", DEFAULT_REVSTATUS)
    If Len(revStatus) = 0 Then
        abortedReason = "RevStatus not selected."
        GoTo FailRollback
    End If
    SetByHeader loComps, lr, "RevStatus", revStatus

    imsStatus = Prompt_ListValue("NR_IMSStatus", "Select IMSStatus (required).", "New Component (" & compId & ")", imsDefault)
    If Len(imsStatus) = 0 Then
        abortedReason = "IMSStatus not selected."
        GoTo FailRollback
    End If
    SetByHeader loComps, lr, "IMSStatus", imsStatus

    ' Required numeric fields
    moq1 = Prompt_Long("Enter MOQ1 (required).", "New Component (" & compId & ")", DEFAULT_MOQ1, 1, 1000000)
    If moq1 = -1 Then
        abortedReason = "MOQ1 not provided."
        GoTo FailRollback
    End If
    SetByHeader loComps, lr, "MOQ1", moq1

    costMOQ1 = Prompt_Double("Enter CostPerUOMMOQ1 (required).", "New Component (" & compId & ")", DEFAULT_COST_MOQ1, 0, 1000000000#)
    If costMOQ1 < 0 Then
        abortedReason = "CostPerUOMMOQ1 not provided."
        GoTo FailRollback
    End If
    SetByHeader loComps, lr, "CostPerUOMMOQ1", costMOQ1

    ' Optional: ComponentLT
    If ColumnExists(loComps, "ComponentLT") Then
        Dim ltVal As Long
        ltVal = Prompt_Long("Enter ComponentLT (days).", "New Component (" & compId & ")", CLng(val(CStr(pickDfltLT))), 0, 3650)
        If ltVal = -1 Then
            abortedReason = "ComponentLT not provided."
            GoTo FailRollback
        End If
        SetByHeader loComps, lr, "ComponentLT", ltVal
    End If

    ' IsBuildable (computed)
    Dim buildable As Boolean
    buildable = True
    If Len(Trim$(ourPN)) = 0 Then buildable = False
    If Len(Trim$(ourRev)) = 0 Then buildable = False
    If Len(Trim$(desc)) = 0 Then buildable = False
    If Len(Trim$(pickId)) = 0 Then buildable = False
    If Len(Trim$(pickName)) = 0 Then buildable = False
    If Len(Trim$(uom)) = 0 Then buildable = False
    If Len(Trim$(revStatus)) = 0 Then buildable = False
    If Len(Trim$(imsStatus)) = 0 Then buildable = False
    If moq1 < 1 Then buildable = False
    If costMOQ1 < 0 Then buildable = False

    SetByHeader loComps, lr, "IsBuildable", buildable

    createdOk = True
    MsgBox "Component created: " & compId & vbCrLf & _
           "Supplier: " & pickName & " [" & pickId & "]" & vbCrLf & _
           "IsBuildable: " & IIf(buildable, "TRUE", "FALSE"), vbInformation, "New Component"
    Exit Sub

FailRollback:
    On Error Resume Next
    If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    GoTo Aborted

Aborted:
    If Not createdOk Then
        If Len(Trim$(abortedReason)) > 0 Then
            MsgBox "No new component created." & vbCrLf & "Reason: " & abortedReason, vbInformation, "New Component"
        Else
            MsgBox "No new component created.", vbInformation, "New Component"
        End If
    End If
    Exit Sub

EH:
    On Error Resume Next
    If Not lr Is Nothing Then lr.Delete
    On Error GoTo 0
    MsgBox "No new component created." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New Component"
End Sub

'==========================
' Schema DefaultValue lookup (TBL_SCHEMA)
'==========================
Private Function GetSchemaDefaultValue(ByVal tabName As String, ByVal tableName As String, ByVal columnHeader As String) As String
    Const SCHEMA_TABLE As String = "TBL_SCHEMA"
    Const H_TAB As String = "TAB_NAME"
    Const H_TBL As String = "TABLE_NAME"
    Const H_COL As String = "COLUMN_HEADER"
    Const H_DEF As String = "DefaultValue"

    Dim lo As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long
    Dim idxTab As Long, idxTbl As Long, idxCol As Long, idxDef As Long
    Dim arr As Variant

    GetSchemaDefaultValue = vbNullString

    Set wb = ThisWorkbook
    Set lo = Nothing

    ' Find schema table anywhere
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(SCHEMA_TABLE)
        On Error GoTo 0
        If Not lo Is Nothing Then Exit For
    Next ws

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxTab = GetColIndex(lo, H_TAB)
    idxTbl = GetColIndex(lo, H_TBL)
    idxCol = GetColIndex(lo, H_COL)
    idxDef = GetColIndex(lo, H_DEF)
    If idxTab = 0 Or idxTbl = 0 Or idxCol = 0 Or idxDef = 0 Then Exit Function

    arr = lo.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(r, idxTab))), tabName, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(arr(r, idxTbl))), tableName, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(arr(r, idxCol))), columnHeader, vbTextCompare) = 0 Then

            GetSchemaDefaultValue = Trim$(CStr(arr(r, idxDef)))
            Exit Function
        End If
    Next r
End Function

'==========================
' Gate wrapper (safe)
'==========================
Private Function GateReady_Safe(Optional ByVal showUserMessage As Boolean = True) As Boolean
    On Error GoTo EH
    GateReady_Safe = M_Core_Gate.Gate_Ready(showUserMessage)
    Exit Function
EH:
    MsgBox "GateReady_Safe failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New Component"
    GateReady_Safe = False
End Function

'==========================
' REQUIRED TEXT PROMPT
'==========================
Private Function Prompt_RequiredText(ByVal prompt As String, ByVal title As String, ByVal defaultValue As String) As String
    Dim resp As String
Retry:
    resp = InputBox(prompt & vbCrLf & "(Leave blank and click Cancel to abort.)", title, defaultValue)
    resp = Trim$(resp)

    If Len(resp) = 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("This field is required." & vbCrLf & vbCrLf & _
                     "Yes = Try again" & vbCrLf & _
                     "No  = Cancel creation (rollback)", _
                     vbYesNo + vbExclamation, title)
        If ans = vbYes Then GoTo Retry
        Prompt_RequiredText = vbNullString
        Exit Function
    End If

    Prompt_RequiredText = resp
End Function

'==========================
' LIST PROMPTS (Named Range)
'==========================
Private Function Prompt_ListValue(ByVal namedRange As String, ByVal prompt As String, ByVal title As String, ByVal defaultValue As String) As String
    Dim arr As Variant
    Dim choices() As String
    Dim i As Long, n As Long
    Dim menu As String
    Dim resp As String
    Dim idx As Long

    Prompt_ListValue = vbNullString

    arr = GetNamedRangeValues(namedRange)
    n = UBound(arr, 1)

    If n <= 0 Then
        MsgBox "Named range '" & namedRange & "' has no values.", vbExclamation, title
        Exit Function
    End If

    ReDim choices(1 To n)
    For i = 1 To n
        choices(i) = Trim$(CStr(arr(i, 1)))
    Next i

Retry:
    menu = prompt & vbCrLf & vbCrLf
    If Len(defaultValue) > 0 Then
        menu = menu & "Default: " & defaultValue & vbCrLf & vbCrLf
    End If
    menu = menu & "Choose by number, or type an exact value from the list:" & vbCrLf & vbCrLf

    For i = 1 To n
        menu = menu & CStr(i) & ") " & choices(i) & vbCrLf
        If i >= 30 Then
            menu = menu & "...(list truncated; extend UI later if needed)" & vbCrLf
            Exit For
        End If
    Next i

    resp = InputBox(menu, title, defaultValue)
    resp = Trim$(resp)

    If Len(resp) = 0 Then
        If Len(defaultValue) > 0 Then
            Prompt_ListValue = defaultValue
        Else
            Prompt_ListValue = vbNullString
        End If
        Exit Function
    End If

    If IsNumeric(resp) Then
        idx = CLng(resp)
        If idx >= 1 And idx <= n Then
            If Len(choices(idx)) > 0 Then
                Prompt_ListValue = choices(idx)
                Exit Function
            End If
        End If
        MsgBox "Invalid selection number. Try again.", vbExclamation, title
        GoTo Retry
    End If

    For i = 1 To n
        If StrComp(choices(i), resp, vbTextCompare) = 0 Then
            Prompt_ListValue = choices(i)
            Exit Function
        End If
    Next i

    MsgBox "Value not found in '" & namedRange & "'. Please select from the list.", vbExclamation, title
    GoTo Retry
End Function

Private Function Prompt_Long(ByVal prompt As String, ByVal title As String, ByVal defaultValue As Long, ByVal minValue As Long, ByVal maxValue As Long) As Long
    Dim resp As String
    Dim v As Double

Retry:
    resp = Trim$(InputBox(prompt & vbCrLf & "(Min=" & CStr(minValue) & ", Max=" & CStr(maxValue) & ")", title, CStr(defaultValue)))

    If Len(resp) = 0 Then
        Prompt_Long = defaultValue
        Exit Function
    End If

    If Not IsNumeric(resp) Then
        MsgBox "Please enter a whole number.", vbExclamation, title
        GoTo Retry
    End If

    v = CDbl(resp)
    If v <> Fix(v) Then
        MsgBox "Please enter a whole number (no decimals).", vbExclamation, title
        GoTo Retry
    End If

    If v < minValue Or v > maxValue Then
        MsgBox "Value out of range. Must be between " & CStr(minValue) & " and " & CStr(maxValue) & ".", vbExclamation, title
        GoTo Retry
    End If

    Prompt_Long = CLng(v)
End Function

Private Function Prompt_Double(ByVal prompt As String, ByVal title As String, ByVal defaultValue As Double, ByVal minValue As Double, ByVal maxValue As Double) As Double
    Dim resp As String
    Dim v As Double

Retry:
    resp = Trim$(InputBox(prompt & vbCrLf & "(Min=" & CStr(minValue) & ", Max=" & CStr(maxValue) & ")", title, CStr(defaultValue)))

    If Len(resp) = 0 Then
        Prompt_Double = defaultValue
        Exit Function
    End If

    If Not IsNumeric(resp) Then
        MsgBox "Please enter a number.", vbExclamation, title
        GoTo Retry
    End If

    v = CDbl(resp)
    If v < minValue Or v > maxValue Then
        MsgBox "Value out of range. Must be between " & CStr(minValue) & " and " & CStr(maxValue) & ".", vbExclamation, title
        GoTo Retry
    End If

    Prompt_Double = v
End Function

Private Sub RequireNamedRange(ByVal namedRange As String)
    Dim nm As Name
    On Error GoTo EH
    Set nm = ThisWorkbook.names(namedRange)
    If nm Is Nothing Then Err.Raise vbObjectError + 5800, "RequireNamedRange", "Named range not found: " & namedRange
    Exit Sub
EH:
    MsgBox "RequireNamedRange failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New Component"
    Err.Raise vbObjectError + 5801, "RequireNamedRange", "Named range not found: " & namedRange
End Sub

Private Function GetNamedRangeValues(ByVal namedRange As String) As Variant
    Dim rng As Range
    Dim v As Variant
    Dim outArr() As Variant
    Dim r As Long, c As Long, n As Long

    Set rng = ThisWorkbook.names(namedRange).RefersToRange
    v = rng.value

    If IsArray(v) Then
        ReDim outArr(1 To (UBound(v, 1) * UBound(v, 2)), 1 To 1)
        n = 0
        For r = 1 To UBound(v, 1)
            For c = 1 To UBound(v, 2)
                If Len(Trim$(CStr(v(r, c)))) > 0 Then
                    n = n + 1
                    outArr(n, 1) = v(r, c)
                End If
            Next c
        Next r
        If n = 0 Then
            ReDim outArr(1 To 0, 1 To 1)
        ElseIf n < UBound(outArr, 1) Then
            ReDim Preserve outArr(1 To n, 1 To 1)
        End If
        GetNamedRangeValues = outArr
    Else
        ReDim outArr(1 To 1, 1 To 1)
        outArr(1, 1) = v
        GetNamedRangeValues = outArr
    End If
End Function

Private Function GetUserNameSafe() As String
    Dim u As String
    u = Trim$(Environ$("Username"))
    If Len(u) = 0 Then u = Application.userName
    If Len(Trim$(u)) = 0 Then u = "UNKNOWN"
    GetUserNameSafe = u
End Function

'==========================
' SUPPLIER PICKER (IMPROVED)
'==========================
Private Function SupplierPick_ByName(ByVal loSupp As ListObject, ByRef supplierId As String, ByRef supplierName As String, ByRef supplierDefaultLT As Variant) As Boolean
    Const title As String = "Pick Supplier"
    Const MAX_SHOW As Long = 25
    Const SAMPLE_SHOW As Long = 12

    Dim idxId As Long, idxName As Long, idxLT As Long
    Dim arrId As Variant, arrName As Variant, arrLT As Variant

    Dim term As String, termNorm As String
    Dim i As Long, hitCount As Long
    Dim hitsIdx() As Long

    Dim menu As String
    Dim choiceText As String
    Dim choiceN As Long

    SupplierPick_ByName = False
    supplierId = vbNullString
    supplierName = vbNullString
    supplierDefaultLT = vbNullString

    If loSupp Is Nothing Then
        MsgBox "Suppliers table reference is Nothing.", vbExclamation, title
        Exit Function
    End If

    If loSupp.DataBodyRange Is Nothing Then
        MsgBox "Suppliers table '" & loSupp.Name & "' has no rows.", vbExclamation, title
        Exit Function
    End If

    idxId = GetColIndex(loSupp, "SupplierID")
    idxName = GetColIndex(loSupp, "SupplierName")
    idxLT = GetColIndex(loSupp, "SupplierDefaultLT")

    If idxId = 0 Or idxName = 0 Or idxLT = 0 Then
        MsgBox "Supplier picker cannot run because required headers were not found." & vbCrLf & _
               "Expected: SupplierID, SupplierName, SupplierDefaultLT", vbExclamation, title
        Exit Function
    End If

    arrId = loSupp.ListColumns(idxId).DataBodyRange.value
    arrName = loSupp.ListColumns(idxName).DataBodyRange.value
    arrLT = loSupp.ListColumns(idxLT).DataBodyRange.value

RetrySearch:
    term = InputBox( _
        "Supplier search:" & vbCrLf & _
        "- Type part of the supplier name (e.g., dig, digi key, b&b, thread)." & vbCrLf & _
        "- Leave blank and click OK to cancel.", _
        title)

    term = Trim$(term)
    If Len(term) = 0 Then Exit Function

    termNorm = NormalizeForMatch(term)

    hitCount = 0
    ReDim hitsIdx(1 To 1)

    For i = 1 To UBound(arrName, 1)
        Dim candidate As String, candNorm As String
        candidate = CStr(arrName(i, 1))
        candNorm = NormalizeForMatch(candidate)

        If SupplierMatch(candNorm, termNorm) Then
            hitCount = hitCount + 1
            If hitCount = 1 Then
                hitsIdx(1) = i
            Else
                ReDim Preserve hitsIdx(1 To hitCount)
                hitsIdx(hitCount) = i
            End If
        End If
    Next i

    If hitCount = 0 Then
        Dim msg As String, nShow As Long
        nShow = WorksheetFunction.Min(SAMPLE_SHOW, UBound(arrName, 1))

        msg = "No suppliers matched: '" & term & "'" & vbCrLf & vbCrLf & _
              "Try fewer characters, different ordering, or omit punctuation." & vbCrLf & vbCrLf & _
              "Sample SupplierName values (first " & CStr(nShow) & "):" & vbCrLf

        For i = 1 To nShow
            msg = msg & "  - " & CStr(arrName(i, 1)) & vbCrLf
        Next i

        MsgBox msg, vbExclamation, title
        GoTo RetrySearch
    End If

    If hitCount = 1 Then
        i = hitsIdx(1)
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Found 1 match:" & vbCrLf & _
                     CStr(arrName(i, 1)) & "  [" & CStr(arrId(i, 1)) & "]" & vbCrLf & vbCrLf & _
                     "Yes = Use this supplier" & vbCrLf & _
                     "No  = Re-search", _
                     vbYesNo + vbQuestion, title)
        If ans = vbNo Then GoTo RetrySearch

        supplierId = CStr(arrId(i, 1))
        supplierName = CStr(arrName(i, 1))
        supplierDefaultLT = arrLT(i, 1)
        SupplierPick_ByName = True
        Exit Function
    End If

RetryPick:
    menu = "Multiple suppliers matched '" & term & "'. Choose a number:" & vbCrLf & vbCrLf

    For i = 1 To hitCount
        menu = menu & CStr(i) & ") " & CStr(arrName(hitsIdx(i), 1)) & "  [" & CStr(arrId(hitsIdx(i), 1)) & "]" & vbCrLf
        If i >= MAX_SHOW Then
            menu = menu & vbCrLf & "(Showing first " & CStr(MAX_SHOW) & ". Refine your search to narrow results.)"
            Exit For
        End If
    Next i

    choiceText = InputBox(menu & vbCrLf & "Enter a number (or leave blank to re-search).", title)
    choiceText = Trim$(choiceText)

    If Len(choiceText) = 0 Then GoTo RetrySearch

    If Not IsNumeric(choiceText) Then
        MsgBox "Please enter a valid number from the list.", vbExclamation, title
        GoTo RetryPick
    End If

    choiceN = CLng(choiceText)
    If choiceN < 1 Or choiceN > hitCount Then
        MsgBox "Choice out of range. Pick a number shown in the list.", vbExclamation, title
        GoTo RetryPick
    End If

    i = hitsIdx(choiceN)
    supplierId = CStr(arrId(i, 1))
    supplierName = CStr(arrName(i, 1))
    supplierDefaultLT = arrLT(i, 1)
    SupplierPick_ByName = True
End Function

Private Function SupplierMatch(ByVal candNorm As String, ByVal termNorm As String) As Boolean
    Dim cNoSp As String, tNoSp As String
    Dim tokens() As String
    Dim i As Long
    Dim tok As String

    SupplierMatch = False
    If Len(termNorm) = 0 Then Exit Function

    If InStr(1, candNorm, termNorm, vbTextCompare) > 0 Then
        SupplierMatch = True
        Exit Function
    End If

    cNoSp = Replace(candNorm, " ", vbNullString)
    tNoSp = Replace(termNorm, " ", vbNullString)

    If Len(tNoSp) > 0 Then
        If InStr(1, cNoSp, tNoSp, vbTextCompare) > 0 Then
            SupplierMatch = True
            Exit Function
        End If
    End If

    tokens = Split(termNorm, " ")
    For i = LBound(tokens) To UBound(tokens)
        tok = Trim$(tokens(i))
        If Len(tok) > 0 Then
            If InStr(1, candNorm, tok, vbTextCompare) = 0 Then Exit Function
        End If
    Next i

    SupplierMatch = True
End Function

Private Function NormalizeForMatch(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, out As String

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

'==========================
' GENERIC TABLE HELPERS
'==========================
Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 5200, "RequireColumn", "Missing column '" & header & "' in table '" & lo.Name & "'."
    End If
End Sub

Private Function ColumnExists(ByVal lo As ListObject, ByVal header As String) As Boolean
    ColumnExists = (GetColIndex(lo, header) > 0)
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Err.Raise vbObjectError + 5201, "SetByHeader", "Missing column '" & header & "' in table '" & lo.Name & "'."
    lr.Range.Cells(1, idx).value = v
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

Private Function ValueExistsInColumn(ByVal lo As ListObject, ByVal header As String, ByVal valueText As String) As Boolean
    Dim idx As Long, rng As Range
    ValueExistsInColumn = False
    idx = GetColIndex(lo, header)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set rng = lo.ListColumns(idx).DataBodyRange
    ValueExistsInColumn = (Application.WorksheetFunction.CountIf(rng, valueText) > 0)
End Function

Private Function GenerateNextId(ByVal lo As ListObject, ByVal header As String, ByVal prefix As String, ByVal padDigits As Long) As String
    Dim idx As Long, maxN As Long
    Dim arr As Variant
    Dim i As Long, s As String, n As Long

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

Private Function PNRevComboExists(ByVal lo As ListObject, ByVal ourPN As String, ByVal ourRev As String) As Boolean
    Dim idxPn As Long, idxRev As Long
    Dim arrPn As Variant, arrRev As Variant
    Dim i As Long

    PNRevComboExists = False
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxPn = GetColIndex(lo, "OurPN")
    idxRev = GetColIndex(lo, "OurRev")
    If idxPn = 0 Or idxRev = 0 Then Exit Function

    arrPn = lo.ListColumns(idxPn).DataBodyRange.value
    arrRev = lo.ListColumns(idxRev).DataBodyRange.value

    For i = 1 To UBound(arrPn, 1)
        If StrComp(Trim$(CStr(arrPn(i, 1))), Trim$(ourPN), vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(arrRev(i, 1))), Trim$(ourRev), vbTextCompare) = 0 Then
            PNRevComboExists = True
            Exit Function
        End If
    Next i
    
    
End Function


Private Sub SortTable_ByColumn(ByVal lo As ListObject, ByVal header As String, Optional ByVal ascending As Boolean = True)
    Const PROC_NAME As String = "M_Data_Comps_Entry.SortTable_ByColumn"

    Dim ws As Worksheet
    Dim sortKey As Range

    On Error GoTo EH

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set ws = lo.Parent

    ' Ensure the column exists
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 5900, PROC_NAME, "Sort column not found: '" & header & "' in table '" & lo.Name & "'."
    End If

    Set sortKey = lo.ListColumns(header).DataBodyRange

    With lo.Sort
        .SortFields.Clear
        .SortFields.Add key:=sortKey, SortOn:=xlSortOnValues, Order:=IIf(ascending, xlAscending, xlDescending), DataOption:=xlSortNormal
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

CleanExit:
    Exit Sub

EH:
    ' Fail "soft" (sorting should not block record creation)
    ' If you prefer to log, replace with M_Core_Logging.LogEvent(...) here.
    MsgBox "SortTable_ByColumn failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, PROC_NAME
    Resume CleanExit
End Sub


