Attribute VB_Name = "M_Data_Suppliers_Entry"
Option Explicit

'===============================================================================
' Module: M_Data_Suppliers_Entry
'
' Purpose:
'   Create new Supplier records using schema-driven requirements from SCHEMA.TBL_SCHEMA.
'   Avoids hard-coded validation rules so changes are made by editing schema rows.
'
' Inputs (Tabs/Tables/Headers):
'   - Suppliers!TBL_SUPPLIERS with columns:
'       SupplierID, SupplierStatus, SupplierName, ASLStatus, SupplierContact, SupplierDefaultLT
'       (plus audit columns stamped by this procedure if present in constants)
'   - SCHEMA!TBL_SCHEMA with headers including:
'       TAB_NAME, TABLE_NAME, COLUMN_HEADER, IsRequired, DefaultValue, UserEditable,
'       DataType, Unique, HelperName, MinValue, MaxValue, MaxLength, Pattern
'
' Outputs / Side effects:
'   - Appends a new row to Suppliers.TBL_SUPPLIERS
'   - Stamps audit fields if columns exist (CreatedAt/By, UpdatedAt/By)
'   - Logs key actions to Log.TBL_LOG via M_Core_Logging
'
' Preconditions / Postconditions:
'   - Gate should be enforced by UI layer prior to calling this worker
'   - Schema rows for the 5 supplier fields exist and are unique
'
' Errors & Guards:
'   - Fails fast if Suppliers table or required schema rows are missing
'   - Blocks create with clear messages on schema rule violations
'   - Deletes the newly inserted row if validation fails after insert
'
' Version: v3.5.3
' Author: Keith + GPT
' Date: 2025-12-27
'===============================================================================

Private Const MODULE_VERSION As String = "3.5.3"

Public Sub NewSupplier()
    Const PROC_NAME As String = "M_Data_Suppliers_Entry.NewSupplier"

    ' SupplierID format per your standard: SUP-#### (max+1)
    Const SUPPLIER_ID_PREFIX As String = "SUP-"
    Const SUPPLIER_ID_PAD As Long = 4

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow

    Dim supplierId As String
    Dim userId As String
    Dim nowStamp As Date

    On Error GoTo EH

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(SH_SUPPLIERS)
    Set lo = ws.ListObjects(TBL_SUPPLIERS)

    '--- Validate table has the columns we will write (fail fast)
    RequireColumn lo, "SupplierID"
    RequireColumn lo, "SupplierName"
    RequireColumn lo, "SupplierStatus"
    RequireColumn lo, "ASLStatus"
    RequireColumn lo, "SupplierDefaultLT"
    ' Optional but commonly present:
    ' SupplierContact

    '--- Load schema specs (your pasted rows must exist and be unique)
    Dim sSupplierID As Object
    Dim sSupplierName As Object
    Dim sSupplierStatus As Object
    Dim sASLStatus As Object
    Dim sSupplierLT As Object

    Set sSupplierID = SchemaRow_Get("Suppliers", "TBL_SUPPLIERS", "SupplierID")
    Set sSupplierName = SchemaRow_Get("Suppliers", "TBL_SUPPLIERS", "SupplierName")
    Set sSupplierStatus = SchemaRow_Get("Suppliers", "TBL_SUPPLIERS", "SupplierStatus")
    Set sASLStatus = SchemaRow_Get("Suppliers", "TBL_SUPPLIERS", "ASLStatus")
    Set sSupplierLT = SchemaRow_Get("Suppliers", "TBL_SUPPLIERS", "SupplierDefaultLT")

    '--- SupplierID must be non-editable per schema (UserEditable = N)
    If Schema_IsUserEditable(sSupplierID) Then
        Err.Raise vbObjectError + 801, PROC_NAME, "Schema violation: SupplierID should be non-editable (UserEditable=N)."
    End If

    '--- Generate SupplierID (max+1) and enforce uniqueness
    supplierId = Supplier_GenerateNextId(lo, SUPPLIER_ID_PREFIX, SUPPLIER_ID_PAD)
    If Len(supplierId) = 0 Then Err.Raise vbObjectError + 802, PROC_NAME, "Failed to generate SupplierID."
    If ValueExistsInColumn(lo, "SupplierID", supplierId) Then Err.Raise vbObjectError + 803, PROC_NAME, "Generated SupplierID already exists: " & supplierId

    '--- Create the row now, but be prepared to delete if validation fails
    Set lr = lo.ListRows.Add
    SetByHeader lo, lr, "SupplierID", supplierId

    '--- Prompt + validate fields per schema
    Dim v As Variant

    ' SupplierName (Required + Unique per your row)
    v = Prompt_Validate_SchemaValue(sSupplierName, "Supplier Name", supplierId, "New Supplier")
    If Len(Trim$(CStr(v))) = 0 Then GoTo FailAndRollback
    If Schema_IsUnique(sSupplierName) Then
        If ValueExistsInColumn(lo, "SupplierName", CStr(v)) Then
            MsgBox "SupplierName must be unique. '" & CStr(v) & "' already exists.", vbExclamation, "New Supplier"
            GoTo FailAndRollback
        End If
    End If
    SetByHeader lo, lr, "SupplierName", v

    ' SupplierStatus (Required + HelperName list)
    v = Prompt_Validate_SchemaValue(sSupplierStatus, "Supplier Status", supplierId, "New Supplier")
    If Len(Trim$(CStr(v))) = 0 Then GoTo FailAndRollback
    SetByHeader lo, lr, "SupplierStatus", v

    ' ASLStatus (Required + HelperName list)
    v = Prompt_Validate_SchemaValue(sASLStatus, "ASL Status", supplierId, "New Supplier")
    If Len(Trim$(CStr(v))) = 0 Then GoTo FailAndRollback
    SetByHeader lo, lr, "ASLStatus", v

    ' SupplierContact (optional; not in your “four rows”, but exists in schema—keep it simple)
    If ColumnExists(lo, "SupplierContact") Then
        Dim contactName As String
        contactName = Trim$(InputBox("Supplier Contact (optional).", "New Supplier (" & supplierId & ")"))
        SetByHeader lo, lr, "SupplierContact", contactName
    End If

    ' SupplierDefaultLT (Required Integer; Min/Max only if present in schema)
    v = Prompt_Validate_SchemaValue(sSupplierLT, "Default Lead Time (days)", supplierId, "New Supplier")
    If Len(Trim$(CStr(v))) = 0 Then GoTo FailAndRollback
    SetByHeader lo, lr, "SupplierDefaultLT", v

    '--- Audit stamps if those columns exist
    userId = SafeUserId()
    nowStamp = Now
    StampAuditIfPresent lo, lr, userId, nowStamp

    M_Core_Logging.LogInfo PROC_NAME, "Created Supplier", "SupplierID=" & supplierId & "; ModuleVersion=" & MODULE_VERSION
    Exit Sub

FailAndRollback:
    On Error Resume Next
    lr.Delete
    On Error GoTo EH
    M_Core_Logging.LogWarn PROC_NAME, "Create Supplier blocked; row rolled back", "SupplierID=" & supplierId & "; ModuleVersion=" & MODULE_VERSION
    Exit Sub

EH:
    M_Core_Logging.LogError PROC_NAME, "Error creating supplier", "Err " & Err.Number & ": " & Err.Description, Err.Number
    MsgBox "New Supplier failed. See Log sheet for details." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New Supplier"
End Sub

'===============================================================================
' Schema-driven prompt + validate
'===============================================================================

Private Function Prompt_Validate_SchemaValue(ByVal sRow As Object, ByVal label As String, ByVal supplierId As String, ByVal formTitle As String) As Variant
    Dim isReq As Boolean
    Dim dt As String
    Dim defVal As String
    Dim helperName As String
    Dim minV As Variant, maxV As Variant
    Dim hasMin As Boolean, hasMax As Boolean

    Dim raw As String
    Dim outVal As Variant

    isReq = Schema_IsRequired(sRow)
    dt = UCase$(Schema_Str(sRow, "DataType"))
    defVal = Schema_Str(sRow, "DefaultValue")
    helperName = Schema_Str(sRow, "HelperName")

    hasMin = Schema_HasValue(sRow, "MinValue")
    hasMax = Schema_HasValue(sRow, "MaxValue")
    If hasMin Then minV = sRow("MinValue")
    If hasMax Then maxV = sRow("MaxValue")

    raw = InputBox(label & IIf(isReq, " (required).", " (optional)."), formTitle & " (" & supplierId & ")", defVal)

    raw = Trim$(raw)

    If isReq And Len(raw) = 0 Then
        MsgBox label & " is required.", vbExclamation, formTitle
        Prompt_Validate_SchemaValue = vbNullString
        Exit Function
    End If

    If Len(raw) = 0 Then
        Prompt_Validate_SchemaValue = vbNullString
        Exit Function
    End If

    Select Case dt
        Case "TEXT", "CODE"
            outVal = raw

        Case "INTEGER"
            If Not IsNumeric(raw) Then
                MsgBox label & " must be an integer.", vbExclamation, formTitle
                Prompt_Validate_SchemaValue = vbNullString
                Exit Function
            End If
            outVal = CLng(raw)

        Case "DECIMAL", "NUMBER", "DOUBLE"
            If Not IsNumeric(raw) Then
                MsgBox label & " must be a number.", vbExclamation, formTitle
                Prompt_Validate_SchemaValue = vbNullString
                Exit Function
            End If
            outVal = CDbl(raw)

        Case "DATE"
            If Not IsDate(raw) Then
                MsgBox label & " must be a date.", vbExclamation, formTitle
                Prompt_Validate_SchemaValue = vbNullString
                Exit Function
            End If
            outVal = CDate(raw)

        Case Else
            ' Unknown DataType token ? treat as text (boring, non-blocking)
            outVal = raw
    End Select

    ' HelperName allowed values (named range)
    If Len(helperName) > 0 Then
        If Not ValueInNamedRange(ThisWorkbook, helperName, CStr(outVal)) Then
            MsgBox label & " must be one of the allowed values in " & helperName & ".", vbExclamation, formTitle
            Prompt_Validate_SchemaValue = vbNullString
            Exit Function
        End If
    End If

    ' Min/Max bounds only if present in schema
    If hasMin Then
        If Not CompareGE(outVal, minV) Then
            MsgBox label & " must be >= " & CStr(minV) & ".", vbExclamation, formTitle
            Prompt_Validate_SchemaValue = vbNullString
            Exit Function
        End If
    End If
    If hasMax Then
        If Not CompareLE(outVal, maxV) Then
            MsgBox label & " must be <= " & CStr(maxV) & ".", vbExclamation, formTitle
            Prompt_Validate_SchemaValue = vbNullString
            Exit Function
        End If
    End If

    Prompt_Validate_SchemaValue = outVal
End Function

'===============================================================================
' Schema table access (private; no new module)
'===============================================================================

Private Function SchemaRow_Get(ByVal tabName As String, ByVal tableName As String, ByVal columnHeader As String) As Object
    Const PROC_NAME As String = "SchemaRow_Get"

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As ListRow
    Dim dic As Object
    Dim idxTab As Long, idxTable As Long, idxCol As Long
    Dim matchCount As Long

    Set ws = ThisWorkbook.Worksheets("SCHEMA")
    Set lo = ws.ListObjects("TBL_SCHEMA")

    idxTab = GetColIndexOrRaise(lo, "TAB_NAME")
    idxTable = GetColIndexOrRaise(lo, "TABLE_NAME")
    idxCol = GetColIndexOrRaise(lo, "COLUMN_HEADER")

    matchCount = 0
    For Each r In lo.ListRows
        If StrComp(Trim$(CStr(r.Range.Cells(1, idxTab).value)), tabName, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(r.Range.Cells(1, idxTable).value)), tableName, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(r.Range.Cells(1, idxCol).value)), columnHeader, vbTextCompare) = 0 Then

            matchCount = matchCount + 1

            Set dic = CreateObject("Scripting.Dictionary")
            LoadSchemaRowToDictionary lo, r, dic
            Set SchemaRow_Get = dic
        End If
    Next r

    If matchCount = 0 Then
        Err.Raise vbObjectError + 820, PROC_NAME, "No schema row found for " & tabName & "." & tableName & "." & columnHeader
    End If

    If matchCount > 1 Then
        Err.Raise vbObjectError + 821, PROC_NAME, "Multiple schema rows found for " & tabName & "." & tableName & "." & columnHeader
    End If
End Function

Private Sub LoadSchemaRowToDictionary(ByVal lo As ListObject, ByVal r As ListRow, ByVal dic As Object)
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        dic(lc.Name) = r.Range.Cells(1, lc.Index).value
    Next lc
End Sub

Private Function Schema_IsRequired(ByVal sRow As Object) As Boolean
    Schema_IsRequired = ToBool(sRow("IsRequired"), False)
End Function

Private Function Schema_IsUserEditable(ByVal sRow As Object) As Boolean
    Schema_IsUserEditable = ToBool(sRow("UserEditable"), True)
End Function

Private Function Schema_IsUnique(ByVal sRow As Object) As Boolean
    Schema_IsUnique = ToBool(sRow("Unique"), False)
End Function

Private Function Schema_Str(ByVal sRow As Object, ByVal key As String) As String
    If sRow.Exists(key) Then
        Schema_Str = Trim$(CStr(sRow(key)))
    Else
        Schema_Str = vbNullString
    End If
End Function

Private Function Schema_HasValue(ByVal sRow As Object, ByVal key As String) As Boolean
    If Not sRow.Exists(key) Then
        Schema_HasValue = False
        Exit Function
    End If
    Schema_HasValue = Not IsBlankOrError(sRow(key))
End Function

'===============================================================================
' SupplierID generation (MAX trailing number + 1)
'===============================================================================

Private Function Supplier_GenerateNextId(ByVal loSuppliers As ListObject, ByVal prefix As String, ByVal padDigits As Long) As String
    Dim maxN As Long
    Dim arr As Variant
    Dim i As Long
    Dim s As String
    Dim n As Long
    Dim idx As Long

    maxN = 0
    idx = GetColIndex(loSuppliers, "SupplierID")
    If idx = 0 Then Exit Function

    If loSuppliers.DataBodyRange Is Nothing Then
        Supplier_GenerateNextId = prefix & Right$(String$(padDigits, "0") & "1", padDigits)
        Exit Function
    End If

    arr = loSuppliers.ListColumns(idx).DataBodyRange.value

    If IsArray(arr) Then
        For i = 1 To UBound(arr, 1)
            s = Trim$(CStr(arr(i, 1)))
            n = TrailingNumber(s)
            If n > maxN Then maxN = n
        Next i
    Else
        s = Trim$(CStr(arr))
        n = TrailingNumber(s)
        If n > maxN Then maxN = n
    End If

    Supplier_GenerateNextId = prefix & Right$(String$(padDigits, "0") & CStr(maxN + 1), padDigits)
End Function

Private Function TrailingNumber(ByVal s As String) As Long
    Dim i As Long
    Dim ch As String
    Dim digits As String

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

'===============================================================================
' Table + validation helpers
'===============================================================================

Private Sub RequireColumn(ByVal lo As ListObject, ByVal header As String)
    If GetColIndex(lo, header) = 0 Then
        Err.Raise vbObjectError + 840, "RequireColumn", "Missing column '" & header & "' in table '" & lo.Name & "'."
    End If
End Sub

Private Function ColumnExists(ByVal lo As ListObject, ByVal header As String) As Boolean
    ColumnExists = (GetColIndex(lo, header) > 0)
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal lr As ListRow, ByVal header As String, ByVal v As Variant)
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Err.Raise vbObjectError + 841, "SetByHeader", "Missing column '" & header & "' in " & lo.Name
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

Private Function GetColIndexOrRaise(ByVal lo As ListObject, ByVal header As String) As Long
    Dim idx As Long
    idx = GetColIndex(lo, header)
    If idx = 0 Then Err.Raise vbObjectError + 842, "GetColIndexOrRaise", "Schema missing header: " & header
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

Private Function ValueInNamedRange(ByVal wb As Workbook, ByVal rangeName As String, ByVal valueText As String) As Boolean
    Dim rng As Range
    Dim c As Range

    On Error GoTo EH
    Set rng = wb.names(rangeName).RefersToRange

    ValueInNamedRange = False
    For Each c In rng.Cells
        If StrComp(Trim$(CStr(c.value)), Trim$(valueText), vbTextCompare) = 0 Then
            ValueInNamedRange = True
            Exit Function
        End If
    Next c
    Exit Function
EH:
    MsgBox "ValueInNamedRange failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Supplier Entry"
    ValueInNamedRange = False
End Function

Private Function CompareGE(ByVal a As Variant, ByVal b As Variant) As Boolean
    On Error GoTo EH
    CompareGE = (a >= b)
    Exit Function
EH:
    MsgBox "CompareGE failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Supplier Entry"
    CompareGE = False
End Function

Private Function CompareLE(ByVal a As Variant, ByVal b As Variant) As Boolean
    On Error GoTo EH
    CompareLE = (a <= b)
    Exit Function
EH:
    MsgBox "CompareLE failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Supplier Entry"
    CompareLE = False
End Function

Private Function IsBlankOrError(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankOrError = True
    ElseIf IsEmpty(v) Then
        IsBlankOrError = True
    ElseIf VarType(v) = vbString And Len(Trim$(CStr(v))) = 0 Then
        IsBlankOrError = True
    Else
        IsBlankOrError = False
    End If
End Function

Private Function ToBool(ByVal v As Variant, ByVal defaultVal As Boolean) As Boolean
    Dim s As String
    If IsBlankOrError(v) Then
        ToBool = defaultVal
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        ToBool = CBool(v)
        Exit Function
    End If
    s = LCase$(Trim$(CStr(v)))
    Select Case s
        Case "y", "yes", "true", "1"
            ToBool = True
        Case "n", "no", "false", "0"
            ToBool = False
        Case Else
            ToBool = defaultVal
    End Select
End Function

Private Function SafeUserId() As String
    Dim u As String
    u = Trim$(Environ$("Username"))
    If Len(u) = 0 Then u = "UNKNOWN"
    SafeUserId = u
End Function

Private Sub StampAuditIfPresent(ByVal lo As ListObject, ByVal lr As ListRow, ByVal userId As String, ByVal ts As Date)
    ' Uses your existing constants (COL_CREATED_AT etc.). If those columns don’t exist, it silently skips.
    On Error Resume Next
    If ColumnExists(lo, COL_CREATED_AT) Then SetByHeader lo, lr, COL_CREATED_AT, ts
    If ColumnExists(lo, COL_CREATED_BY) Then SetByHeader lo, lr, COL_CREATED_BY, userId
    If ColumnExists(lo, COL_UPDATED_AT) Then SetByHeader lo, lr, COL_UPDATED_AT, ts
    If ColumnExists(lo, COL_UPDATED_BY) Then SetByHeader lo, lr, COL_UPDATED_BY, userId
    On Error GoTo 0
End Sub


