Attribute VB_Name = "M_Core_Utils"
Option Explicit
'*******************************************************************************
' Module:      M_Core_Utils
'
' Purpose:
'   Central helper utilities for the Tracker workbook. This module provides
'   reusable, non-business-specific functions that support safe workbook
'   interaction, ListObject access, dictionary-based joins, string handling,
'   and light user interaction (confirmation prompts).
'
' Inputs (Tabs/Tables/Headers):
'   - Indirectly works with any worksheet/ListObject passed in.
'   - No hard-coded sheet or table names; all references are via arguments.
'
' Outputs / Side effects:
'   - ConfirmProceed displays a MsgBox and returns a Boolean.
'   - SafeGetListObject returns a ListObject or Nothing without raising.
'   - SafeGetValue/SafeSetValue read/write cells in ListObjects safely.
'   - ListObjectToArray/ArrayToListObject move data between tables and arrays.
'   - BuildDictionaryByColumn returns a Scripting.Dictionary keyed by column
'     values (e.g., CompID, OurPN+OurRev).
'   - Utils_GenerateActivityId returns a unique-ish string for logging.
'
' Preconditions / Postconditions:
'   - Callers must pass valid Workbook/Worksheet/ListObject references.
'   - Functions attempt to handle missing tables/columns gracefully, logging
'     errors via M_Core_Logging.LogEvent where appropriate.
'   - No direct modification of workbook structure (no Add/Delete sheets).
'
' Errors & Guards:
'   - Public functions use structured error handling:
'       On Error GoTo EH
'       ...
'     and log via LogEvent on failure.
'   - Most functions fail-soft (return Nothing, 0, or defaultValue) rather
'     than raising unhandled errors.
'
' Version:     v0.1.0
' Author:      ChatGPT (assistant)
' Date:        2025-11-29
'
' @spec
'   Purpose: Provide reusable utilities for safe and maintainable VBA code,
'            particularly around ListObjects, values, and dictionary joins.
'   Inputs: Various Worksheet/ListObject/Workbook/value parameters as needed.
'   Outputs: Values, Lists, Dictionaries, and user confirmations.
'   Preconditions: M_Core_Logging and M_Core_Constants compiled and available.
'   Postconditions: No structural changes; minimal side effects beyond logging
'                   and any requested value writes.
'   Errors: Logged via LogEvent; functions return reasonable defaults.
'   Version: v0.1.0
'   Author: ChatGPT
'   Date: 2025-11-29
'*******************************************************************************

'===============================================================================
' Public API - User confirmation
'===============================================================================

Public Function ConfirmProceed( _
    ByVal prompt As String, _
    Optional ByVal TITLE As String = "Confirm Action", _
    Optional ByVal defaultNo As Boolean = True) As Boolean
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Display a Yes/No confirmation dialog before destructive or high-impact
    '   operations. Returns True if user clicks Yes, False otherwise.
    '
    ' Errors:
    '   Any error is logged and function returns False.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "ConfirmProceed"
    
    Dim buttons As VbMsgBoxStyle
    Dim defaultButton As VbMsgBoxStyle
    Dim answer As VbMsgBoxResult
    
    On Error GoTo EH
    
    buttons = vbYesNo + vbQuestion
    If defaultNo Then
        defaultButton = vbDefaultButton2
    Else
        defaultButton = vbDefaultButton1
    End If
    
    answer = MsgBox(prompt, buttons + defaultButton, TITLE)
    ConfirmProceed = (answer = vbYes)
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, "Error in ConfirmProceed", Err.Description, Err.Number
    ConfirmProceed = False
    Resume CleanExit
End Function

'===============================================================================
' Public API - Table and column helpers
'===============================================================================

Public Function SafeGetListObject( _
    ByVal ws As Worksheet, _
    ByVal tableName As String) As ListObject
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Safely retrieve a ListObject by name from a worksheet. Returns Nothing
    '   if the table is not found instead of raising an error.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "SafeGetListObject"
    
    On Error GoTo EH
    
    If ws Is Nothing Then
        GoTo CleanExit
    End If
    
    On Error Resume Next
    Set SafeGetListObject = ws.ListObjects(tableName)
    On Error GoTo EH
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error retrieving ListObject '" & tableName & "'", Err.Description, Err.Number
    Set SafeGetListObject = Nothing
    Resume CleanExit
End Function

Public Function IsTablePresent( _
    ByVal wb As Workbook, _
    ByVal sheetName As String, _
    ByVal tableName As String) As Boolean
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Check if a sheet and table exist in the workbook, returning True/False
    '   instead of raising an error.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "IsTablePresent"
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    On Error GoTo EH
    
    If wb Is Nothing Then GoTo CleanExit
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo EH
    
    If ws Is Nothing Then GoTo CleanExit
    
    Set lo = SafeGetListObject(ws, tableName)
    IsTablePresent = Not (lo Is Nothing)
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error checking table presence: " & sheetName & "." & tableName, _
             Err.Description, Err.Number
    IsTablePresent = False
    Resume CleanExit
End Function

Public Function IsColumnPresent( _
    ByVal lo As ListObject, _
    ByVal columnName As String) As Boolean
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Check if a ListObject contains a column with the specified header.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "IsColumnPresent"
    
    Dim lc As ListColumn
    
    On Error GoTo EH
    
    If lo Is Nothing Then GoTo CleanExit
    
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            IsColumnPresent = True
            Exit For
        End If
    Next lc
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error checking column presence: " & columnName, Err.Description, Err.Number
    IsColumnPresent = False
    Resume CleanExit
End Function

Public Function GetColumnIndexByName( _
    ByVal lo As ListObject, _
    ByVal columnName As String, _
    Optional ByVal caseSensitive As Boolean = False) As Long
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Return the index of a ListObject column by header name. Returns 0
    '   if the column is not found.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "GetColumnIndexByName"
    
    Dim lc As ListColumn
    Dim compareMode As VbCompareMethod
    
    On Error GoTo EH
    
    If lo Is Nothing Then GoTo CleanExit
    
    If caseSensitive Then
        compareMode = vbBinaryCompare
    Else
        compareMode = vbTextCompare
    End If
    
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, compareMode) = 0 Then
            GetColumnIndexByName = lc.Index
            Exit For
        End If
    Next lc
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error getting column index: " & columnName, Err.Description, Err.Number
    GetColumnIndexByName = 0
    Resume CleanExit
End Function

'===============================================================================
' Public API - Safe value getters/setters
'===============================================================================

Public Function SafeGetValue( _
    ByVal lo As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    Optional ByVal defaultValue As Variant) As Variant
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Safely read a value from a ListObject's DataBodyRange by row index and
    '   column header. Returns defaultValue if out of range or column missing.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "SafeGetValue"
    
    Dim colIndex As Long
    
    On Error GoTo EH
    
    SafeGetValue = defaultValue
    
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit
    If rowIndex < 1 Or rowIndex > lo.DataBodyRange.Rows.Count Then GoTo CleanExit
    
    colIndex = GetColumnIndexByName(lo, columnName)
    If colIndex = 0 Then GoTo CleanExit
    
    SafeGetValue = lo.DataBodyRange.Cells(rowIndex, colIndex).value
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error in SafeGetValue for column '" & columnName & "'", Err.Description, Err.Number
    SafeGetValue = defaultValue
    Resume CleanExit
End Function

Public Sub SafeSetValue( _
    ByVal lo As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    ByVal newValue As Variant)
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Safely write a value into a ListObject's DataBodyRange by row index
    '   and column header. Does nothing if row/column is invalid.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "SafeSetValue"
    
    Dim colIndex As Long
    
    On Error GoTo EH
    
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit
    If rowIndex < 1 Or rowIndex > lo.DataBodyRange.Rows.Count Then GoTo CleanExit
    
    colIndex = GetColumnIndexByName(lo, columnName)
    If colIndex = 0 Then GoTo CleanExit
    
    lo.DataBodyRange.Cells(rowIndex, colIndex).value = newValue
    
CleanExit:
    Exit Sub
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error in SafeSetValue for column '" & columnName & "'", Err.Description, Err.Number
    Resume CleanExit
End Sub

'===============================================================================
' Public API - Bulk array helpers
'===============================================================================

Public Function ListObjectToArray( _
    ByVal lo As ListObject) As Variant
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Return the DataBodyRange of a ListObject as a 2D Variant array.
    '   Returns an uninitialized Variant if there is no data.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "ListObjectToArray"
    
    On Error GoTo EH
    
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit
    
    ListObjectToArray = lo.DataBodyRange.value
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, "Error in ListObjectToArray", Err.Description, Err.Number
    Erase ListObjectToArray
    Resume CleanExit
End Function

Public Sub ArrayToListObject( _
    ByVal lo As ListObject, _
    ByVal data As Variant, _
    Optional ByVal clearExisting As Boolean = True)
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Write a 2D Variant array back into a ListObject.DataBodyRange. If
    '   clearExisting is True, existing rows are deleted and new rows added
    '   to match the array size.
    '
    ' Notes:
    '   - Expects data to be 2D, 1-based (typical Worksheet/Range arrays).
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "ArrayToListObject"
    
    Dim rowCount As Long
    Dim colCount As Long
    Dim currentRows As Long
    Dim i As Long
    
    On Error GoTo EH
    
    If lo Is Nothing Then GoTo CleanExit
    If IsEmpty(data) Then GoTo CleanExit
    
    rowCount = UBound(data, 1) - LBound(data, 1) + 1
    colCount = UBound(data, 2) - LBound(data, 2) + 1
    
    Application.EnableEvents = False
    
    If clearExisting Then
        ' Delete all existing rows
        If Not lo.DataBodyRange Is Nothing Then
            currentRows = lo.DataBodyRange.Rows.Count
            For i = currentRows To 1 Step -1
                lo.ListRows(i).Delete
            Next i
        End If
    End If
    
    ' Add rows to match array rowCount
    For i = 1 To rowCount
        lo.ListRows.Add
    Next i
    
    ' Now write the array into the DataBodyRange
    lo.DataBodyRange.Resize(rowCount, colCount).value = data
    
CleanExit:
    On Error Resume Next
    Application.EnableEvents = True
    Exit Sub
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, "Error in ArrayToListObject", Err.Description, Err.Number
    Resume CleanExit
End Sub

'===============================================================================
' Public API - Dictionary helpers
'===============================================================================

Public Function BuildDictionaryByColumn( _
    ByVal lo As ListObject, _
    ByVal keyColumnName As String, _
    Optional ByVal includeBlanks As Boolean = False) As Object
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Build a Scripting.Dictionary keyed by the values in keyColumnName
    '   for each row in the ListObject's DataBodyRange.
    '
    '   - Key:   Value of keyColumnName (Variant)
    '   - Item:  Long row index (1-based within DataBodyRange)
    '
    ' Notes:
    '   - Caller can then use the dictionary to perform fast joins:
    '       rowIndex = dic(key)
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "BuildDictionaryByColumn"
    
    Dim dic As Object
    Dim data As Variant
    Dim rowCount As Long
    Dim i As Long
    Dim colIndex As Long
    Dim keyValue As Variant
    
    On Error GoTo EH
    
    Set dic = CreateObject("Scripting.Dictionary")
    
    If lo Is Nothing Then GoTo CleanExit
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit
    
    colIndex = GetColumnIndexByName(lo, keyColumnName)
    If colIndex = 0 Then GoTo CleanExit
    
    data = lo.DataBodyRange.value
    rowCount = UBound(data, 1)
    
    For i = 1 To rowCount
        keyValue = data(i, colIndex)
        If (Not includeBlanks) And IsEmpty(keyValue) Then
            ' skip blank keys
        ElseIf Not dic.Exists(keyValue) Then
            dic.Add keyValue, i
        Else
            ' Duplicate keys: caller can decide how to handle.
            ' For now we leave first occurrence.
        End If
    Next i
    
    Set BuildDictionaryByColumn = dic
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error in BuildDictionaryByColumn for column '" & keyColumnName & "'", _
             Err.Description, Err.Number
    Set BuildDictionaryByColumn = Nothing
    Resume CleanExit
End Function

'===============================================================================
' Public API - String / ID helpers
'===============================================================================

Public Function NormalizeString( _
    ByVal value As String, _
    Optional ByVal toUpper As Boolean = False) As String
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Trim leading/trailing spaces and normalize case.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "NormalizeString"
    
    On Error GoTo EH
    
    value = Trim$(value)
    If toUpper Then
        NormalizeString = UCase$(value)
    Else
        NormalizeString = value
    End If
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, "Error in NormalizeString", Err.Description, Err.Number
    NormalizeString = value
    Resume CleanExit
End Function

Public Function Utils_GenerateActivityId( _
    ByVal procName As String) As String
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Generate a simple activity identifier string for logging or grouping
    '   related events. Not globally unique, but sufficient for workbook use.
    '
    ' Format:
    '   <procName>_yyyymmdd_hhnnss
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "Utils_GenerateActivityId"
    
    On Error GoTo EH
    
    Utils_GenerateActivityId = procName & "_" & Format$(Now, "yyyymmdd_hhnnss")
    
CleanExit:
    Exit Function
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error in Utils_GenerateActivityId", Err.Description, Err.Number
    Utils_GenerateActivityId = procName & "_ERROR"
    Resume CleanExit
End Function

'===============================================================================
' Test harness
'===============================================================================

Public Sub Test_Core_Utils()
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Light sanity checks for core utilities. This is not a full unit test
    '   suite, but a quick way to verify basic behavior:
    '     - Confirm sheet/table presence
    '     - Try building a dictionary on a known table/column
    '     - Read/write a value via SafeGetValue/SafeSetValue
    '
    ' Notes:
    '   - Update sheet/table/column names below to match a small, safe area
    '     of the workbook (e.g., TBL_SUPPLIERS on Suppliers sheet).
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "Test_Core_Utils"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dic As Object
    Dim sampleValue As Variant
    Dim testRow As Long
    
    On Error GoTo EH
    
    Set wb = ThisWorkbook
    ' Adjust these to something safe in your workbook
    Set ws = wb.Worksheets("Suppliers")
    Set lo = SafeGetListObject(ws, "TBL_SUPPLIERS")
    
    If lo Is Nothing Then
        MsgBox "Test_Core_Utils: TBL_SUPPLIERS not found on Suppliers sheet.", vbExclamation, PROC_NAME
        GoTo CleanExit
    End If
    
    ' Build dictionary on SupplierID
    Set dic = BuildDictionaryByColumn(lo, "SupplierID")
    If dic Is Nothing Then
        MsgBox "Test_Core_Utils: Failed to build dictionary on SupplierID.", vbExclamation, PROC_NAME
        GoTo CleanExit
    End If
    
    If dic.Count > 0 Then
        ' Grab first key and test SafeGetValue
        Dim firstKey As Variant
        firstKey = dic.Keys()(0)
        testRow = dic.Item(firstKey)
        
        sampleValue = SafeGetValue(lo, testRow, "SupplierName", "N/A")
        Debug.Print "Test_Core_Utils: SupplierID=" & firstKey & " Name=" & sampleValue
    End If
    
    MsgBox "Test_Core_Utils completed. Check Immediate Window for details.", vbInformation, PROC_NAME
    
CleanExit:
    Exit Sub
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, "Error in Test_Core_Utils", Err.Description, Err.Number
    MsgBox "Error in Test_Core_Utils: " & Err.Description, vbCritical, PROC_NAME
    Resume CleanExit
End Sub


