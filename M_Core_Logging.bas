Attribute VB_Name = "M_Core_Logging"
Option Explicit
'*******************************************************************************
' Module:      M_Core_Logging
' Procedure:   LogEvent (public entry point)
'
' Purpose:
'   Centralized logging utility for the Tracker workbook. Appends rows to
'   Log.TBL_LOG using the schema defined in Schema 3.4.3. All modules should
'   call LogEvent (or the LogInfo/LogWarn/LogError wrappers) instead of
'   writing to the log table directly.
'
' Inputs (Tabs/Tables/Headers):
'   - Sheet:  SH_LOG
'   - Table:  TBL_LOG
'   - Columns (from M_Core_Constants):
'       COL_LOG_TIMESTAMP
'       COL_LOG_LEVEL
'       COL_LOG_PROC
'       COL_LOG_MESSAGE
'       COL_LOG_DETAILS
'       COL_LOG_ERRNUM
'       COL_LOG_ACTIVITY_ID
'       COL_LOG_REPEAT_COUNT
'       COL_LOG_USER_ID
'       COL_LOG_WORKBOOK
'       COL_LOG_VERSION
'       COL_LOG_OTHER
'
' Outputs / Side effects:
'   - Appends a single row to Log.TBL_LOG for each call to LogEvent.
'   - Does not modify any other tables or sheets.
'
' Preconditions:
'   - Sheet SH_LOG exists.
'   - Table TBL_LOG exists on SH_LOG and has the expected columns.
'   - M_Core_Constants is present and compiled.
'
' Postconditions:
'   - One new log row is added, or (on failure) a Debug.Print message is emitted.
'
' Errors & Guards:
'   - Errors inside LogEvent are trapped locally; logging failures will not
'     throw unhandled errors back into business code.
'   - A simple re-entry guard prevents recursive logging if LogEvent itself fails.
'
' Version:     v0.1.0
' Author:      ChatGPT (assistant)
' Date:        2025-11-28
'
' @spec
'   Purpose: Centralize structured logging to TBL_LOG.
'   Inputs: procName, level, message, details, errNum, activityId
'   Outputs: One appended row in TBL_LOG per call.
'   Preconditions: SH_LOG/TBL_LOG exist and match TBL_SCHEMA.
'   Postconditions: Log row captured or graceful fallback to Debug.Print.
'   Errors: Logging failures are swallowed with Debug.Print; no recursion.
'   Version: v0.1.0
'   Author: ChatGPT
'   Date: 2025-11-28
'*******************************************************************************

' Re-entry guard to avoid recursive logging if something inside LogEvent fails
Private mIsLogging As Boolean
Private Const LOG_DIAGNOSTICS As Boolean = True

'===============================================================================
' Public API
'===============================================================================

Public Sub LogEvent( _
    ByVal procName As String, _
    ByVal level As String, _
    ByVal message As String, _
    Optional ByVal details As String = "", _
    Optional ByVal errNum As Long = 0, _
    Optional ByVal activityId As String = "")

    Const PROC_NAME As String = "LogEvent"
    
    Dim wb As Workbook
    Dim wsLog As Worksheet
    Dim loLog As ListObject
    Dim newRow As ListRow
    
    Dim nowStamp As Date
    Dim currentUser As String
    Dim logActivityId As String
    Dim safeLevel As String
    
    On Error GoTo EH
    
    ' Prevent recursive logging
    If mIsLogging Then
        GoTo CleanExit
    End If
    mIsLogging = True
    
    Set wb = ThisWorkbook
    nowStamp = Now
    
    ' Simple user resolution (can be enhanced later to map to TBL_USERS)
    currentUser = Environ$("Username")
    If Len(Trim$(currentUser)) = 0 Then
        currentUser = "UNKNOWN"
    End If
    
    ' Normalize level to known tokens if possible
    safeLevel = UCase$(Trim$(level))
    Select Case safeLevel
        Case LOG_LEVEL_INFO, LOG_LEVEL_WARN, LOG_LEVEL_ERROR
            ' ok
        Case Else
            safeLevel = LOG_LEVEL_INFO
    End Select
    
    ' Activity ID: if none supplied, generate a simple one
    If Len(Trim$(activityId)) = 0 Then
        logActivityId = GenerateActivityId(procName)
    Else
        logActivityId = activityId
    End If
    
    ' Get the log table
    Set wsLog = wb.Worksheets(SH_LOG)
    Set loLog = wsLog.ListObjects(TBL_LOG)
    
    ' Append a new row
    Set newRow = loLog.ListRows.Add
    
    With newRow.Range
        ' Timestamp
        SafeSetColumnValue loLog, newRow, COL_LOG_TIMESTAMP, nowStamp
        ' Level
        SafeSetColumnValue loLog, newRow, COL_LOG_LEVEL, safeLevel
        ' Proc
        SafeSetColumnValue loLog, newRow, COL_LOG_PROC, procName
        ' Message
        SafeSetColumnValue loLog, newRow, COL_LOG_MESSAGE, message
        ' Details
        SafeSetColumnValue loLog, newRow, COL_LOG_DETAILS, details
        ' ErrNum
        If errNum <> 0 Then
            SafeSetColumnValue loLog, newRow, COL_LOG_ERRNUM, errNum
        Else
            SafeSetColumnValue loLog, newRow, COL_LOG_ERRNUM, vbNullString
        End If
        ' ActivityId
        SafeSetColumnValue loLog, newRow, COL_LOG_ACTIVITY_ID, logActivityId
        ' RepeatCount (for now always 1; can be aggregated later)
        SafeSetColumnValue loLog, newRow, COL_LOG_REPEAT_COUNT, 1
        ' UserID
        SafeSetColumnValue loLog, newRow, COL_LOG_USER_ID, currentUser
        ' Workbook
        SafeSetColumnValue loLog, newRow, COL_LOG_WORKBOOK, wb.Name
        ' Version
        SafeSetColumnValue loLog, newRow, COL_LOG_VERSION, APP_VERSION
        ' Other (spare)
        SafeSetColumnValue loLog, newRow, COL_LOG_OTHER, vbNullString
    End With

    SortLogTableByTimestamp loLog
    
CleanExit:
    mIsLogging = False
    Exit Sub
    
EH:
    ' If logging itself fails, we do NOT raise further errors. Just debug-print.
    Debug.Print "Logging error in " & PROC_NAME & ": " & Err.Number & " - " & Err.Description
    If LOG_DIAGNOSTICS Then
        MsgBox _
            "Logging failed inside " & PROC_NAME & "." & vbCrLf & _
            "Err " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
            "Check these items:" & vbCrLf & _
            "1) Sheet exists: " & SH_LOG & vbCrLf & _
            "2) Table exists: " & TBL_LOG & vbCrLf & _
            "3) Sheet/Table not protected or locked" & vbCrLf & _
            "4) Log table columns match constants in M_Core_Constants", _
            vbExclamation, "Log Diagnostics"
    End If
    MsgBox "Logging error." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Logging"
    Resume CleanExit
End Sub

'-------------------------------------------------------------------------------
' Convenience wrappers
'-------------------------------------------------------------------------------

Public Sub LogInfo(ByVal procName As String, ByVal message As String, Optional ByVal details As String = "")
    LogEvent procName, LOG_LEVEL_INFO, message, details, 0
End Sub

Public Sub LogWarn(ByVal procName As String, ByVal message As String, Optional ByVal details As String = "")
    LogEvent procName, LOG_LEVEL_WARN, message, details, 0
End Sub

Public Sub LogError( _
    ByVal procName As String, _
    ByVal message As String, _
    Optional ByVal details As String = "", _
    Optional ByVal errNum As Long = 0)
    
    LogEvent procName, LOG_LEVEL_ERROR, message, details, errNum
End Sub

'===============================================================================
' Private helpers
'===============================================================================

' Generate a simple ActivityId string combining procedure, date, and time.
' This is not a GUID, but is sufficient to group related log entries.
Private Function GenerateActivityId(ByVal procName As String) As String
    Dim ts As String
    ts = Format$(Now, "yyyymmdd_hhnnss")
    GenerateActivityId = procName & "_" & ts
End Function

' Safely set a column value by column header name, if the column exists.
' If the column is missing (schema drift), this fails silently for robustness.
Private Sub SafeSetColumnValue( _
    ByVal lo As ListObject, _
    ByVal row As ListRow, _
    ByVal columnName As String, _
    ByVal value As Variant)
    
    Dim lc As ListColumn
    On Error GoTo CleanExit
    
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            row.Range.Cells(1, lc.Index).value = value
            Exit For
        End If
    Next lc
    
CleanExit:
    On Error GoTo 0
End Sub

' Keep newest log entries at the top by sorting Timestamp descending.
Private Sub SortLogTableByTimestamp(ByVal lo As ListObject)
    Dim tsColumn As ListColumn
    
    On Error GoTo CleanExit
    
    If lo Is Nothing Then
        GoTo CleanExit
    End If
    
    Set tsColumn = lo.ListColumns(COL_LOG_TIMESTAMP)
    If tsColumn Is Nothing Then
        GoTo CleanExit
    End If
    
    If lo.ListRows.Count <= 1 Then
        GoTo CleanExit
    End If
    
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tsColumn.Range, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With
    
CleanExit:
    On Error GoTo 0
End Sub


