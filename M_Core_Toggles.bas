Attribute VB_Name = "M_Core_Toggles"
Option Explicit
'*******************************************************************************
' Module:      M_Core_Toggles
'
' Purpose:
'   Centralized control of Excel application state (ScreenUpdating, EnableEvents,
'   Calculation) for Tracker macros. Provides a simple, safe pattern for
'   entering/exiting a "critical section" where performance and stability are
'   prioritized.
'
'   Callers should:
'       1) Call App_EnterCriticalSection at the start of a public macro.
'       2) Call App_ExitCriticalSection in the CleanExit/EH path.
'
'   Nested calls are supported: state is only captured/restored for the
'   outermost caller.
'
' Inputs (Tabs/Tables/Headers):
'   - None. Operates only on Application-level properties.
'
' Outputs / Side effects:
'   - Temporarily sets:
'       Application.ScreenUpdating = False
'       Application.EnableEvents   = False
'       Application.Calculation    = xlCalculationManual
'   - Restores original values when the outermost critical section exits.
'
' Preconditions / Postconditions:
'   - Callers must ensure every App_EnterCriticalSection has a matching
'     App_ExitCriticalSection, even when errors occur (use standard error
'     handler pattern).
'   - After App_ExitCriticalSection unwinds the last nested level, application
'     state is restored to what it was before the first Enter.
'
' Errors & Guards:
'   - Errors are logged via M_Core_Logging.LogEvent and swallowed; the module
'     makes a best effort to avoid leaving Excel in a bad state.
'
' Version:     v0.1.0
' Author:      ChatGPT (assistant)
' Date:        2025-11-29
'
' @spec
'   Purpose: Provide safe, centralized toggling of Excel global state for
'            performance and stability during macro execution.
'   Inputs: activityId (optional identifier for logging correlation).
'   Outputs: None; modifies Application global state and maintains internal
'            depth tracking.
'   Preconditions: M_Core_Logging and M_Core_Constants compiled and available.
'   Postconditions: Application state restored when outermost caller exits.
'   Errors: Logged via LogEvent; module attempts to restore state on failure.
'   Version: v0.1.0
'   Author: ChatGPT
'   Date: 2025-11-29
'*******************************************************************************

'===============================================================================
' Module-level state
'===============================================================================

Private m_prevScreenUpdating As Boolean
Private m_prevEnableEvents   As Boolean
Private m_prevCalcMode       As XlCalculation
Private m_stateDepth         As Long
Private m_stateInitialized   As Boolean

'===============================================================================
' Public API - Critical section control
'===============================================================================

Public Sub App_EnterCriticalSection( _
    Optional ByVal activityId As String = "")
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Enter a critical section for macro execution. On the first (outermost)
    '   call, capture the current Application state and apply a performant,
    '   safe configuration:
    '       ScreenUpdating = False
    '       EnableEvents   = False
    '       Calculation    = xlCalculationManual
    '
    '   Nested calls only increment depth; state is not re-applied.
    '
    ' Errors:
    '   Any error is logged via LogEvent and the procedure exits without
    '   raising further errors.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "App_EnterCriticalSection"
    
    On Error GoTo EH
    
    ' If no activityId provided, generate one for potential logging correlation.
    If Len(activityId) = 0 Then
        On Error Resume Next
        activityId = Utils_GenerateActivityId(PROC_NAME)
        On Error GoTo EH
    End If
    
    ' Only capture and modify Application state on the outermost entry.
    If m_stateDepth = 0 Then
        m_prevScreenUpdating = Application.ScreenUpdating
        m_prevEnableEvents = Application.EnableEvents
        m_prevCalcMode = Application.Calculation
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        m_stateInitialized = True
    End If
    
    m_stateDepth = m_stateDepth + 1
    
CleanExit:
    Exit Sub
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error entering critical section", Err.Description, Err.Number, activityId
    MsgBox "Error entering critical section." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Application State"
    ' Do not raise further; better to leave state unchanged than crash.
    Resume CleanExit
End Sub

Public Sub App_ExitCriticalSection( _
    Optional ByVal activityId As String = "")
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Exit a previously entered critical section. Decrements the internal
    '   depth counter; when it reaches zero, restores the original Application
    '   state captured on the outermost entry.
    '
    ' Errors:
    '   Any error is logged via LogEvent and the procedure exits without
    '   raising further errors.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "App_ExitCriticalSection"
    
    On Error GoTo EH
    
    If m_stateDepth > 0 Then
        m_stateDepth = m_stateDepth - 1
    End If
    
    If m_stateDepth = 0 And m_stateInitialized Then
        Application.ScreenUpdating = m_prevScreenUpdating
        Application.EnableEvents = m_prevEnableEvents
        Application.Calculation = m_prevCalcMode
        
        m_stateInitialized = False
    End If
    
CleanExit:
    Exit Sub
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, _
             "Error exiting critical section", Err.Description, Err.Number, activityId
    MsgBox "Error exiting critical section." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Application State"
    ' Best effort to restore state even on error.
    Resume CleanExit
End Sub

Public Function App_IsInCriticalSection() As Boolean
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Convenience function to check whether the application is currently in
    '   a critical section (i.e., App_EnterCriticalSection has been called
    '   more times than App_ExitCriticalSection).
    '-------------------------------------------------------------------------
    App_IsInCriticalSection = (m_stateDepth > 0)
End Function

'===============================================================================
' Test harness
'===============================================================================

Public Sub Test_Core_Toggles()
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Smoke test for M_Core_Toggles. Demonstrates that:
    '       - Entering a critical section turns off ScreenUpdating/Events and
    '         sets calculation to Manual.
    '       - Exiting restores the original settings.
    '
    '   This is a non-destructive test and is safe to run at any time, but
    '   should not be wired to normal user buttons.
    '-------------------------------------------------------------------------
    Const PROC_NAME As String = "Test_Core_Toggles"
    
    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalc As XlCalculation
    Dim msg As String
    
    On Error GoTo EH
    
    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    originalCalc = Application.Calculation
    
    msg = "Before Enter:" & vbCrLf & _
          "  ScreenUpdating = " & CStr(originalScreenUpdating) & vbCrLf & _
          "  EnableEvents   = " & CStr(originalEnableEvents) & vbCrLf & _
          "  Calculation    = " & CalcModeToString(originalCalc)
    Debug.Print msg
    
    App_EnterCriticalSection (Utils_GenerateActivityId(PROC_NAME))
    
    msg = "Inside critical section:" & vbCrLf & _
          "  ScreenUpdating = " & CStr(Application.ScreenUpdating) & vbCrLf & _
          "  EnableEvents   = " & CStr(Application.EnableEvents) & vbCrLf & _
          "  Calculation    = " & CalcModeToString(Application.Calculation)
    Debug.Print msg
    
    App_ExitCriticalSection
    
    msg = "After Exit:" & vbCrLf & _
          "  ScreenUpdating = " & CStr(Application.ScreenUpdating) & vbCrLf & _
          "  EnableEvents   = " & CStr(Application.EnableEvents) & vbCrLf & _
          "  Calculation    = " & CalcModeToString(Application.Calculation)
    Debug.Print msg
    
    MsgBox "Test_Core_Toggles completed. Check Immediate Window for state snapshots.", _
           vbInformation, PROC_NAME
    
CleanExit:
    Exit Sub
    
EH:
    On Error Resume Next
    LogEvent PROC_NAME, LOG_LEVEL_ERROR, "Error in Test_Core_Toggles", Err.Description, Err.Number
    MsgBox "Error in Test_Core_Toggles." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, PROC_NAME
    Resume CleanExit
End Sub

'===============================================================================
' Private helpers
'===============================================================================

Private Function CalcModeToString(ByVal mode As XlCalculation) As String
    '-------------------------------------------------------------------------
    ' Purpose:
    '   Helper to convert Application.Calculation mode to a readable string.
    '-------------------------------------------------------------------------
    Select Case mode
        Case xlCalculationAutomatic
            CalcModeToString = "Automatic"
        Case xlCalculationManual
            CalcModeToString = "Manual"
        Case xlCalculationSemiautomatic
            CalcModeToString = "Semiautomatic"
        Case Else
            CalcModeToString = CStr(mode)
    End Select
End Function


