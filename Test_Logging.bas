Attribute VB_Name = "Test_Logging"
Public Sub Test_Logging()
    Const PROC_NAME As String = "Test_Logging"
    
    LogInfo PROC_NAME, "This is an INFO test.", "Detail: nothing special."
    LogWarn PROC_NAME, "This is a WARN test.", "Detail: check thresholds."
    LogError PROC_NAME, "This is an ERROR test.", "Detail: simulated failure.", 1234
End Sub

