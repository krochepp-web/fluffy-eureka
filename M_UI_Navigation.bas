Attribute VB_Name = "M_UI_Navigation"
'this bas module has the following UI_GoToSheet"XYZ"() Where XYZ is a Worksheet in this workbook:
'1. Landing Tab
'2. Schema_Check (output, individual workbook schema errors) Tab
'3. SCHEMA (input, with column header definitions) Tab
'4. Core_Tests (output) Tab
'5. Workbook_Schema (output, each Tab, Table, and Column Header in the workbook) Tab
'6. Next...If Not Gate_Ready(True) Then Exit Sub
'7.
'8.
'9.
'10.
'11.
'12.


Sub UI_GoToBOM_TEMPLATE()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("BOM_TEMPLATE").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToUsers()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("Users").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToData_Check()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("Data_Check").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub


Sub UI_GoToSheetLanding()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Landing").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetSchema_Check()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("Schema_Check").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetSCHEMA()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("SCHEMA").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetCore_Tests()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("Core_Tests").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetWorkbook_Schema()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
    Worksheets("Workbook_Schema").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub


Sub UI_GoToSheetAUTO()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("AUTO").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetSuppliers()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Suppliers").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetComps()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Comps").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetRHistory()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("RHistory").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetHelpers()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Helpers").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetDev_ModuleCatalog()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Dev_ModuleCatalog").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheetLockdown_Preview()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Lockdown_Preview").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

Sub UI_GoToSheeDev_ProcedureCatalog()
    ' Replace "YourSheetName" with the actual name of the worksheet tab
    ' For example: Worksheets("Data").Activate
      Worksheets("Dev_ProcedureCatalog").Activate
    
    ' Reset the view so that cell A1 is in the upper-left corner
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Optional: Select cell A1 (uncomment if needed)
    ' Range("A1").Select
End Sub

