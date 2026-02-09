VERSION 5.00
Begin VB.UserForm UF_ComponentPicker
   Caption         =   "Component Picker"
   ClientHeight    =   4500
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   330
      Left            =   6360
      TabIndex        =   13
      Top             =   4080
      Width           =   900
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "Add Selected"
      Height          =   330
      Left            =   5160
      TabIndex        =   12
      Top             =   4080
      Width           =   1140
   End
   Begin VB.TextBox txtQtyPer
      Height          =   300
      Left            =   4200
      TabIndex        =   10
      Top             =   410
      Width           =   840
   End
   Begin VB.CheckBox chkActiveOnly
      Caption         =   "Active Only"
      Height          =   240
      Left            =   5400
      TabIndex        =   11
      Top             =   420
      Width           =   1200
   End
   Begin VB.TextBox txtMaxResults
      Height          =   300
      Left            =   6600
      TabIndex        =   9
      Top             =   420
      Width           =   720
   End
   Begin VB.TextBox txtRev
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Top             =   420
      Width           =   720
   End
   Begin VB.TextBox txtSearch
      Height          =   300
      Left            =   720
      TabIndex        =   5
      Top             =   420
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "Search"
      Height          =   330
      Left            =   3600
      TabIndex        =   8
      Top             =   400
      Width           =   540
   End
   Begin VB.ListBox lstResults
      Height          =   3000
      Left            =   180
      TabIndex        =   14
      Top             =   900
      Width           =   7200
   End
   Begin VB.Label lblMax
      Caption         =   "Max"
      Height          =   240
      Left            =   6600
      TabIndex        =   4
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblQty
      Caption         =   "QtyPer"
      Height          =   240
      Left            =   4200
      TabIndex        =   3
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblRev
      Caption         =   "Rev"
      Height          =   240
      Left            =   2760
      TabIndex        =   2
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblSearch
      Caption         =   "Search"
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblHelp
      Caption         =   "Select one or more rows, then Add Selected."
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   3780
      Width           =   3900
   End
End
Attribute VB_Name = "UF_ComponentPicker"
Option Explicit

Private mWb As Workbook

Public Sub InitForm(ByVal wb As Workbook)
    Set mWb = wb
    txtSearch.Value = vbNullString
    txtRev.Value = vbNullString
    txtQtyPer.Value = "1"
    txtMaxResults.Value = "250"
    chkActiveOnly.Value = True
    ConfigureList
    RefreshResults
End Sub

Private Sub ConfigureList()
    lstResults.ColumnCount = 7
    lstResults.ColumnWidths = "0 pt;80 pt;50 pt;220 pt;50 pt;220 pt;60 pt"
    lstResults.MultiSelect = fmMultiSelectExtended
End Sub

Private Sub cmdRefresh_Click()
    RefreshResults
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then RefreshResults
End Sub

Private Sub cmdAdd_Click()
    Dim qtyPer As Double
    qtyPer = ParseQtyPer()
    If qtyPer <= 0 Then Exit Sub

    If lstResults.ListCount = 0 Then
        MsgBox "No results to add.", vbInformation, "Component Picker"
        Exit Sub
    End If

    Dim loBom As ListObject
    Set loBom = M_Data_BOMs_Picker.GetActiveBomTable_Picker()

    Dim i As Long
    Dim added As Long
    For i = 0 To lstResults.ListCount - 1
        If lstResults.Selected(i) Then
            M_Data_BOMs_Picker.Picker_AddComponentToActiveBOM _
                loBom, _
                CStr(lstResults.List(i, 0)), _
                CStr(lstResults.List(i, 1)), _
                CStr(lstResults.List(i, 2)), _
                CStr(lstResults.List(i, 3)), _
                CStr(lstResults.List(i, 4)), _
                qtyPer, _
                CStr(lstResults.List(i, 5)), _
                CStr(lstResults.List(i, 6))
            added = added + 1
        End If
    Next i

    MsgBox "Components added: " & added, vbInformation, "Component Picker"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshResults()
    On Error GoTo EH

    Dim outCount As Long
    Dim results As Variant
    results = M_Data_BOMs_Picker.Picker_GetResults( _
        mWb, _
        LCase$(Trim$(txtSearch.Value)), _
        Trim$(txtRev.Value), _
        CBool(chkActiveOnly.Value), _
        CLng(ParseLong(txtMaxResults.Value, 250)), _
        outCount)

    lstResults.Clear
    If outCount = 0 Then Exit Sub
    lstResults.List = results
    Exit Sub

EH:
    MsgBox "Search failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Component Picker"
End Sub

Private Function ParseQtyPer() As Double
    Dim valText As String
    valText = Trim$(txtQtyPer.Value)
    If Len(valText) = 0 Or Not IsNumeric(valText) Then
        MsgBox "QtyPer must be a number > 0.", vbExclamation, "Component Picker"
        ParseQtyPer = -1
        Exit Function
    End If
    ParseQtyPer = CDbl(valText)
    If ParseQtyPer <= 0 Then
        MsgBox "QtyPer must be > 0.", vbExclamation, "Component Picker"
        ParseQtyPer = -1
    End If
End Function

Private Function ParseLong(ByVal v As String, ByVal defaultVal As Long) As Long
    If Len(Trim$(v)) = 0 Then
        ParseLong = defaultVal
    ElseIf IsNumeric(v) Then
        ParseLong = CLng(v)
    Else
        ParseLong = defaultVal
    End If
End Function
