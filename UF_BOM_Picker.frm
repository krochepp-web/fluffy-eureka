VERSION 5.00
Begin VB.UserForm UF_BOM_Picker
   Caption         =   "Component Picker"
   ClientHeight    =   4200
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   3000
   End
   Begin VB.TextBox txtRev
      Height          =   285
      Left            =   3300
      TabIndex        =   1
      Top             =   300
      Width           =   900
   End
   Begin VB.CheckBox chkActiveOnly
      Caption         =   "Active only"
      Height          =   240
      Left            =   4320
      TabIndex        =   2
      Top             =   330
      Width           =   1200
   End
   Begin VB.TextBox txtMax
      Height          =   285
      Left            =   5700
      TabIndex        =   3
      Top             =   300
      Width           =   900
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "Refresh"
      Height          =   300
      Left            =   5700
      TabIndex        =   4
      Top             =   660
      Width           =   900
   End
   Begin VB.ListBox lstResults
      Height          =   2200
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   6960
   End
   Begin VB.TextBox txtQty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   3300
      Width           =   900
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "Add to BOM"
      Height          =   330
      Left            =   1200
      TabIndex        =   7
      Top             =   3270
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   330
      Left            =   2520
      TabIndex        =   8
      Top             =   3270
      Width           =   900
   End
   Begin VB.Label lblSearch
      Caption         =   "Search (desc/notes/PN)"
      Height          =   240
      Left            =   120
      Top             =   60
      Width           =   3000
   End
   Begin VB.Label lblRev
      Caption         =   "Rev"
      Height          =   240
      Left            =   3300
      Top             =   60
      Width           =   600
   End
   Begin VB.Label lblMax
      Caption         =   "Max"
      Height          =   240
      Left            =   5700
      Top             =   60
      Width           =   600
   End
   Begin VB.Label lblQty
      Caption         =   "QtyPer"
      Height          =   240
      Left            =   120
      Top             =   3060
      Width           =   900
   End
End
Attribute VB_Name = "UF_BOM_Picker"
Option Explicit

Private Sub UserForm_Initialize()
    Me.chkActiveOnly.value = True
    Me.txtMax.Text = "250"
    Me.txtQty.Text = "1"

    Me.lstResults.ColumnCount = 6
    Me.lstResults.ColumnWidths = "0 pt;80 pt;50 pt;240 pt;60 pt;70 pt"
    Me.lstResults.MultiSelect = fmMultiSelectMulti

    RefreshResults
End Sub

Private Sub cmdRefresh_Click()
    RefreshResults
End Sub

Private Sub cmdAdd_Click()
    Dim qtyPer As Double
    Dim i As Long
    Dim added As Long

    qtyPer = ParseQty(Me.txtQty.Text)
    If qtyPer <= 0 Then
        MsgBox "QtyPer must be > 0.", vbExclamation, "Component Picker"
        Exit Sub
    End If

    If Me.lstResults.ListCount = 0 Then Exit Sub

    For i = 0 To Me.lstResults.ListCount - 1
        If Me.lstResults.Selected(i) Then
            Dim pn As String
            Dim rev As String
            pn = CStr(Me.lstResults.List(i, 1))
            rev = CStr(Me.lstResults.List(i, 2))
            If Len(pn) > 0 And Len(rev) > 0 Then
                M_Data_BOMs_Picker.AddComponentToActiveBOM pn, rev, qtyPer
                added = added + 1
            End If
        End If
    Next i

    If added > 0 Then
        MsgBox "Added " & added & " component(s) to the active BOM.", vbInformation, "Component Picker"
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub RefreshResults()
    Dim searchText As String
    Dim revFilter As String
    Dim activeOnly As Boolean
    Dim maxResults As Long
    Dim outArr() As Variant
    Dim outCount As Long
    Dim i As Long

    searchText = Trim$(Me.txtSearch.Text)
    revFilter = Trim$(Me.txtRev.Text)
    activeOnly = CBool(Me.chkActiveOnly.value)
    maxResults = ParseLong(Me.txtMax.Text, 250)

    M_Data_BOMs_Picker.Picker_GetResultsArray searchText, revFilter, activeOnly, maxResults, outArr, outCount

    Me.lstResults.Clear
    If outCount <= 0 Then Exit Sub

    For i = 1 To outCount
        Me.lstResults.AddItem CStr(outArr(i, 1))
        Me.lstResults.List(Me.lstResults.ListCount - 1, 1) = CStr(outArr(i, 2))
        Me.lstResults.List(Me.lstResults.ListCount - 1, 2) = CStr(outArr(i, 3))
        Me.lstResults.List(Me.lstResults.ListCount - 1, 3) = CStr(outArr(i, 4))
        Me.lstResults.List(Me.lstResults.ListCount - 1, 4) = CStr(outArr(i, 5))
        Me.lstResults.List(Me.lstResults.ListCount - 1, 5) = CStr(outArr(i, 7))
    Next i
End Sub

Private Function ParseQty(ByVal s As String) As Double
    If Len(Trim$(s)) = 0 Then
        ParseQty = -1#
    ElseIf Not IsNumeric(s) Then
        ParseQty = -1#
    Else
        ParseQty = CDbl(s)
    End If
End Function

Private Function ParseLong(ByVal s As String, ByVal defaultVal As Long) As Long
    If Len(Trim$(s)) = 0 Then
        ParseLong = defaultVal
    ElseIf Not IsNumeric(s) Then
        ParseLong = defaultVal
    Else
        ParseLong = CLng(s)
    End If
End Function
