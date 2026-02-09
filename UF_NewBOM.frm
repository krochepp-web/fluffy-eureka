VERSION 5.00
Begin VB.UserForm UF_NewBOM
   Caption         =   "New BOM"
   ClientHeight    =   2700
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3480
      TabIndex        =   9
      Top             =   2250
      Width           =   960
   End
   Begin VB.CommandButton cmdCreate
      Caption         =   "Create BOM"
      Height          =   330
      Left            =   2340
      TabIndex        =   8
      Top             =   2250
      Width           =   1080
   End
   Begin VB.TextBox txtTADesc
      Height          =   390
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   3120
   End
   Begin VB.TextBox txtTARev
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   1260
      Width           =   1200
   End
   Begin VB.TextBox txtTAPN
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   900
      Width           =   3120
   End
   Begin VB.TextBox txtTAID
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   540
      Width           =   3120
   End
   Begin VB.Label lblTADesc
      Caption         =   "Description"
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   1740
      Width           =   1020
   End
   Begin VB.Label lblTARev
      Caption         =   "TA Rev"
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label lblTAPN
      Caption         =   "TA Part Number"
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label lblTAID
      Caption         =   "TA ID"
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "UF_NewBOM"
Option Explicit

Private mWb As Workbook

Public Sub InitForm(ByVal wb As Workbook)
    Set mWb = wb
    txtTAID.Value = vbNullString
    txtTAPN.Value = vbNullString
    txtTARev.Value = vbNullString
    txtTADesc.Value = vbNullString
End Sub

Private Sub cmdCreate_Click()
    On Error GoTo EH

    M_Data_BOMs_Entry.Create_BOM_For_Assembly_FromInputs _
        Trim$(txtTAID.Value), _
        Trim$(txtTAPN.Value), _
        Trim$(txtTARev.Value), _
        Trim$(txtTADesc.Value)

    Unload Me
    Exit Sub

EH:
    MsgBox "New BOM creation failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New BOM"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
