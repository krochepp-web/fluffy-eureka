VERSION 5.00
Begin VB.UserForm UF_BOM_Create
   Caption         =   "Create New BOM"
   ClientHeight    =   3000
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTaId
      Height          =   285
      Left            =   1500
      TabIndex        =   0
      Top             =   240
      Width           =   2800
   End
   Begin VB.TextBox txtTaPn
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   720
      Width           =   2800
   End
   Begin VB.TextBox txtTaRev
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      Top             =   1200
      Width           =   2800
   End
   Begin VB.TextBox txtTaDesc
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Top             =   1680
      Width           =   2800
   End
   Begin VB.CommandButton cmdCreate
      Caption         =   "Create BOM"
      Height          =   330
      Left            =   1500
      TabIndex        =   4
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2850
      TabIndex        =   5
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label lblTaId
      Caption         =   "TA ID"
      Height          =   240
      Left            =   240
      Top             =   270
      Width           =   1200
   End
   Begin VB.Label lblTaPn
      Caption         =   "TA Part Number"
      Height          =   240
      Left            =   240
      Top             =   750
      Width           =   1200
   End
   Begin VB.Label lblTaRev
      Caption         =   "TA Revision"
      Height          =   240
      Left            =   240
      Top             =   1230
      Width           =   1200
   End
   Begin VB.Label lblTaDesc
      Caption         =   "TA Description"
      Height          =   240
      Left            =   240
      Top             =   1710
      Width           =   1200
   End
End
Attribute VB_Name = "UF_BOM_Create"
Option Explicit

Public Cancelled As Boolean

Public Property Get TAID() As String
    TAID = Trim$(Me.txtTaId.Text)
End Property

Public Property Get TAPN() As String
    TAPN = Trim$(Me.txtTaPn.Text)
End Property

Public Property Get TARev() As String
    TARev = Trim$(Me.txtTaRev.Text)
End Property

Public Property Get TADesc() As String
    TADesc = Trim$(Me.txtTaDesc.Text)
End Property

Private Sub cmdCreate_Click()
    If Len(Trim$(Me.txtTaId.Text)) = 0 Then
        MsgBox "TA ID is required.", vbExclamation, "Create BOM"
        Me.txtTaId.SetFocus
        Exit Sub
    End If

    If Len(Trim$(Me.txtTaPn.Text)) = 0 Then
        MsgBox "TA Part Number is required.", vbExclamation, "Create BOM"
        Me.txtTaPn.SetFocus
        Exit Sub
    End If

    If Len(Trim$(Me.txtTaRev.Text)) = 0 Then
        MsgBox "TA Revision is required.", vbExclamation, "Create BOM"
        Me.txtTaRev.SetFocus
        Exit Sub
    End If

    If Len(Trim$(Me.txtTaDesc.Text)) = 0 Then
        MsgBox "TA Description is required.", vbExclamation, "Create BOM"
        Me.txtTaDesc.SetFocus
        Exit Sub
    End If

    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Cancelled = True
End Sub
